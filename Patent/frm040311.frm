VERSION 5.00
Begin VB.Form frm040311 
   BorderStyle     =   1  '單線固定
   Caption         =   "核准(駁)簿"
   ClientHeight    =   3555
   ClientLeft      =   3225
   ClientTop       =   1425
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4155
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1600
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1776
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2880
      Width           =   330
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3240
      TabIndex        =   12
      Top             =   50
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2430
      TabIndex        =   11
      Top             =   50
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1128
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2580
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2280
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1020
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2280
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   900
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1980
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1404
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1380
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2244
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1080
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1032
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1080
      Width           =   1020
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1068
      MaxLength       =   1
      TabIndex        =   1
      Top             =   780
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1044
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "是否含新型申請：            (Y : 是)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   25
      Top             =   1740
      Width           =   3180
   End
   Begin VB.Label Label1 
      Caption         =   "是否依承辦人跳頁：            (Y : 是)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   105
      TabIndex        =   24
      Top             =   2940
      Width           =   3180
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   1950
      TabIndex        =   23
      Top             =   2130
      Width           =   1665
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   2412
      TabIndex        =   22
      Top             =   1440
      Width           =   1668
   End
   Begin VB.Line Line2 
      X1              =   2100
      X2              =   2220
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Line Line1 
      X1              =   2055
      X2              =   2175
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦人2.准駁日)"
      Height          =   180
      Index           =   8
      Left            =   1470
      TabIndex        =   21
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "(1.准 2.駁)"
      Height          =   180
      Index           =   7
      Left            =   1425
      TabIndex        =   20
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      Height          =   180
      Index           =   6
      Left            =   105
      TabIndex        =   19
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   5
      Left            =   105
      TabIndex        =   18
      Top             =   2340
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   4
      Left            =   105
      TabIndex        =   17
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質代碼："
      Height          =   180
      Index           =   3
      Left            =   108
      TabIndex        =   16
      Top             =   1440
      Width           =   1416
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日期："
      Height          =   180
      Index           =   2
      Left            =   108
      TabIndex        =   15
      Top             =   1140
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "准/駁代碼："
      Height          =   180
      Index           =   1
      Left            =   108
      TabIndex        =   14
      Top             =   840
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   13
      Top             =   540
      Width           =   948
   End
End
Attribute VB_Name = "frm040311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 10) As String, strTemp1 As Variant, strTemp2 As Variant, StrTemp8(0 To 1) As String, k As Integer
Dim PLeft(0 To 7) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 4) As String, StrTemp5(0 To 4) As String, StrTemp6(0 To 4) As String
'Add By Cheng 2002/09/11
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
'Add By Cheng 2003/02/14
Dim strTmp3 As String
Dim strPrintTitle As Boolean  '2011/10/3 add by sonia 已印過最上層表頭

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0 '確定
         'Add By Cheng 2002/09/11
         blnClkSure = False
        If Len(txt1(0)) = 0 Then
           s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
           txt1(0).SetFocus
           Exit Sub
        Else
            If Len(txt1(1)) = 0 Then
               s = MsgBox("准駁代碼不可空白!!", , "USER 輸入錯誤")
               txt1(1).SetFocus
               Exit Sub
            Else
               'Add By Cheng 2002/03/19
               If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
                  Me.txt1(2).SetFocus
                  txt1_GotFocus 2
                  Exit Sub
               End If
               If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
               'Add By Cheng 2002/09/11
               If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
                  If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
                     MsgBox "准駁日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(2).SetFocus
                     txt1_GotFocus 2
                     Exit Sub
                  End If
               End If
                            
               If Len(txt1(3)) = 0 Then
                  s = MsgBox("准駁日期區間不可空白!!", , "USER 輸入錯誤")
                   txt1(2).SetFocus
                   txt1_GotFocus (2)
                  Exit Sub
               Else
                  If txt1(4) <> "" Then
                     'Add By Cheng 2002/09/11
                     lbl1(0) = GetPrjState6HM("P", txt1(4))
                     If lbl1(0) = "" Then
                        MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
                        Me.txt1(4).SetFocus
                        txt1_GotFocus 4
                        Exit Sub
                     End If
                  End If
                  If txt1(5) <> "" Then
                     'edit by nickc 2007/02/02 不用 dll 了
                     'If objPublicData.GetStaff(txt1(5), strExc(0)) Then
                     If ClsPDGetStaffN(txt1(5), strExc(0)) Then
                        lbl1(1) = strExc(0)
                     Else
                        lbl1(1) = ""
                        Me.txt1(5).SetFocus
                        txt1_GotFocus 5
                        Exit Sub
                     End If
                  End If
                  If Me.txt1(6).Text <> "" And Me.txt1(7).Text <> "" Then
                     If Me.txt1(6).Text > Me.txt1(7).Text Then
                        MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(6).SetFocus
                        txt1_GotFocus 6
                        Exit Sub
                     End If
                  End If
                   
                   If Len(txt1(8)) = 0 Then
                       s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                       txt1(8).SetFocus
                       Exit Sub
                   Else
                       Screen.MousePointer = vbHourglass
                       Me.Enabled = False
                       ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
                       Process
                       Me.Enabled = True
                       Screen.MousePointer = vbDefault
                   End If
                End If
            End If
        End If
      Case 1
           Unload Me
      Case Else
   End Select
End Sub

Sub Process()
   strSql = "DELETE FROM R040311 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute strSql
   strSQL1 = ""
   strSQL2 = ""
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/2
   End If
   'Modify By Cheng 2003/02/06
   '改抓進度檔的實際結果
   'strSQL1 = strSQL1 + " AND PA16='" & txt1(1) & "' "
   strSQL1 = strSQL1 + " AND CP24='" & txt1(1) & "' "
   If Val(txt1(1)) = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.准" 'Add By Sindy 2010/12/2
   Else
      pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.駁" 'Add By Sindy 2010/12/2
   End If
   If Len(Trim(txt1(2))) <> 0 Then
       'Modify By Cheng 2003/02/06
       '改抓進度檔的准駁日
   '   strSQL1 = strSQL1 + " AND PA20>=" & Val(ChangeTStringToWString(txt1(2))) & " "
       strSQL1 = strSQL1 + " AND CP25>=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   If Len(Trim(txt1(3))) <> 0 Then
       'Modify By Cheng 2003/02/06
       '改抓進度檔的准駁日
   '   strSQL1 = strSQL1 & " AND PA20<=" & Val(ChangeTStringToWString(txt1(3))) & " "
       strSQL1 = strSQL1 & " AND CP25<=" & Val(ChangeTStringToWString(txt1(3))) & " "
   End If
   If Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/2
   End If
   If Len(txt1(4)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP10='" & txt1(4) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & lbl1(0) 'Add By Sindy 2010/12/2
   End If
   If Len(txt1(5)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP14='" & txt1(5) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(1) 'Add By Sindy 2010/12/2
   End If
   'Add By Cheng 2002/12/12
   '除非有指定承辦人, 否則承辦人為陳玲玲及莊敏惠的資料不印
   If Me.txt1(5).Text = "" Then
       'Modified by Morgan 2013/10/23 考慮程序新人
       'strSQL1 = strSQL1 + " AND CP14<>'81002' And CP14<>'73017' "
       strSQL1 = strSQL1 + " AND NVL(S1.ST05,' ')<>'75' "
       'end 2013/10/23
   End If
   If Len(txt1(6)) <> 0 Then
       strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(6) & "' "
   End If
   If Len(txt1(7)) <> 0 Then
       strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(7) & "' "
   End If
   If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/2
   End If
   '2011/4/18 add by sonia
   If txt1(10) = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Label1(10) & ":" & txt1(10)
   Else
      strSQL1 = strSQL1 + " AND CP10<>'102' "
   End If
   '2011/4/18 end
   
   'Add By Cheng 2003/05/13
   '只抓計案件數的資料
   strSQL1 = strSQL1 & " And CP26 Is Null "
   CheckOC
   'Modify By Cheng 2002/12/12
   'strSQL = "SELECT S1.ST02," & SQLDate("PA20") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04,NVL(PA05,NVL(PA06,PA07)),PA11,decode(pa09,'000',cpm03,cpm04),decode(CP23,'1','准','2','駁',''),S2.ST02,'" & strUserNum & "' FROM PATENT,CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND cp14=S1.ST01(+) AND cp13=S2.ST01(+) " & strSQL1
   'Modify By Cheng 2003/02/06
   'strSQL = "SELECT S1.ST02," & SQLDate("PA20") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04,NVL(PA05,NVL(PA06,PA07)),PA11,decode(pa09,'000',cpm03,cpm04),decode(CP23,'1','准','2','駁',''),S2.ST02,'" & strUserNum & "' FROM PATENT,CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND cp14=S1.ST01(+) AND cp13=S2.ST01(+) " & strSQL1
   'Modify By Cheng 2003/02/14
   '承辦人員工名稱改成承辦人員工代號
   'strSQL = "SELECT S1.ST02," & SQLDate("CP25") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04,NVL(PA05,NVL(PA06,PA07)),PA11,decode(pa09,'000',cpm03,cpm04),decode(CP23,'1','准','2','駁',''),S2.ST02,'" & strUserNum & "' FROM PATENT,CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1
   'Modify By Cheng 2003/05/13
   'strSQL = "SELECT CP14," & SQLDate("CP25") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04,NVL(PA05,NVL(PA06,PA07)),PA11,decode(pa09,'000',cpm03,cpm04),decode(CP23,'1','准','2','駁',''),S2.ST02,'" & strUserNum & "' FROM PATENT,CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1
   strSql = "SELECT CP14," & SQLDate("CP25") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04,NVL(PA05,NVL(PA06,PA07)),PA11,decode(pa09,'000',cpm03,cpm04),decode(CP23,'1','准','2','駁',''),S2.ST02,'" & strUserNum & "', CP01, CP10 FROM PATENT,CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1
   cnnConnection.Execute "Insert Into R040311 " & strSql
   strSql = "Select * From R040311 Where ID='" & strUserNum & "' "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/2
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      ShowNoData
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
CheckOC
   If Val(txt1(8)) = 1 Then
       PrintData1      '承辦人
   Else
       PrintData2      '准駁日
   End If
   ShowPrintOk
   Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 7200
   Printer.CurrentY = iPrint
   If Val(txt1(1)) = 1 Then
       Printer.Print "核 准 簿"
   Else
       Printer.Print "核 駁 簿"
   End If
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 6400
   Printer.CurrentY = iPrint
   Printer.Print "准駁日期：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
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
   strPrintTitle = True '2011/10/3 ADD BY SONIA
   Debug.Print str(Page) & "  " & iPrint
End Sub

Sub PrintTitle1()
   If iPrint >= 8600 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
   End If
   GetPleft1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "准駁日"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "預估准駁"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
   End If
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
   End If
   strPrintTitle = False '2011/10/3 ADD BY SONIA
End Sub

Sub PrintTitle2()
   If iPrint >= 8600 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
   End If
   GetPleft2
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "准駁日"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "預估准駁"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle2
   End If
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle2
   End If
   strPrintTitle = False '2011/10/3 ADD BY SONIA
End Sub

Sub PrintData1()
'Add By Cheng 2002/12/12
Dim blnFirstRec As Boolean '第一筆資料

   '2011/5/12 modify by sonia 加入所別排序
   'strSql = "SELECT * FROM R040311 WHERE ID='" & strUserNum & "' ORDER BY R027001,R027002,R027003 "
   strSql = "SELECT R040311.*,ST06 FROM R040311,STAFF WHERE ID='" & strUserNum & "' AND R027001=ST01(+) ORDER BY ST06,R027001,R027002,R027003 "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   strTemp3(0) = " ": strTmp3 = " "
   strTemp3(1) = " "
   strTemp3(2) = " "
   strTemp3(3) = " "  '2011/5/12 ADD BY SONIA
   Page = 1
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           'Add By Cheng 2002/12/12
           '目前資料指向第一筆
           blnFirstRec = True
           PrintTitle
           'PrintTitle1
           'strTemp3(0) = CheckStr(.Fields(0))
           'strTemp3(1) = CheckStr(.Fields(1))
           'strTemp3(2) = CheckStr(.Fields(2))
           Do While .EOF = False
               For i = 0 To 7
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               '2011/5/12 ADD BY SONIA
               strTemp(8) = CheckStr(.Fields("ST06"))
               '2011/5/12 END
               strTmp3 = "" & strTemp(0)
               'Modify By Cheng 2003/02/14
               '將員工代號轉成員工姓名
   '            strTemp(0) = StrToStr(strTemp(0), 4)
               strTemp(0) = StrToStr(GetStaffName(strTemp(0), True), 4)
               strTemp(3) = StrToStr(strTemp(3), 16)
               strTemp(5) = StrToStr(strTemp(5), 4)
               strTemp(7) = StrToStr(strTemp(7), 4)
               '若承辦人不同
               If strTemp3(0) <> strTemp(0) Then
                   '非第一筆資料
                   If blnFirstRec = False Then
                       'Add By Cheng 2002/12/12
                       '2011/5/12 ADD BY SONIA 先依別所不同跳頁再考慮是否依承辦人跳頁
                       'If Me.Txt1(9).Text = "Y" And blnFirstRec = False Then
                       If strPrintTitle = False Then   '2011/10/3 ADD BY SONIA
                           If strTemp3(3) <> strTemp(8) Then
                               Page = Page + 1
                               Printer.NewPage
                               PrintTitle
                           ElseIf Me.txt1(9).Text = "Y" And blnFirstRec = False Then
                               Page = Page + 1
                               Printer.NewPage
                               PrintTitle
                           End If
                       End If                           '2011/10/3 ADD BY SONIA
                       strTemp3(0) = strTemp(0)
                       strTemp3(1) = strTemp(1)
                       strTemp3(2) = strTemp(2)
                       strTemp3(3) = strTemp(8)  '2011/5/12 add by sonia
                       'PrintTotil1
                       PrintTitle1
                   '第一筆資料
                   Else
                       strTemp3(0) = strTemp(0)
                       strTemp3(1) = strTemp(1)
                       strTemp3(2) = strTemp(2)
                       strTemp3(3) = strTemp(8)  '2011/5/12 add by sonia
                       PrintTitle1
                   End If
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
               If iPrint > 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
                   PrintTitle1
               End If
               PrintDatil1
               .MoveNext
               'Add By Cheng 2002/12/12
               '目前資料非指向第一筆
               blnFirstRec = False
               '若還有資料
               If .EOF = False Then
                   '若承辦人不同
                   'Modify By Cheng 2003/02/14
   '                If strTemp3(0) <> CheckStr(.Fields(0)) Then
                   If strTemp3(0) <> GetStaffName(CheckStr(.Fields(0)), True) Then
                       PrintTotil1
                   End If
               End If
           Loop
       End With
   End If
   PrintTotil1
   Printer.EndDoc
End Sub

Sub PrintTotil1()
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
   End If
   'Modify By Cheng 2003/02/14
   'strSQL = "SELECT R027006,COUNT(R027006) FROM R040311 WHERE R027001='" & strTemp3(0) & "' AND ID='" & strUserNum & "' GROUP BY R027006 "
   'Modify By Cheng 2003/05/13
   'strSQL = "SELECT R027006,COUNT(R027006) FROM R040311 WHERE R027001='" & strTmp3 & "' AND ID='" & strUserNum & "' GROUP BY R027006 "
   strSql = "SELECT R027010,COUNT(R027010), Decode(CPM03,'（無）',CPM04,CPM03) FROM R040311, CasePropertyMap WHERE R027009=CPM01(+) And R027010=CPM02(+) And R027001='" & strTmp3 & "' AND ID='" & strUserNum & "' GROUP BY R027010, Decode(CPM03,'（無）',CPM04,CPM03) Order By 1 "
   CheckOC2
   adoRecordset1.CursorLocation = adUseClient
   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
       With adoRecordset1
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 4
                   'Modify By Cheng 2003/05/13
   '                StrTemp5(i) = CheckStr(.Fields(0))
                   StrTemp5(i) = "" & .Fields(2).Value
                   StrTemp6(i) = CheckStr(.Fields(1))
                   Printer.CurrentX = 500 + (i * 2500)
                   Printer.CurrentY = iPrint
                   Printer.Print StrToStr(StrTemp5(i), 6)
                   Printer.CurrentX = 2500 + (i * 2500) - Printer.TextWidth(StrTemp6(i))
                   Printer.CurrentY = iPrint
                   Printer.Print StrTemp6(i)
                   .MoveNext
                   If .EOF = True Then
                       Exit For
                   End If
               Next i
               iPrint = iPrint + 300
               If iPrint >= 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
                   If .EOF = False Then
                       PrintTitle1
                   End If
               End If
               Loop
       End With
   End If
   CheckOC2
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
   End If
   
End Sub

Sub PrintTotil2()
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
   End If
   'Modify By Cheng 2003/05/13
   'strSQL = "SELECT R027006,COUNT(R027006) FROM R040311 WHERE R027002='" & strTemp3(0) & "' AND ID='" & strUserNum & "' GROUP BY R027006 "
   strSql = "SELECT R027010,COUNT(R027010), Decode(CPM03,'（無）',CPM04,CPM03) FROM R040311,CasePropertyMap WHERE R027009=CPM01(+) And R027010=CPM02(+) And R027002='" & strTemp3(0) & "' AND ID='" & strUserNum & "' GROUP BY R027010, Decode(CPM03,'（無）',CPM04,CPM03) Order By 1 "
   CheckOC2
   adoRecordset1.CursorLocation = adUseClient
   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
       With adoRecordset1
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 4
                   'Modify By C
   '                StrTemp5(i) = CheckStr(.Fields(0))
                   StrTemp5(i) = CheckStr(.Fields(2))
                   StrTemp6(i) = CheckStr(.Fields(1))
                   Printer.CurrentX = 500 + (i * 2500)
                   Printer.CurrentY = iPrint
                   Printer.Print StrToStr(StrTemp5(i), 6)
                   Printer.CurrentX = 2500 + (i * 2500) - Printer.TextWidth(StrTemp6(i))
                   Printer.CurrentY = iPrint
                   Printer.Print StrTemp6(i)
                   .MoveNext
                   If .EOF = True Then
                       Exit For
                   End If
               Next i
               iPrint = iPrint + 300
               If iPrint >= 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
                   If .EOF = False Then
                       PrintTitle2
                   End If
               End If
               Loop
       End With
   End If
   CheckOC2
   iPrint = iPrint + 300
   If iPrint > 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle2
   End If
            
End Sub

Sub GetPleft1()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 1800
   PLeft(2) = 3000
   PLeft(3) = 5000
   PLeft(4) = 9000
   PLeft(5) = 11000
   PLeft(6) = 12500
   PLeft(7) = 14000
End Sub

Sub PrintDatil1()
   For i = 0 To 7
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTemp(i)
   Next i
   iPrint = iPrint + 300
End Sub

Sub GetPleft2()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 1700
   PLeft(2) = 3700
   PLeft(3) = 7700
   PLeft(4) = 9700
   PLeft(5) = 11200
   PLeft(6) = 12700
   PLeft(7) = 14200
End Sub

Sub PrintDatil2()
   For i = 0 To 7
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTemp(i)
   Next i
   iPrint = iPrint + 300
End Sub

Sub PrintData2()
   strSql = "SELECT R027002,R027003,R027004,R027005,R027006,R027007,R027001,R027008 FROM R040311 WHERE ID='" & strUserNum & "' ORDER BY 1,2,3 "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   strTemp3(0) = " "
   Page = 1
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           PrintTitle
           'strTemp3(0) = CheckStr(.Fields(0))
           Do While .EOF = False
               For i = 0 To 7
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               strTemp(4) = StrToStr(strTemp(4), 4)
               strTemp(2) = StrToStr(strTemp(2), 16)
               'Modify By Cheng 2003/02/14
               '將員工代號轉成員工名稱
   '            strTemp(6) = StrToStr(strTemp(6), 4)
               strTemp(6) = StrToStr(GetStaffName(strTemp(6), True), 4)
               strTemp(7) = StrToStr(strTemp(7), 4)
               If strTemp3(0) <> strTemp(0) Then
                   strTemp3(0) = strTemp(0)
                   'PrintTotil2
                   PrintTitle2
               Else
                   strTemp(0) = ""
               End If
               If iPrint > 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
                   PrintTitle2
               End If
               PrintDatil2
               .MoveNext
               If .EOF = False Then
                   If strTemp3(0) <> CheckStr(.Fields(0)) Then
                       PrintTotil2
                   End If
               End If
           Loop
       End With
   End If
   PrintTotil2
   Printer.EndDoc
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040311 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/11
   Select Case Index
      Case 1 '准駁代號
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      Case 8 '列印順序
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      'Add By Cheng 2002/12/12
      Case 9 '是否依承辦人跳頁
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 0
        strTemp1 = Split(UCase(GetSystemKindByNick), ",")
        strTemp2 = Split(UCase(txt1(0)), ",")
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
               txt1(0).SetFocus
               txt1(0).SelStart = 0
               txt1(0).SelLength = Len(txt1(0))
               Exit Sub
           End If
        Next i
      Case 3, 7 '准駁日期, 申請國家
         'Modify By Cheng 2002/09/11
         If blnClkSure = False Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      '2011/4/18 add by sonia
      Case 4
         If txt1(Index) = "102" Then txt1(10) = "Y"
      '2011/4/18 end
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
           Select Case Val(txt1(1))
           Case 1, 2
           Case Else
                s = MsgBox("准駁代碼只能 1 或 2 !!", , "USER 輸入錯誤")
                Cancel = True
           End Select
      Case 2, 3 '准駁日期起, 迄
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Cancel = True
         End If
      Case 4
         If txt1(Index) <> "" Then
            lbl1(0) = GetPrjState6HM("P", txt1(Index))
            If lbl1(0) = "" Then
               MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
               Cancel = True
            End If
         End If
      Case 5
         If txt1(Index) <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetStaff(txt1(Index), strExc(0)) Then
            If ClsPDGetStaffN(txt1(Index), strExc(0)) Then
               lbl1(1) = strExc(0)
            Else
               lbl1(1) = ""
               Cancel = True
            End If
         End If
      Case 8
           Select Case Val(txt1(8))
           Case 1, 2
           Case Else
                s = MsgBox("列印順序只能 1 或 2 !!", , "USER 輸入錯誤")
                Cancel = True
           End Select
   End Select
   If Cancel Then TextInverse txt1(Index)
End Sub
