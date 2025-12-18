VERSION 5.00
Begin VB.Form frm090622 
   BorderStyle     =   1  '單線固定
   Caption         =   "商申承辦人內部及機關收發文統計表"
   ClientHeight    =   960
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3840
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4620
      TabIndex        =   6
      Text            =   "ALL"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2820
      TabIndex        =   3
      Top             =   90
      Width           =   945
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1830
      TabIndex        =   2
      Top             =   90
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   1
      Left            =   2250
      MaxLength       =   7
      TabIndex        =   1
      Top             =   570
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   570
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "民國年"
      Height          =   180
      Left            =   3210
      TabIndex        =   5
      Top             =   660
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1980
      X2              =   2670
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收/發文區間："
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   1125
   End
End
Attribute VB_Name = "frm090622"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/14 Form2.0已檢查 (無需修改的物件); Printer列印未改
'Create by Sindy 2013/5/29
Option Explicit

Dim strTemp(0 To 18) As String
Dim PLeft(0 To 18) As Integer
Dim iPrint As Integer, intPage As Integer
Dim iTop As Integer, iEnd As Integer, iTop2 As Integer

Dim LineH As Integer

Private Type Cal950705
    UserName As String
    MonthCountIn As Double
    YearCountIn As Double
    MonthCountOut As Double
    CPM As String
    CPMCode As String
    TotMonthCountIn As Double
    TotYearCountIn As Double
    TotMonthCountOut As Double
End Type
Dim oCal() As Cal950705
Dim intCol As Integer
Dim intRunPage As Integer


Private Sub cmdOK_Click(Index As Integer)
Dim bolCancel As Boolean 'Add By Sindy 2016/6/14
Select Case Index
Case 0
     If Trim(txt1(0).Text) = "" Then
        MsgBox "起始日期不可空白！", vbExclamation
        txt1(0).SetFocus
        Exit Sub
     'Add By Sindy 2016/6/14
     Else
        bolCancel = False
        Call txt1_Validate(0, bolCancel)
        If bolCancel = True Then
           txt1(0).SetFocus
           Exit Sub
        End If
     End If
     '2016/6/14 END
     If Trim(txt1(1).Text) = "" Then
        MsgBox "迄止日期不可空白！", vbExclamation
        txt1(1).SetFocus
        Exit Sub
     'Add By Sindy 2016/6/14
     Else
        bolCancel = False
        Call txt1_Validate(1, bolCancel)
        If bolCancel = True Then
           txt1(1).SetFocus
           Exit Sub
        End If
     End If
     '2016/6/14 END
     
     Me.Enabled = False
     Screen.MousePointer = vbHourglass
     ClearQueryLog (Me.Name)
     StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub StrMenu()
Dim strSql As String
Dim orsTmp As New ADODB.Recordset
Dim oI As Integer
Dim oJ As Integer
Dim oK As Integer
Dim ChkCPM As Boolean
Dim intRow As Integer

If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(1)
   InsertQueryLog ("")
End If
LineH = 270

'95.商標處商申
'97.商標處商爭 Add By Sindy 2016/6/14 + 97
'Add By Sindy 2016/6/14 + 人員離職當月也要出現
strSql = "select * from staff where st05 in('95','97') and (st04='1' or (st04='2' and " & Val(Left(txt1(1), Len(txt1(1)) - 2)) + 191100 & "<=substr(st51,1,6))) and st01<'F' order by st15,st01 "
Set orsTmp = New ADODB.Recordset
With orsTmp
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
         ReDim oCal(.RecordCount + 1, 0) As Cal950705
         .MoveFirst
         Do While Not .EOF
             oCal(.AbsolutePosition, 0).UserName = CheckStr(.Fields("st02"))
             .MoveNext
         Loop
         oCal(.RecordCount + 1, 0).UserName = "合計"
         '95.商標處商申
         '97.商標處商爭 Add By Sindy 2016/6/14 + st05='95' ==> st05 in('95','97')
         '抓取承辦人員為商申人員的承辦案件量
         '本月收文
         'Modify By Sindy 2016/6/14
         '取消 decode(cp01,'CFT','國外商標',decode(cp01||tm10,'T020','大陸','')||nvl(cpm03,cpm04)) CPM
         '+ and tm10='000' 或 and sp09='000'
         '95.商標處商申
         '97.商標處商爭 Add By Sindy 2016/6/14 + 97
         'Add By Sindy 2016/6/14 + 人員離職當月也要出現
         'Modify By Sindy 2019/10/15 + 有關「商申內部及機關來函收發文之件數統計」,請針對以下案件性質增加A類收文之件數統計：
'         1. 延期－303
'         2. 補正－201
'         3. 放棄專用權－206
'         4. 檢送同意書－211
         strSql = " select cp14,st02,cp10,CPM,count(oNow05) TNow05,count(oAll05) TAll05,count(oNow27) TNow27 from ("
         strSql = strSql & " select cp14,st02,cp10,cpm03 CPM,oNow05,oAll05,oNow27,cp01 from ("
         strSql = strSql & "       select cp10,cp14,cp09 oNow05,null oAll05,null oNow27,cp01 FROM CASEPROGRESS,trademark WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and CP05>=" & ChangeTStringToWString(txt1(0)) & "  AND CP05<=" & ChangeTStringToWString(txt1(1)) & " AND ((substr(CP09,1,1) < 'D' AND substr(CP09,1,1) >= 'B') or (substr(CP09,1,1) = 'A' and cp10 in('303','201','206','211'))) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 2) & ") And CP10 not in('1001','1701') and tm10='000' and CP57 is null "
         strSql = strSql & " union select cp10,cp14,cp09 oNow05,null oAll05,null oNow27,cp01 FROM CASEPROGRESS,servicepractice WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and CP05>=" & ChangeTStringToWString(txt1(0)) & "  AND CP05<=" & ChangeTStringToWString(txt1(1)) & " AND ((substr(CP09,1,1) < 'D' AND substr(CP09,1,1) >= 'B') or (substr(CP09,1,1) = 'A' and cp10 in('303','201','206','211'))) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 5) & ") And CP10 not in('1001','1701') and sp09='000' and CP57 is null "
         '累計收文
         strSql = strSql & " union select cp10,cp14,null oNow05,cp09 oAll05,null oNow27,cp01 FROM CASEPROGRESS,trademark WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and CP05>=" & Mid(ChangeTStringToWString(txt1(0)), 1, 4) & "0101  AND CP05<=" & ChangeTStringToWString(txt1(1)) & " AND ((substr(CP09,1,1) < 'D' AND substr(CP09,1,1) >= 'B') or (substr(CP09,1,1) = 'A' and cp10 in('303','201','206','211'))) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 2) & ") And CP10 not in('1001','1701') and tm10='000' and CP57 is null "
         strSql = strSql & " union select cp10,cp14,null oNow05,cp09 oAll05,null oNow27,cp01 FROM CASEPROGRESS,servicepractice WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and CP05>=" & Mid(ChangeTStringToWString(txt1(0)), 1, 4) & "0101  AND CP05<=" & ChangeTStringToWString(txt1(1)) & " AND ((substr(CP09,1,1) < 'D' AND substr(CP09,1,1) >= 'B') or (substr(CP09,1,1) = 'A' and cp10 in('303','201','206','211'))) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 5) & ") And CP10 not in('1001','1701') and sp09='000' and CP57 is null "
         '本月發文
         strSql = strSql & " union select cp10,cp14,null oNow05,null oAll05,cp09 oNow27,cp01 FROM CASEPROGRESS,trademark WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and CP27>=" & ChangeTStringToWString(txt1(0)) & "  AND CP27<=" & ChangeTStringToWString(txt1(1)) & " AND ((substr(CP09,1,1) < 'D' AND substr(CP09,1,1) >= 'B') or (substr(CP09,1,1) = 'A' and cp10 in('303','201','206','211'))) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 2) & ") And CP10 not in('1001','1701') and tm10='000' and CP57 is null "
         strSql = strSql & " union select cp10,cp14,null oNow05,null oAll05,cp09 oNow27,cp01 FROM CASEPROGRESS,servicepractice WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and CP27>=" & ChangeTStringToWString(txt1(0)) & "  AND CP27<=" & ChangeTStringToWString(txt1(1)) & " AND ((substr(CP09,1,1) < 'D' AND substr(CP09,1,1) >= 'B') or (substr(CP09,1,1) = 'A' and cp10 in('303','201','206','211'))) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 5) & ") And CP10 not in('1001','1701') and sp09='000' and CP57 is null "
         strSql = strSql & ") AAA,staff,casepropertymap" & _
                           " where cp14=st01(+) and st05 in('95','97') and st01<'F'" & _
                           " and (st04='1' or (st04='2' and " & Val(Left(txt1(1), Len(txt1(1)) - 2)) + 191100 & "<=substr(st51,1,6)))" & _
                           " and cp01=cpm01(+) and cp10=cpm02(+)) BBB" & _
                           " group by cp14,st02,cp10,CPM order by cp10 "
         '2016/6/14 ENd
         Set orsTmp = New ADODB.Recordset
         If .State = 1 Then .Close
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount <> 0 Then
             .MoveFirst
             Do While Not .EOF
                 For oI = 1 To UBound(oCal, 1)
                     If CheckStr(.Fields("st02")) = oCal(oI, 0).UserName Then
                         ChkCPM = False
                         If UBound(oCal, 2) = 0 Then
                             ChkCPM = False
                         Else
                             '檢查有無案件性質
                             For oJ = 1 To UBound(oCal, 2)
                                 'If oCal(oI, oJ).CPMCode = CheckStr(.Fields("oSys")) Then
                                     If oCal(oI, oJ).CPM = CheckStr(.Fields("CPM")) Then
                                         oCal(oI, oJ).MonthCountIn = oCal(oI, oJ).MonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                         oCal(oI, oJ).MonthCountOut = oCal(oI, oJ).MonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                         oCal(oI, oJ).YearCountIn = oCal(oI, oJ).YearCountIn + Val(CheckStr(.Fields("TAll05")))
                                         oCal(UBound(oCal, 1), oJ).MonthCountIn = oCal(UBound(oCal, 1), oJ).MonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                         oCal(UBound(oCal, 1), oJ).MonthCountOut = oCal(UBound(oCal, 1), oJ).MonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                         oCal(UBound(oCal, 1), oJ).YearCountIn = oCal(UBound(oCal, 1), oJ).YearCountIn + Val(CheckStr(.Fields("TAll05")))
                                         oCal(oI, 0).TotMonthCountIn = oCal(oI, 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                         oCal(oI, 0).TotMonthCountOut = oCal(oI, 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                         oCal(oI, 0).TotYearCountIn = oCal(oI, 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                                         oCal(UBound(oCal, 1), 0).TotMonthCountIn = oCal(UBound(oCal, 1), 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                         oCal(UBound(oCal, 1), 0).TotMonthCountOut = oCal(UBound(oCal, 1), 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                         oCal(UBound(oCal, 1), 0).TotYearCountIn = oCal(UBound(oCal, 1), 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                                         ChkCPM = True
                                         Exit For
                                     End If
                                 'End If
                             Next oJ
                         End If
                         If ChkCPM = False Then
                             '還沒新增任何案件性質時
                             ReDim Preserve oCal(UBound(oCal, 1), UBound(oCal, 2) + 1) As Cal950705
                             For oK = 1 To UBound(oCal, 1)
                                 oCal(oK, UBound(oCal, 2)).CPM = CheckStr(.Fields("CPM"))
                                 'oCal(oK, UBound(oCal, 2)).CPMCode = CheckStr(.Fields("oSys"))
                             Next oK
                             oCal(oI, UBound(oCal, 2)).MonthCountIn = CheckStr(.Fields("TNow05"))
                             oCal(oI, UBound(oCal, 2)).MonthCountOut = CheckStr(.Fields("TNow27"))
                             oCal(oI, UBound(oCal, 2)).YearCountIn = CheckStr(.Fields("TAll05"))
                             oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountIn = oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountIn + Val(CheckStr(.Fields("TNow05")))
                             oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountOut = oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountOut + Val(CheckStr(.Fields("TNow27")))
                             oCal(UBound(oCal, 1), UBound(oCal, 2)).YearCountIn = oCal(UBound(oCal, 1), UBound(oCal, 2)).YearCountIn + Val(CheckStr(.Fields("TAll05")))
                             oCal(oI, 0).TotMonthCountIn = oCal(oI, 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                             oCal(oI, 0).TotMonthCountOut = oCal(oI, 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                             oCal(oI, 0).TotYearCountIn = oCal(oI, 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                             oCal(UBound(oCal, 1), 0).TotMonthCountIn = oCal(UBound(oCal, 1), 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                             oCal(UBound(oCal, 1), 0).TotMonthCountOut = oCal(UBound(oCal, 1), 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                             oCal(UBound(oCal, 1), 0).TotYearCountIn = oCal(UBound(oCal, 1), 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                         End If
                         Exit For
                     End If
                 Next oI
                 .MoveNext
             Loop
         Else
             MsgBox "沒有資料可以列印！", vbExclamation
             Exit Sub
         End If
    Else
         MsgBox "商標處無人在職！沒有資料可印", vbExclamation
         Exit Sub
    End If
End With
'印表
'Add By Sindy 2011/11/1
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
Printer.PaperSize = 9  'PDF
'2011/11/1 End
'intPage = 1
'Add By Sindy 2010/6/1
intRunPage = 1
'Modify By Sindy 2017/8/14
'If UBound(oCal, 1) >= intCol + 1 Then
'   intRunPage = 2
'End If
If (UBound(oCal, 1) / intCol) > Int(UBound(oCal, 1) / intCol) Then
   intRunPage = Int(UBound(oCal, 1) / intCol) + 1
Else
   intRunPage = Int(UBound(oCal, 1) / intCol)
End If
'2017/8/14 END
For intPage = 1 To intRunPage
   If intPage >= 2 Then Printer.NewPage
'2010/6/1 End
   PrintTitle
   For oI = 0 To intCol * 3
       strTemp(oI) = ""
   Next oI
   'Add By Sindy 2010/6/1
   If intPage = 1 Then '第1頁
      If UBound(oCal, 1) < intCol + 1 Then
         intRow = UBound(oCal, 1)
      Else
         intRow = intCol
      End If
'   End If
'   If intPage = 2 Then '第2頁
'      intRow = UBound(oCal, 1) - intCol
'   End If
   'Modify By Sindy 2017/8/14
   Else
      If intPage = intRunPage Then '最後一頁
         intRow = UBound(oCal, 1) - (intCol * (intPage - 1))
      Else
         intRow = intCol
      End If
   End If
   '2017/8/14 END
   '2010/6/1 End
   For oJ = 1 To UBound(oCal, 2)
       'Add By Sindy 2011/11/1
       'Modify By Sindy 2012/1/2
       'If oJ >= 33 Then
       If (oJ Mod 33) = 0 Then
       '2012/1/2 End
          ShowLine True
          iEnd = iPrint - 20
          PrintEndLine
          Printer.NewPage
          PrintTitle
       End If
       '2011/11/1 End
       
       strTemp(0) = oCal(1, oJ).CPM
       Printer.CurrentX = PLeft(0) + 20
       Printer.CurrentY = iPrint
       Printer.Print StrToStr(strTemp(0), 6) '案件性質
       'For oI = 1 To UBound(oCal, 1)
       For oI = 1 To intRow
           'Debug.Print oCal(oI, 0).UserName; oCal(oI, oJ).CPM; oCal(oI, oJ).MonthCountIn; oCal(oI, oJ).YearCountIn; oCal(oI, oJ).MonthCountOut
           If intPage = 1 Then
               strTemp(1) = IIf(oCal(oI, oJ).MonthCountIn = 0, "", oCal(oI, oJ).MonthCountIn)
               strTemp(2) = IIf(oCal(oI, oJ).YearCountIn = 0, "", oCal(oI, oJ).YearCountIn)
               strTemp(3) = IIf(oCal(oI, oJ).MonthCountOut = 0, "", oCal(oI, oJ).MonthCountOut)
'           ElseIf intPage = 2 Then
'               strTemp(1) = IIf(oCal(oI + intCol, oJ).MonthCountIn = 0, "", oCal(oI + intCol, oJ).MonthCountIn)
'               strTemp(2) = IIf(oCal(oI + intCol, oJ).YearCountIn = 0, "", oCal(oI + intCol, oJ).YearCountIn)
'               strTemp(3) = IIf(oCal(oI + intCol, oJ).MonthCountOut = 0, "", oCal(oI + intCol, oJ).MonthCountOut)
           'Modify By Sindy 2017/8/14
           Else
               strTemp(1) = IIf(oCal(oI + (intCol * (intPage - 1)), oJ).MonthCountIn = 0, "", oCal(oI + (intCol * (intPage - 1)), oJ).MonthCountIn)
               strTemp(2) = IIf(oCal(oI + (intCol * (intPage - 1)), oJ).YearCountIn = 0, "", oCal(oI + (intCol * (intPage - 1)), oJ).YearCountIn)
               strTemp(3) = IIf(oCal(oI + (intCol * (intPage - 1)), oJ).MonthCountOut = 0, "", oCal(oI + (intCol * (intPage - 1)), oJ).MonthCountOut)
           '2017/8/14 END
           End If
           Printer.CurrentX = PLeft(1 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(1))) - 20
           Printer.CurrentY = iPrint
           Printer.Print strTemp(1)
           Printer.CurrentX = PLeft(2 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(2))) - 20
           Printer.CurrentY = iPrint
           Printer.Print strTemp(2)
           Printer.CurrentX = PLeft(3 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(3))) - 20
           Printer.CurrentY = iPrint
           Printer.Print strTemp(3)
       Next oI
       ShowLine True
       iPrint = iPrint + LineH
       If iPrint + LineH >= Printer.ScaleHeight Then
           iEnd = iPrint + LineH
           PrintEndLine
           PrintTitle
       End If
   Next oJ
   ShowLine True
   Printer.CurrentX = PLeft(0) + 20
   Printer.CurrentY = iPrint
   Printer.Print "個人總計"
   'For oI = 1 To UBound(oCal, 1)
   For oI = 1 To intRow
       'Debug.Print oCal(oI, 0).UserName; oCal(oI, oJ).CPM; oCal(oI, oJ).MonthCountIn; oCal(oI, oJ).YearCountIn; oCal(oI, oJ).MonthCountOut
       If intPage = 1 Then
            strTemp(1) = IIf(oCal(oI, 0).TotMonthCountIn = 0, "", oCal(oI, 0).TotMonthCountIn)
            strTemp(2) = IIf(oCal(oI, 0).TotYearCountIn = 0, "", oCal(oI, 0).TotYearCountIn)
            strTemp(3) = IIf(oCal(oI, 0).TotMonthCountOut = 0, "", oCal(oI, 0).TotMonthCountOut)
'       ElseIf intPage = 2 Then
'            strTemp(1) = IIf(oCal(oI + intCol, 0).TotMonthCountIn = 0, "", oCal(oI + intCol, 0).TotMonthCountIn)
'            strTemp(2) = IIf(oCal(oI + intCol, 0).TotYearCountIn = 0, "", oCal(oI + intCol, 0).TotYearCountIn)
'            strTemp(3) = IIf(oCal(oI + intCol, 0).TotMonthCountOut = 0, "", oCal(oI + intCol, 0).TotMonthCountOut)
       'Modify By Sindy 2017/8/14
       Else
            strTemp(1) = IIf(oCal(oI + (intCol * (intPage - 1)), 0).TotMonthCountIn = 0, "", oCal(oI + (intCol * (intPage - 1)), 0).TotMonthCountIn)
            strTemp(2) = IIf(oCal(oI + (intCol * (intPage - 1)), 0).TotYearCountIn = 0, "", oCal(oI + (intCol * (intPage - 1)), 0).TotYearCountIn)
            strTemp(3) = IIf(oCal(oI + (intCol * (intPage - 1)), 0).TotMonthCountOut = 0, "", oCal(oI + (intCol * (intPage - 1)), 0).TotMonthCountOut)
       '2017/8/14 END
       End If
       Printer.CurrentX = PLeft(1 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(1))) - 20
       Printer.CurrentY = iPrint
       Printer.Print strTemp(1)
       Printer.CurrentX = PLeft(2 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(2))) - 20
       Printer.CurrentY = iPrint
       Printer.Print strTemp(2)
       Printer.CurrentX = PLeft(3 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(3))) - 20
       Printer.CurrentY = iPrint
       Printer.Print strTemp(3)
   Next oI
   iPrint = iPrint + LineH
   iPrint = iPrint + 20
   ShowLine
   iEnd = iPrint - 20
   PrintEndLine
Next intPage
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
intCol = 6
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Printer.DrawWidth = 1
   Set frm090622 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) <> "" Then 'Add By Sindy 2023/12/8 +if
      If CheckIsTaiwanDate(txt1(Index), True) = False Then
         Cancel = True
         txt1(Index).SetFocus
         txt1_GotFocus Index
         Exit Sub
      End If
      If Index = 4 Then
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            Cancel = True
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
         End If
      End If
   End If
End Sub

Sub PrintTitle()
Dim i As Integer
Dim intRow As Integer

GetPleft
iPrint = 0
Printer.DrawWidth = 10
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(GetTitleNick & "商申內部及機關來函收發文之件數統計") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商申內部及機關來函收發文之件數統計"
iPrint = iPrint + Printer.TextHeight(GetTitleNick & "商申內部及機關來函收發文之件數統計")
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 230
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "日期：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "頁　　數：" & Printer.Page 'intPage
iPrint = iPrint + 230
iTop = iPrint - 20
ShowLine
'Add By Sindy 2010/6/1
If intPage = 1 Then '第1頁
   If UBound(oCal, 1) < intCol + 1 Then
      intRow = UBound(oCal, 1)
   Else
      intRow = intCol
   End If
'End If
'If intPage = 2 Then '第2頁
'   intRow = UBound(oCal, 1) - intCol
'End If
'Modify By Sindy 2017/8/14
Else
   If intPage = intRunPage Then '最後一頁
      intRow = UBound(oCal, 1) - (intCol * (intPage - 1))
   Else
      intRow = intCol
   End If
End If
'2017/8/14 END
'2010/6/1 End
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
   Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + ((((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) * 3) - Printer.TextWidth(oCal(i, 0).UserName)) / 2)
   Printer.CurrentY = iPrint
   Printer.Print oCal(i + (intCol * (intPage - 1)), 0).UserName
'   If intPage = 1 Then
'      Printer.Print oCal(i, 0).UserName
'   ElseIf intPage = 2 Then
'      Printer.Print oCal(i + intCol, 0).UserName
'   End If
Next i
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (PLeft(1), iPrint - 20)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iPrint - 20)
iTop2 = iPrint - 20
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("本")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "本"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("累")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "累"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("本")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "本"
Next i
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0) + (((PLeft(1) - PLeft(0)) - (Printer.TextWidth("案件性質"))) / 2)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("月")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "月"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("計")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "計"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("月")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "月"
Next i
iPrint = iPrint + 300
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("收")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "收"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("收")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "收"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("送")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "送"
Next i
iPrint = iPrint + 300
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("文")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "文"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("文")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "文"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("件")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "件"
Next i
iPrint = iPrint + 300
ShowLine True
End Sub

Sub GetPleft()
Dim i As Integer
Erase PLeft
PLeft(0) = 0
For i = 1 To intCol * 3
    PLeft(i) = 900 + ((Printer.TextWidth("　　") * 2) * i - 1)
Next i
End Sub

Sub ShowLine(Optional oContent As Boolean = False)
If oContent = False Then
    Printer.DrawWidth = 10
    Printer.CurrentX = 0
    Printer.CurrentY = iPrint
    Printer.Line (PLeft(0), iPrint - 20)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iPrint - 20)
Else
    Printer.DrawWidth = 10
    Printer.Line (PLeft(0), iPrint - 20)-(PLeft(1), iPrint - 20)
    Printer.DrawWidth = 2
    Printer.Line (PLeft(1), iPrint - 20)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iPrint - 20)
End If
End Sub

Sub PrintEndLine()
Dim i As Integer
For i = 0 To UBound(PLeft)
    If (i - 1) Mod 3 = 0 Or (i - 1) Mod 3 = -1 Then
        Printer.Line (PLeft(i), iTop)-(PLeft(i), iEnd)
    Else
        Printer.DrawWidth = 2
        Printer.Line (PLeft(i), iTop2)-(PLeft(i), iEnd)
        Printer.DrawWidth = 10
    End If
Next i
Printer.Line (PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iTop)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iEnd)
End Sub
