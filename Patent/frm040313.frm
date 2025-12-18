VERSION 5.00
Begin VB.Form frm040313 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問客戶委辦案件明細表"
   ClientHeight    =   1095
   ClientLeft      =   5865
   ClientTop       =   1770
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3285
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2364
      TabIndex        =   5
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1536
      TabIndex        =   4
      Top             =   20
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   3
      Top             =   792
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   2
      Top             =   768
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   1
      Top             =   480
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   0
      Top             =   480
      Width           =   945
   End
   Begin VB.Line Line2 
      X1              =   2112
      X2              =   2232
      Y1              =   936
      Y2              =   936
   End
   Begin VB.Line Line1 
      X1              =   2112
      X2              =   2247
      Y1              =   612
      Y2              =   612
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   180
      Index           =   1
      Left            =   288
      TabIndex        =   7
      Top             =   828
      Width           =   768
   End
   Begin VB.Label Label1 
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   528
      Width           =   936
   End
End
Attribute VB_Name = "frm040313"
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
Dim PLeft(0 To 8) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 4) As String, StrTemp5(0 To 4) As String, StrTemp6(0 To 4) As String
'Add By Cheng 2002/09/11
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
      'Add By Cheng 2002/09/11
      blnClkSure = False
           
     If Len(txt1(0)) = 0 Or Len(txt1(1)) = 0 Then
        s = MsgBox("客戶編號區間不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1_GotFocus (0)
        Exit Sub
     Else
        If Mid(txt1(0), 1, 6) <> Mid(txt1(1), 1, 6) Then
            s = MsgBox("客戶編號前 6 碼必須相同!!", , "USER 輸入錯誤")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        Else
            'Add By Cheng 2002/09/11
            If Me.txt1(0).Text <> "" And Me.txt1(1).Text <> "" Then
               If Me.txt1(0).Text > Me.txt1(1).Text Then
                  MsgBox "客戶編號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(0).SetFocus
                  txt1_GotFocus 0
                  Exit Sub
               End If
            End If
            
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
               'Modify by Morgan 2010/8/11 百年蟲
               'If Me.txt1(2).Text > Me.txt1(3).Text Then
               If Val(txt1(2)) > Val(txt1(3)) Then
                  MsgBox "收文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(2).SetFocus
                  txt1_GotFocus 2
                  Exit Sub
               End If
            End If
            
            If Len(txt1(3)) = 0 Then
                s = MsgBox("收文日區間不可空白!!", , "USER 輸入錯誤")
                txt1(2).SetFocus
                txt1_GotFocus (2)
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
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
'cnnConnection.Execute "DELETE FROM R040313 WHERE ID='" & strUserNum & "' "
Screen.MousePointer = vbHourglass
strSQL1 = ""
strSQL2 = ""
strTemp3(4) = " "

If Len(Trim(txt1(0))) <> 0 Then
   strSQL1 = strSQL1 & " AND HC05>='" & GetNewFagent(txt1(0)) & "' "
End If
If Len(Trim(txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " AND HC05<='" & GetNewFagent(txt1(1)) & "' "
End If
If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/12/2
End If
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(3))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(3))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/2
End If

strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)), " & _
          "CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,HC06,NVL(CPM03,CPM04),ST02,CP16,CP18,'',CP53,CP54 FROM HIRECASE,CUSTOMER,CASEPROGRESS,STAFF,CASEPROPERTYMAP WHERE hC01=Cp01(+) AND hC02=Cp02(+) AND hC03=Cp03(+) AND hC04=Cp04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(HC05,1,8)=cu01(+) AND decode(SUBSTR(HC05,9,1),null,'0',substr(hc05,9,1))=cu02(+) AND CP13=ST01(+) AND CP57 IS NULL " & strSQL1 & " ORDER BY CP05,A "
CheckOC
Page = 1
strTemp3(0) = " "    '控制重複不印      收文日
strTemp3(1) = " "    '控制重複不印      本所案號
strTemp3(2) = "0"    '計算小計          費用
strTemp3(3) = "0"    '計算小計          點數
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/2
    adoRecordset.MoveFirst
    strTemp(0) = CheckStr(adoRecordset.Fields(0))
    strTemp3(4) = strTemp(0)
    PrintTitle
    Do While adoRecordset.EOF = False
        For i = 0 To 8
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
         'Modify By Cheng 2002/06/10
'        strSQL2 = ""
        strSQL2 = " And CP09<'B' "
        If Len(CheckStr(adoRecordset.Fields(9))) <> 0 Then
            strSQL2 = strSQL2 + " AND CP05>=" & Val(CheckStr(adoRecordset.Fields(9)))
        End If
        If Len(CheckStr(adoRecordset.Fields(10))) <> 0 Then
            strSQL2 = strSQL2 + " AND CP05<=" & Val(CheckStr(adoRecordset.Fields(10)))
        End If
        strSQL2 = strSQL2 + " AND CR01='" & SystemNumber(strTemp(2), 1) & "' AND CR02='" & SystemNumber(strTemp(2), 2) & "' AND CR03='" & SystemNumber(strTemp(2), 3) & "' AND CR04='" & SystemNumber(strTemp(2), 4) & "' "
        strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
        strTemp(3) = StrToStr(strTemp(3), 14)
        strTemp(4) = StrToStr(strTemp(4), 4)
        strTemp(8) = StrToStr(strTemp(8), 7)
        PrintDatil
        If iPrint >= 10000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle
        End If
        strTemp3(2) = str(Val(strTemp3(2)) + Val(strTemp(6)))
        strTemp3(3) = str(Val(strTemp3(3)) + Val(strTemp(7)))
        If Val(CheckStr(adoRecordset.Fields(9))) <> 0 And Val(CheckStr(adoRecordset.Fields(10))) <> 0 Then
            'Modify By Cheng 2002/06/10
            '系統類別多加"CFP","CPS"
'            strSQL = "SELECT CP05,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS B,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04),ST02,CP16,CP18,NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)),NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)),NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) FROM CASERELATION,PATENT,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CASEPROGRESS,STAFF,CASEPROPERTYMAP WHERE Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) AND cp01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA27,1,8)=C1.CU01(+) AND SUBSTR(PA27,9,1)=C1.CU02(+) AND SUBSTR(PA28,1,8)=C2.CU01(+) AND SUBSTR(PA28,9,1)=C2.CU02(+) AND SUBSTR(PA29,1,8)=C3.CU01(+) AND SUBSTR(PA29,9,1)=C3.CU02(+) AND SUBSTR(PA30,1,8)=C4.CU01(+) AND SUBSTR(PA30,9,1)=C4.CU02(+) AND CP13=ST01(+) AND CP57 IS NULL AND PA01='P' and cp01='P' AND cr05=PA01(+) AND cr06=PA02(+) AND cr07=PA03(+) AND cr08=PA04(+) " & strSQL2
'            strSQL = strSQL + "  UNION ALL SELECT CP05,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS B,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04),ST02,CP16,CP18,NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)),' ',' ' FROM CASERELATION,SERVICEPRACTICE,CUSTOMER C1,CUSTOMER C2,CASEPROGRESS,STAFF,CASEPROPERTYMAP WHERE sP01=cP01(+) AND sP02=cP02(+) AND sp03=cP03(+) AND sP04=cP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP58,1,8)=C1.CU01(+) AND SUBSTR(SP58,9,1)=C1.CU02(+) AND SUBSTR(SP59,1,8)=C2.CU01(+) AND SUBSTR(SP59,9,1)=C2.CU02(+) AND CP13=ST01(+) AND CP57 IS NULL AND SP01='PS' and cp01='PS' AND cr05=SP01(+) AND cr06=SP02(+) AND cr07=SP03(+) AND cr08=SP04(+) " & strSQL2
'            strSQL = strSQL + " ORDER BY CP05,B "
            strSql = "SELECT CP05,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS B,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04),ST02,CP16,CP18,NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)),NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)),NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)),CP09 FROM CASERELATION,PATENT,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CASEPROGRESS,STAFF,CASEPROPERTYMAP WHERE Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) AND cp01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA27,1,8)=C1.CU01(+) AND SUBSTR(PA27,9,1)=C1.CU02(+) AND SUBSTR(PA28,1,8)=C2.CU01(+) AND SUBSTR(PA28,9,1)=C2.CU02(+) AND SUBSTR(PA29,1,8)=C3.CU01(+) AND SUBSTR(PA29,9,1)=C3.CU02(+) AND SUBSTR(PA30,1,8)=C4.CU01(+) AND SUBSTR(PA30,9,1)=C4.CU02(+) AND CP13=ST01(+) AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL)) " & _
                     " AND ((PA01='P' and cp01='P') OR (PA01='CFP' and cp01='CFP')) AND cr05=PA01(+) AND cr06=PA02(+) AND cr07=PA03(+) AND cr08=PA04(+) " & strSQL2
            'Modify By Sindy 2011/2/21 增加SP65,SP66
            strSql = strSql + " UNION ALL SELECT CP05,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS B,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04),ST02,CP16,CP18,NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)),NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)),NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)),CP09 FROM CASERELATION,SERVICEPRACTICE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CASEPROGRESS,STAFF,CASEPROPERTYMAP WHERE sP01=cP01(+) AND sP02=cP02(+) AND sp03=cP03(+) AND sP04=cP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP58,1,8)=C1.CU01(+) AND SUBSTR(SP58,9,1)=C1.CU02(+) AND SUBSTR(SP59,1,8)=C2.CU01(+) AND SUBSTR(SP59,9,1)=C2.CU02(+) AND SUBSTR(SP65,1,8)=C3.CU01(+) AND SUBSTR(SP65,9,1)=C3.CU02(+) AND SUBSTR(SP66,1,8)=C4.CU01(+) AND SUBSTR(SP66,9,1)=C4.CU02(+) AND CP13=ST01(+) AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL)) " & _
                     " AND ((SP01='PS' and cp01='PS') OR (SP01='CPS' and cp01='CPS')) AND cr05=SP01(+) AND cr06=SP02(+) AND cr07=SP03(+) AND cr08=SP04(+) " & strSQL2
            strSql = strSql + " ORDER BY 2,1,12 "
            
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                With adoRecordset1
                    .MoveFirst
                    strTemp3(0) = " "
                    strTemp3(1) = " "
                    Do While .EOF = False
                        For i = 0 To 7
                            strTemp(i + 1) = CheckStr(.Fields(i))
                        Next i
                        strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                        strTemp(3) = StrToStr(strTemp(3), 14)
                        strTemp(4) = StrToStr(strTemp(4), 4)
                        strTemp(8) = StrToStr(strTemp(8), 7)
                        strTemp3(2) = str(Val(strTemp3(2)) + Val(strTemp(6)))
                        strTemp3(3) = str(Val(strTemp3(3)) + Val(strTemp(7)))
                        If strTemp3(0) <> strTemp(1) Then
                            strTemp3(0) = strTemp(1)
                            strTemp3(1) = strTemp(2)
                        Else
                            strTemp(1) = ""
                            If strTemp3(1) <> strTemp(2) Then
                                strTemp3(1) = strTemp(2)
                            Else
                                strTemp(2) = ""
                            End If
                        End If
                        PrintDatil
                        If iPrint >= 10000 Then
                            Page = Page + 1
                            Printer.NewPage
                            PrintTitle
                        End If
                        For i = 8 To 10
                            If Len(CheckStr(.Fields(i))) <> 0 Then
                                Printer.CurrentX = PLeft(8)
                                Printer.CurrentY = iPrint
                                Printer.Print StrToStr(CheckStr(.Fields(i)), 5)
                                iPrint = iPrint + 300
                                If iPrint >= 10000 Then
                                    Page = Page + 1
                                    Printer.NewPage
                                    PrintTitle
                                End If
                            End If
                        Next i
                        .MoveNext
                    Loop
               End With
            End If
        End If
        adoRecordset.MoveNext
        If adoRecordset.EOF = False Then
            If strTemp3(4) <> CheckStr(adoRecordset.Fields(0)) Then
                PrintEnd
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
        End If
    Loop
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/2
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
    CheckOC
End If
CheckOC
PrintEnd
Printer.EndDoc
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintEnd()
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "小計："
Printer.CurrentX = PLeft(6) + 800 - Printer.TextWidth(Trim(strTemp3(2)))
Printer.CurrentY = iPrint
Printer.Print Format(Trim(strTemp3(2)), "###,###,###,###.00")
Printer.CurrentX = PLeft(7) + 800 - Printer.TextWidth(Trim(strTemp3(3)))
Printer.CurrentY = iPrint
Printer.Print Format(Trim(strTemp3(3)), "###,###,###,###.00")
iPrint = iPrint + 600
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
strTemp3(2) = "0"    '計算小計          費用
strTemp3(3) = "0"    '計算小計          點數
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
Printer.Print "顧問客戶委辦案件明細表"
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
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "客戶名稱：" & strTemp(0)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "費  用"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "其他申請人"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 3500
PLeft(4) = 7700
PLeft(5) = 9000
PLeft(6) = 10350
PLeft(7) = 12000
PLeft(8) = 13500
End Sub

Sub PrintDatil()
For i = 1 To 5
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
Printer.CurrentX = PLeft(6) + 800 - Printer.TextWidth(strTemp(6))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(6), "###,###,###,###.00")
Printer.CurrentX = PLeft(7) + 800 - Printer.TextWidth(strTemp(7))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(7), "###,###,###,###.00")
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print strTemp(8)
iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040313 = Nothing
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
   Case 3
      'Modify By Cheng 2002/09/11
      If blnClkSure = False Then
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
         End If
      Else
         blnClkSure = False
      End If
   Case 1
      'Modify By Cheng 2002/09/11
      If blnClkSure = False Then
         If Len(txt1(Index - 1)) <> 0 Then
            If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
                s = MsgBox("客戶編號前 6 碼必須相同", , "USER 輸入錯誤")
                txt1(Index - 1).SetFocus
                Exit Sub
            End If
         End If
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
         End If
      Else
         blnClkSure = False
      End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 2, 3 '收文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
   End If
End Select
If Cancel Then TextInverse txt1(Index)
End Sub
