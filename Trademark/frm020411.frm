VERSION 5.00
Begin VB.Form frm020411 
   BorderStyle     =   1  '單線固定
   Caption         =   "各區收/發文達成比較表"
   ClientHeight    =   2670
   ClientLeft      =   5190
   ClientTop       =   2540
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4200
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   180
      TabIndex        =   8
      Top             =   1980
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   9
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1935
      MaxLength       =   5
      TabIndex        =   0
      Top             =   780
      Width           =   930
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3360
      TabIndex        =   4
      Top             =   135
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2565
      TabIndex        =   3
      Top             =   135
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1110
      Width           =   450
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1935
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1110
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(統計非整個月的資料時才需輸入日區間)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   7
      Top             =   1560
      Width           =   3180
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2370
      X2              =   2580
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "收/發文日區間："
      Height          =   180
      Index           =   2
      Left            =   660
      TabIndex        =   6
      Top             =   1140
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "收/發文年月："
      Height          =   180
      Index           =   1
      Left            =   660
      TabIndex        =   5
      Top             =   795
      Width           =   1155
   End
End
Attribute VB_Name = "frm020411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String
Dim strSQL2 As String, iPrint As Double, Page As Integer, strTemp(0 To 9) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 7) As String
Dim PLeft(0 To 7) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, SeekPrint As Integer, SeekPrintL As Integer, StrSQL7 As String, strSQL8 As String
'Add By Cheng 2002/02/22
Dim m_Parts As Integer '1:上表, 2:下表
Dim m_dblSubTotal As Double '各區合計
Dim m_dblTotal As Double '全所合計
Dim strPrinter As String 'Add By Sindy 2015/7/3


Private Sub cmdok_Click(Index As Integer)
'Add By Cheng 2002/02/21
Dim strDateFrom As String
Dim strDateTo As String

Select Case Index
Case 0 '確定
   If CheckKeyIn(1) = -1 Then Exit Sub
   If CheckKeyIn(2) = -1 Then Exit Sub
   If CheckKeyIn(3) = -1 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Printer.EndDoc 'Add By Sindy 2011/11/1
   Me.Enabled = False
   PUB_RestorePrinter Combo1.Text 'Add By Sindy 2015/7/3
   
   TestOk = False
   Page = 1
   'Modify By Cheng 2002/02/21
'   Process
   'Modify By Sindy 2011/2/1
   If Len("" & Me.txt1(1).Text) = 5 Then
      strDateFrom = IIf(Len("" & Me.txt1(2).Text) <= 0, Format(Me.txt1(1).Text, "00000") & "01", Format(Me.txt1(1).Text, "00000") & Format(Me.txt1(2).Text, "00")) + 19110000
      strDateTo = IIf(Len("" & Me.txt1(2).Text) <= 0, Format(Me.txt1(1).Text, "00000") & PUB_GetMonthDays(Mid(Me.txt1(1).Text, 1, 3) + 1911, Mid(Me.txt1(1).Text, 4, 2)), Format(Me.txt1(1).Text, "00000") & Format(Me.txt1(3).Text, "00")) + 19110000
   '2011/2/1 End
   Else
      strDateFrom = IIf(Len("" & Me.txt1(2).Text) <= 0, Format(Me.txt1(1).Text, "0000") & "01", Format(Me.txt1(1).Text, "0000") & Format(Me.txt1(2).Text, "00")) + 19110000
      strDateTo = IIf(Len("" & Me.txt1(2).Text) <= 0, Format(Me.txt1(1).Text, "0000") & PUB_GetMonthDays(Mid(Me.txt1(1).Text, 1, 2) + 1911, Mid(Me.txt1(1).Text, 3, 2)), Format(Me.txt1(1).Text, "0000") & Format(Me.txt1(3).Text, "00")) + 19110000
   End If
   iPrint = 0: m_Parts = 1: m_dblSubTotal = 0: m_dblTotal = 0
   Process strDateFrom, strDateTo
   '若列印整月資料
   If Len("" & Me.txt1(2).Text) <= 0 Then
      m_Parts = 2: m_dblSubTotal = 0: m_dblTotal = 0
      'Modify By Sindy 2011/2/1
      If Len("" & Me.txt1(1).Text) = 5 Then
         strDateFrom = Val(Format(Mid(Me.txt1(1).Text, 1, 3), "000") & "0101") + 19110000
         strDateTo = Val(Format(Me.txt1(1).Text, "00000") & "31") + 19110000
      '2011/2/1 End
      Else
         strDateFrom = Val(Format(Mid(Me.txt1(1).Text, 1, 2), "00") & "0101") + 19110000
         strDateTo = Val(Format(Me.txt1(1).Text, "0000") & "31") + 19110000
      End If
      Process strDateFrom, strDateTo
   End If
   Printer.EndDoc
   ShowPrintOk
   
   PUB_RestorePrinter strPrinter 'Add By Sindy 2015/7/3
   Me.Enabled = True
   Screen.MousePointer = vbDefault
Case 1 '結束
   Unload Me
Case Else
End Select
End Sub

Sub Process1(ByRef strDateFrom As String, ByRef strDateTo As String)
cnnConnection.Execute "DELETE FROM R020411 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
'依列印條件抓案件進度檔中系統類別為"T", 案件性質為"申請"
strSQL1 = strSQL1 + " AND CP01 = 'T' AND CP10='101' "
StrSQL6 = ""
'是否算案件數欄為空白, 無取消收文日的資料
StrSQL6 = StrSQL6 + " AND (CP26='' OR CP26 IS NULL) AND (CP57='' OR CP57 IS NULL) "
CheckOC
StrSQL7 = ""
strSQL8 = ""
'收發文日區間
If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 Then
   StrSQL7 = StrSQL7 + " AND CP05>=" & Val(strDateFrom) & " "
   strSQL8 = strSQL8 + " AND CP27>=" & Val(strDateFrom) & " "
End If
'收發文日區間
If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(3))) <> 0 Then
   StrSQL7 = StrSQL7 + " AND CP05<=" & Val(strDateTo) & " "
   strSQL8 = strSQL8 + " AND CP27<=" & Val(strDateTo) & " "
End If
'整月
If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) = 0 And Len(Trim(txt1(3))) = 0 Then
   'Modify By Cheng 2002/02/19
'   StrSQL7 = StrSQL7 + " AND CP05>=" & Val(Mid(str(Val(ChangeWDateStringToWString(DateAdd("M", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")))))), 1, 8)) & " AND CP05<=" & Val(Mid(str(Val(ChangeWDateStringToWString(DateAdd("M", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "31")))))), 1, 8)) & " "
'   StrSQL8 = StrSQL8 + " AND CP27>=" & Val(Mid(str(Val(ChangeWDateStringToWString(DateAdd("M", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")))))), 1, 8)) & " AND CP27<=" & Val(Mid(str(Val(ChangeWDateStringToWString(DateAdd("M", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "31")))))), 1, 8)) & " "
   StrSQL7 = StrSQL7 + " AND CP05>=" & Val(Mid(strDateFrom, 1, 6) & "01") & " AND CP05<=" & Val(Mid(strDateTo, 1, 6) & Format(PUB_GetMonthDays(Mid(strDateTo, 1, 4), Mid(strDateTo, 5, 2)), "00")) & " "
   strSQL8 = strSQL8 + " AND CP27>=" & Val(Mid(strDateFrom, 1, 6) & "01") & " AND CP27<=" & Val(Mid(strDateTo, 1, 6) & Format(PUB_GetMonthDays(Mid(strDateTo, 1, 4), Mid(strDateTo, 5, 2)), "00")) & " "
End If
'Modify By Cheng 2002/02/21
'依列印條件抓案件進度檔中系統類別為"T", 案件性質為"申請", 是否算案件數欄為空白, 無取消收文日的資料
'計算收文件數時, 只計算CP09<"B"的資料, 計算發文件數時, 則計算CP09<"C"的資料
'抓商標基本檔時, 只抓申請國家為"台灣"(屬國內), 或"大陸"(屬大陸)的資料
'欄位名稱--部門名稱, 智權人員代號, '',收文件數,'',發文件數,'',員工部門代號
'先收後發
'strSQL = "SELECT NVL(A0902,A0903),CP13,'','1','','','',ST03 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',ST03 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL8
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','1','','','',ST03 FROM CASEPROGRESS,SERVICEPRACTICE,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',ST03 FROM CASEPROGRESS,SERVICEPRACTICE,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL8
'Modify By Cheng 2002/04/09
'收文與發文皆抓CP09<"B"的資料
'                    strSQL = "SELECT NVL(A0902,A0903),CP13,'','1','','','',ST03,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND ST03=A0901(+) AND CP13=ST01(+) AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & StrSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',ST03,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND ST03=A0901(+) AND CP13=ST01(+) AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & StrSQL8
'Modify By Cheng 2003/12/30
'加商品類別
'                    strSQL = "SELECT NVL(A0902,A0903),CP13,'','1','','','',CP12,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '101'=CP10 AND CP09<'B' AND CP12=A0901(+) AND CP13=ST01(+) AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',CP12,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '101'=CP10 AND CP09<'B' AND CP12=A0901(+) AND CP13=ST01(+) AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL8
                    strSql = "SELECT NVL(A0902,A0903),CP13,'','1','','','',CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '101'=CP10 AND CP09<'B' AND CP12=A0901(+) AND CP13=ST01(+) AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & StrSQL7
strSql = strSql + " union all select NVL(A0902,A0903),CP13,'','','','1','',CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '101'=CP10 AND CP09<'B' AND CP12=A0901(+) AND CP13=ST01(+) AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL8
'End
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        DoEvents
        Do While .EOF = False
'            For i = 0 To 7
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2002/02/20
'            strTemp(1) = GetPerformanceByNick(Val("0000" & Format(DateAdd("M", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & Format(txt1(2), "00")))), "mm")), Val("0000" & Format(DateAdd("M", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & Format(txt1(3), "00")))), "mm")), txt1(0), strTemp(1))
            'Modify By Cheng 2002/04/09
            '改成列印時才抓目標值
'            strTemp(1) = GetPerformanceByNick_1(Mid(strDateFrom, 5, 2), Mid(strDateTo, 5, 2), "T", strTemp(7))
'            '若非整月, 則要換算目標件數
'            If Len("" & Me.txt1(2).Text) > 0 Then
'               strTemp(1) = strTemp(1) * (GetWorkDay(strDateTo, strDateFrom)) / (GetWorkDay(Mid(strDateTo, 1, 6) & "01", Mid(strDateFrom, 1, 6)) & "31")
'            End If
             strTemp(1) = 0
'            strTemp(2) = str(Val(strTemp(1)) / 2)
            'Add By Cheng 2003/12/30
            '若申請國家為台灣
            If "" & .Fields(8).Value = "000" Then
                '若為收文件數
                If strTemp(3) <> "" Then
                    '若商品類別有值
                    If "" & .Fields(10).Value <> "" Then
                        strTemp(3) = UBound(Split("" & .Fields(10).Value, ",")) + 1
                    End If
                '若為發文件數
                Else
                    '若商品類別有值
                    If "" & .Fields(10).Value <> "" Then
                        strTemp(5) = UBound(Split("" & .Fields(10).Value, ",")) + 1
                    End If
                End If
            End If
            'End
'            strSQL = "INSERT INTO R020411 VALUES ('" & chgsql(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",''," & Val(strTemp(5)) & ",'','" & chgsql(strTemp(7)) & "','" & strUserNum & "') "
            strSql = "INSERT INTO R020411 VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",''," & Val(strTemp(5)) & ",'','" & ChgSQL(strTemp(7)) & "','" & strUserNum & "','" & ChgSQL(strTemp(8)) & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    End If
End With
End Sub
'Modify By Cheng 2002/02/21
'強制傳入西元日期區間
'Sub Process()
Sub Process(ByRef strDateFrom As String, ByRef strDateTo As String)
ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020411 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
'依列印條件抓案件進度檔中系統類別為"T", 案件性質為"申請"
strSQL2 = ""
strSQL1 = strSQL1 + " AND CP01 = 'T' AND CP10='101' "
StrSQL6 = ""
'是否算案件數欄為空白, 無取消收文日的資料
StrSQL6 = StrSQL6 + " AND (CP26='' OR CP26 IS NULL) AND (CP57='' OR CP57 IS NULL) "
StrSQL7 = ""
strSQL8 = ""
'收發文日區間
If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 Then
'   StrSQL7 = StrSQL7 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(2)))) & " "
'   StrSQL8 = StrSQL8 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(2)))) & " "
   StrSQL7 = StrSQL7 + " AND CP05>=" & Val(strDateFrom) & " "
   strSQL8 = strSQL8 + " AND CP27>=" & Val(strDateFrom) & " "
End If
'收發文日區間
If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(3))) <> 0 Then
'   StrSQL7 = StrSQL7 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(3)))) & " "
'   StrSQL8 = StrSQL8 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(3)))) & " "
   StrSQL7 = StrSQL7 + " AND CP05<=" & Val(strDateTo) & " "
   strSQL8 = strSQL8 + " AND CP27<=" & Val(strDateTo) & " "
End If
If (Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0) Or _
   (Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(3))) <> 0) Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & (Val(strDateFrom) - 19110000) & "-" & (Val(strDateTo) - 19110000) 'Add By Sindy 2010/10/19
End If
'整月
If Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) = 0 And Len(Trim(txt1(3))) = 0 Then
   'Modify By Cheng 2002/02/20
'   StrSQL7 = StrSQL7 + " AND CP05>=" & Val(Mid(str(Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(3))))), 1, 6)) & "01 AND CP27<= " & Val(Mid(str(Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(3))))), 1, 6)) & "31 "
'   StrSQL8 = StrSQL8 + " AND CP27>=" & Val(Mid(str(Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(3))))), 1, 6)) & "01 AND CP27<= " & Val(Mid(str(Val(ChangeTStringToWString(txt1(1) & ChgNumByNick(txt1(3))))), 1, 6)) & "31 "
   StrSQL7 = StrSQL7 + " AND CP05>=" & Mid(strDateFrom, 1, 6) & "01" & " AND CP05<= " & Mid(strDateTo, 1, 6) & PUB_GetMonthDays(Mid(strDateTo, 1, 4), Mid(strDateTo, 5, 2)) & " "
   strSQL8 = strSQL8 + " AND CP27>=" & Mid(strDateFrom, 1, 6) & "01" & " AND CP27<= " & Mid(strDateTo, 1, 6) & PUB_GetMonthDays(Mid(strDateTo, 1, 4), Mid(strDateTo, 5, 2)) & " "
   pub_QL05 = pub_QL05 & ";" & Label1(1) & (Val(Mid(strDateFrom, 1, 6)) - 191100) 'Add By Sindy 2010/10/19
End If
CheckOC
'Modify By Cheng 2002/02/21
'依列印條件抓案件進度檔中系統類別為"T", 案件性質為"申請", 是否算案件數欄為空白, 無取消收文日的資料
'計算收文件數時, 只計算CP09<"B"的資料, 計算發文件數時, 則計算CP09<"C"的資料
'抓商標基本檔時, 只抓申請國家為"台灣"(屬國內), 或"大陸"(屬大陸)的資料
'欄位名稱--部門名稱, 智權人員代號, '',收文件數,'',發文件數,'',員工部門代號
'先收後發
'                strSQL = "SELECT NVL(A0902,A0903),CP13,'','1','','','',ST03 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101' AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',ST03 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL8
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','1','','','',ST03 FROM CASEPROGRESS,SERVICEPRACTICE,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806'  AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',ST03 FROM CASEPROGRESS,SERVICEPRACTICE,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806'  AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL8
'                strSQL = "SELECT NVL(A0902,A0903),CP13,'','1','','','',ST03,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND ST03=A0901(+) AND CP13=ST01(+) AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & StrSQL7
'Modify By Cheng 2002/04/08
'發文也抓CP09<"B"
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',ST03,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND ST03=A0901(+) AND CP13=ST01(+) AND CP09<'C' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & StrSQL8
'Modify By Cheng 2003/12/30
'加商品類別
'                    strSQL = "SELECT NVL(A0902,A0903),CP13,'','1','','','',CP12,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',CP12,TM10,CP09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL8
'Modify By Cheng 2004/04/13
'將中四區(S24)的資料合併至中二區(S22)
'                    strSQL = "SELECT NVL(A0902,A0903),CP13,'','1','','','',CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL7
'strSQL = strSQL + " union all select NVL(A0902,A0903),CP13,'','','','1','',CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL8
'Modified by Morgan 2016/2/2 105年1月起又有中四區
'                    strSql = "SELECT NVL(A0902,A0903),CP13,'','1','','','',Decode(CP12,'S24','S22',CP12) As CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND Decode(CP12,'S24','S22',CP12)=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & StrSQL7
'strSql = strSql + " union all select NVL(A0902,A0903),CP13,'','','','1','',Decode(CP12,'S24','S22',CP12) As CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND Decode(CP12,'S24','S22',CP12)=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL8
                    strSql = "SELECT NVL(A0902,A0903),CP13,'','1','','','',CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & StrSQL7
strSql = strSql + " union all select NVL(A0902,A0903),CP13,'','','','1','',CP12,TM10,CP09, TM09 FROM CASEPROGRESS,TRADEMARK,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=ST01(+) AND '101'=CP10 AND CP09<'B' AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "') " & strSQL1 & StrSQL6 & strSQL8
'End

CheckOC 'Add By Sindy 2025/9/2
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2002/02/19
'            strTemp(1) = GetPerformanceByNick(Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & Format(txt1(2), "00"))), "mm")), Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & Format(txt1(3), "00"))), "mm")), txt1(0), strTemp(1))
            
            'Modify By Cheng 2002/04/09
            '改成列印時才去抓目標值
'            strTemp(1) = GetPerformanceByNick_1(Val(Mid(strDateFrom, 5, 2)), Val(Mid(strDateTo, 5, 2)), "T", strTemp(7))
'            '若非整月, 則要換算目標件數
'            If Len("" & Me.txt1(2).Text) > 0 Then
'               strTemp(1) = strTemp(1) * (GetWorkDay(strDateTo, strDateFrom)) / (GetWorkDay(Mid(strDateTo, 1, 6) & "31", Mid(strDateFrom, 1, 6) & "01"))
'            End If
             strTemp(1) = 0
'            If Format(ChangeWStringToWDateString(GetTodayDate), "MM") = Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "MM") Then
'                strTemp(2) = str(Val(strTemp(1)) * 15 / Val(Format(ChangeWStringToWDateString(GetTodayDate), "DD")))
'            Else
'                strTemp(2) = str(Val(strTemp(1)) / 2)
'            End If
            'Modify By Cheng 2003/07/01
            '廣東所歸在非業務區
'            If UCase(Mid(strTemp(7), 1, 1)) <> "S" Then
            If UCase(Mid(strTemp(7), 1, 1)) <> "S" Or strTemp(7) = "S91" Then
               strTemp(0) = "非業務區"
               strTemp(7) = "000"
            End If
            'Add By Cheng 2003/12/30
            '若申請國家為台灣
            If "" & .Fields(8).Value = "000" Then
                '若為收文件數
                If strTemp(3) <> "" Then
                    '若商品類別有值
                    If "" & .Fields(10).Value <> "" Then
                        strTemp(3) = UBound(Split("" & .Fields(10).Value, ",")) + 1
                    End If
                '若為發文件數
                Else
                    '若商品類別有值
                    If "" & .Fields(10).Value <> "" Then
                        strTemp(5) = UBound(Split("" & .Fields(10).Value, ",")) + 1
                    End If
                End If
            End If
            'End
            '欄位--部門別名稱, 目標件數
            strSql = "INSERT INTO R020411 VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",''," & Val(strTemp(5)) & ",'','" & ChgSQL(strTemp(7)) & "','" & strUserNum & "','" & ChgSQL(strTemp(8)) & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
        'Add By Sindy 2025/9/2 檢查是否有"沒有"出現的S部門 ex:114/8 中四區沒有出現,才提出的需求
        strSql = "INSERT INTO R020411" & _
                 " select NVL(A0902,A0903),'','',0,'',0,'',A0901,'" & strUserNum & "','000'" & _
                 " from ACC090,staff" & _
                 " where substr(a0901,1,2) in('S1','S2')" & _
                 " and a0901=st03 and st04='1' and (substr(st01,1,1)>='6' and substr(st01,1,1)<'F')" & _
                 " and not exists(select * FROM R020411 WHERE ID='" & strUserNum & "' and R084001=NVL(A0902,A0903))" & _
                 " group by NVL(A0902,A0903),A0901"
        cnnConnection.Execute strSql, intI
        '2025/9/2 END
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
PrintData strDateFrom, strDateTo
'ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

'Modify By Cheng 2002/02/21
'Sub PrintData()
Sub PrintData(ByRef strDateFrom As String, ByRef strDateTo As String)
Dim strTitle As String
'Modify By Cheng 2002/02/21
'strSQL = "select R084001,max(R084002),MAX(R084003),SUM(R084004),'',SUM(R084006),'',R084008 FROM R020411 WHERE ID='" & strUserNum & "' GROUP BY R084008,R084001"
'欄位--部門名稱, 目標件數, 國內收文件數, 紅/綠, 國內發文件數, 紅/綠, 大陸收文件數, 大陸發文件數
'strSQL = "select R084001,max(R084002),SUM(R084004),'',SUM(R084006),'','','',R084008,R084009 FROM R020411 WHERE ID='" & strUserNum & "' GROUP BY R084008,R084001,R084009"
strSql = "select R084001,max(R084002),SUM(R084004),'',SUM(R084006),'','','',R084008,R084009 FROM R020411 WHERE R084008<>'000' AND ID='" & strUserNum & "' GROUP BY R084008,R084001,R084009"
CheckOC

strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
'        strTemp3 = CheckStr(.Fields(7))
        strTemp3 = CheckStr(.Fields(8))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = ""
            Next i
            'Add By Cheng 2003/03/17
            '業務區別
            strTemp(0) = CheckStr(.Fields(0))
            If .Fields(9) = "000" Then
               'Modify By Cheng 2002/02/21
               For i = 0 To 9
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               .MoveNext
               If .EOF Then GoTo PrintData
            End If
            If .Fields(9) <> "000" Then
               strTemp(6) = CheckStr(.Fields(2))
               strTemp(7) = CheckStr(.Fields(4))
                'Add By Cheng 2002/11/27
                If strTemp(8) = "" Then strTemp(8) = "" & .Fields(8).Value
            Else
               .MovePrevious
            End If
'            If Val(txt1(3)) = 15 Then
'                If Val(strTemp(3)) > Val(strTemp(2)) Then
'                    strTemp(4) = "綠"
'                Else
'                    strTemp(4) = "紅"
'                End If
'                If Val(strTemp(5)) >= Val(strTemp(2)) Then
'                    strTemp(6) = "綠"
'                Else
'                    strTemp(6) = "紅"
'                End If
'            Else
PrintData:
                 'Modify By Cheng 2002/04/11
'                If Val(strTemp(2)) >= Val(strTemp(1)) Then
'                    strTemp(3) = "綠"
'                Else
'                    strTemp(3) = "紅"
'                End If
'                If Val(strTemp(4)) >= Val(strTemp(1)) Then
'                    strTemp(5) = "綠"
'                Else
'                    strTemp(5) = "紅"
'                End If
'            End If
            If StrToStr(strTemp3, 1) <> StrToStr(strTemp(8), 1) Then
                ShowLine
                If strTemp3 <> "000" Then
                  PrintEnd (IIf(Left(strTemp3, 2) = "S1", 3, IIf(Left(strTemp3, 2) = "S2", 4, IIf(Left(strTemp3, 2) = "S3", 5, 6))))
                  m_dblSubTotal = 0
                  ShowLine
                End If
                strTemp3 = strTemp(8)
                '
                'PrintTitle
            End If
            strTemp(0) = StrToStr(strTemp(0), 4)
            'Add By Cheng 2002/04/09
            '取得目標值
            'Modify By Sindy 2010/10/15
            'strTemp(1) = GetPerformanceByNick_1(Val(Mid(strDateFrom, 5, 2)), Val(Mid(strDateTo, 5, 2)), "T", strTemp(8))
            strTemp(1) = GetPerformanceByNick_1(Val(Left(strDateFrom, 6)), Val(Left(strDateTo, 6)), "T", strTemp(8))
            '2010/10/15 End
            '若非整月, 則要換算目標件數
            If Len("" & Me.txt1(2).Text) > 0 Then
               strTemp(1) = strTemp(1) * (GetWorkDay(strDateTo, strDateFrom)) / (GetWorkDay(Mid(strDateTo, 1, 6) & "31", Mid(strDateFrom, 1, 6) & "01"))
            End If
            'Modify By Cheng 2003/04/01
            If Val(strTemp(1)) <> 0 Then
                'Add By Cheng 2002/04/11
                If Val(strTemp(2)) >= Val(strTemp(1)) Then
                    strTemp(3) = "綠"
                Else
                    strTemp(3) = "紅"
                End If
                If Val(strTemp(4)) >= Val(strTemp(1)) Then
                    strTemp(5) = "綠"
                Else
                    strTemp(5) = "紅"
                End If
            Else
                strTemp(3) = ""
                strTemp(5) = ""
            End If
            
            PrintDatil
            m_dblSubTotal = m_dblSubTotal + strTemp(1)
            m_dblTotal = m_dblTotal + strTemp(1)
            ShowLine
            If iPrint >= 14000 Then
                ShowLine
                Page = Page + 1
                Printer.NewPage
                iPrint = 0
                PrintTitle
            End If
            'Add By Cheng 2002/02/21
            If .EOF Then Exit Do
            .MoveNext
        Loop
    End If
End With
ShowLine
'Modify By Cheng 2002/02/20
'PrintEnd (0)
PrintEnd (IIf(Left(strTemp3, 2) = "S1", 3, IIf(Left(strTemp3, 2) = "S2", 4, IIf(Left(strTemp3, 2) = "S3", 5, 6))))
ShowLine
PrintEnd (1) '全所總計
ShowLine
'Add By Cheng 2002/02/21
If m_Parts = 1 Then
   strDateFrom = DateAdd("m", -1, ChangeWStringToWDateString(strDateFrom))
   strDateFrom = Format(Year(strDateFrom), "0000") & Format(Month(strDateFrom), "00") & Format(Day(strDateFrom), "00")
   '若有指定日區間
   If Len("" & Me.txt1(2).Text) > 0 Then
      strDateTo = DateAdd("m", -1, ChangeWStringToWDateString(strDateTo))
      strDateTo = Format(Year(strDateTo), "0000") & Format(Month(strDateTo), "00") & Format(Day(strDateTo), "00")
   '若未指定日區間
   Else
      strDateTo = DateAdd("m", -1, ChangeWStringToWDateString(strDateTo))
      strDateTo = Format(Year(strDateTo), "0000") & Format(Month(strDateTo), "00") & Format(PUB_GetMonthDays(Year(strDateTo), Month(strDateTo)), "00")
   End If
   Process1 strDateFrom, strDateTo '處理上個月同期資料
   '若為整個月
   If Len("" & Me.txt1(2).Text) <= 0 Then
      strTitle = (Mid(strDateFrom, 1, 4) - 1911) & "年" & Mid(strDateFrom, 5, 2) & "月"
   '若有指定日區間
   Else
      strTitle = (Mid(strDateFrom, 1, 4) - 1911) & "年" & Format(Mid(strDateFrom, 5, 2), "00") & "月" & Format(Me.txt1(2).Text, "00") & "日－" & Format(Mid(strDateFrom, 5, 2), "00") & "月" & Format(Me.txt1(3).Text, "00") & "日"
   End If
Else
   strDateFrom = DateAdd("yyyy", -1, ChangeWStringToWDateString(strDateFrom))
   strDateFrom = Format(Year(strDateFrom), "0000") & Format(Month(strDateFrom), "00") & Format(Day(strDateFrom), "00")
   strDateTo = strDateTo - 10000
   strDateTo = Format(Mid(strDateTo, 1, 4), "0000") & Format(Mid(strDateTo, 5, 2), "00") & Format(PUB_GetMonthDays(Mid(strDateTo, 1, 4), Mid(strDateTo, 5, 2)), "00")
   Process1 strDateFrom, strDateTo '處理上個月同期資料
   strTitle = (Mid(strDateFrom, 1, 4) - 1911) & "年1－" & Val(Mid(strDateTo, 5, 2)) & "月"
End If
'PrintEnd (2) '上個月全所總計
PrintEnd_1 2, strTitle  '上個月全所總計
ShowLine
'Modify By Cheng 2002/02/22
'Printer.EndDoc

End Sub

Sub PrintEnd(Strindex As Integer)
'Add By Cheng 2002/02/21
Dim Rs As New ADODB.Recordset

Select Case Strindex
Case 0
'     strSQL = "select '小計',max(R084002),MAX(R084003),SUM(R084004),'',sum(R084006),'' from r020411 where id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
Case 1 '全所(含非智權人員)
      'Modify By Cheng 2002/02/21
'     strSQL = "select '全所(含非智權人員)',max(R084002),MAX(R084003),SUM(R084004),'',sum(R084006),'' from r020411 where id='" & strUserNum & "' "
     strSql = "select '全所(含非智權人員)'," & m_dblTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where R084009='000' AND id='" & strUserNum & "' "
     strSQL1 = "select '全所(含非智權人員)'," & m_dblTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where R084009='020' AND id='" & strUserNum & "' "
Case 2 '同期
      'Modify By Cheng 2002/02/21
'     strSQL = "select '上個月全所總計',max(R084002),MAX(R084003),SUM(R084004),'',sum(R084006),'' from r020411 where id='" & strUserNum & "' "
     strSql = "select '上個月全所總計','',SUM(R084004),'',sum(R084006),'' from r020411 where R084009='000' AND id='" & strUserNum & "' "
     strSQL1 = "select '上個月全所總計','',SUM(R084004),'',sum(R084006),'' from r020411 where R084009='020' AND id='" & strUserNum & "' "

'Add By Cheng 2002/02/20
Case 3 '北區
     strSql = "select '北區合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='000' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
     strSQL1 = "select '北區合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='020' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
Case 4 '中區
     strSql = "select '中區合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='000' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
     strSQL1 = "select '中區合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='020' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
Case 5 '南所
     strSql = "select '南所合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='000' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
     strSQL1 = "select '南所合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='020' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
Case 6 '高所
     strSql = "select '高所合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='000' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
     strSQL1 = "select '高所合計'," & m_dblSubTotal & ",SUM(R084004),'',sum(R084006),'' from r020411 where r084009='020' AND id='" & strUserNum & "' AND substr(R084008,1,2)='" & StrToStr(strTemp3, 1) & "' "
Case Else
     Exit Sub
End Select
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 7
                StrTemp7(i) = ""
            Next i
'            For i = 0 To 6
            For i = 0 To 5
                StrTemp7(i) = CheckStr(.Fields(i))
                If Len(StrTemp7(i)) = 0 Then
                  StrTemp7(i) = "0"
                End If
            Next i
            If Rs.State <> adStateClosed Then Rs.Close
            Set Rs = Nothing
            Rs.CursorLocation = adUseClient
            Rs.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
            If Rs.RecordCount > 0 Then
               StrTemp7(6) = "" & Rs.Fields(2).Value
               StrTemp7(7) = "" & Rs.Fields(4).Value
            End If
            If Rs.State <> adStateClosed Then Rs.Close
            Set Rs = Nothing

'            If Val(txt1(3)) = 15 Then
'                If Val(StrTemp7(3)) > Val(StrTemp7(2)) Then
'                    StrTemp7(4) = "綠"
'                Else
'                    StrTemp7(4) = "紅"
'                End If
'                If Val(StrTemp7(5)) >= Val(StrTemp7(2)) Then
'                    StrTemp7(6) = "綠"
'                Else
'                    StrTemp7(6) = "紅"
'                End If
'            Else
            If Strindex <> 2 Then
                If Val(StrTemp7(2)) >= Val(StrTemp7(1)) Then
                    StrTemp7(3) = "綠"
                Else
                    StrTemp7(3) = "紅"
                End If
                If Val(StrTemp7(4)) >= Val(StrTemp7(1)) Then
                    StrTemp7(5) = "綠"
                Else
                    StrTemp7(5) = "紅"
                End If
            End If
'            End If
                        
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            'Modify By Cheng 2003/03/03
'            Printer.CurrentX = PLeft(1) + 600 - Printer.TextWidth(Format(StrTemp7(1), "####0"))
            Printer.CurrentX = PLeft(1) + 600 - Printer.TextWidth(Format(Val("0" & StrTemp7(1)), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val("0" & StrTemp7(1)), "####0")
            'Modify By Cheng 2003/03/03
'            Printer.CurrentX = PLeft(2) + 600 - Printer.TextWidth(Format(StrTemp7(2), "####0"))
            Printer.CurrentX = PLeft(2) + 600 - Printer.TextWidth(Format(Val("0" & StrTemp7(2)), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val("0" & StrTemp7(2)), "####0")
            Printer.CurrentX = PLeft(3) + 100
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(3)
            'Modify By Cheng 2003/03/03
'            Printer.CurrentX = PLeft(4) + 600 - Printer.TextWidth(Format(StrTemp7(4), "####0"))
            Printer.CurrentX = PLeft(4) + 600 - Printer.TextWidth(Format(Val("0" & StrTemp7(4)), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val("0" & StrTemp7(4)), "####0")
            Printer.CurrentX = PLeft(5) + 100
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(5)
            'Modify By Cheng 2003/03/03
'            Printer.CurrentX = PLeft(6) + 600 - Printer.TextWidth(Format(StrTemp7(6), "####0"))
            Printer.CurrentX = PLeft(6) + 600 - Printer.TextWidth(Format(Val("0" & StrTemp7(6)), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val("0" & StrTemp7(6)), "####0")
            'Modify By Cheng 2003/03
'            Printer.CurrentX = PLeft(7) + 600 - Printer.TextWidth(Format(StrTemp7(7), "####0"))
            Printer.CurrentX = PLeft(7) + 600 - Printer.TextWidth(Format(Val("0" & StrTemp7(7)), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val("0" & StrTemp7(7)), "####0")
            Printer.CurrentX = PLeft(6) - 200
            Printer.CurrentY = iPrint
            Printer.Line (PLeft(6) - 200, iPrint)-(PLeft(6) - 200, iPrint + 300)
            
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                iPrint = 0
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintEnd_1(Strindex As Integer, strTitle As String)
'Add By Cheng 2002/02/21
Dim Rs As New ADODB.Recordset

Select Case Strindex
Case 2 '同期
      'Modify By Cheng 2002/02/21
'     strSQL = "select '上個月全所總計',max(R084002),MAX(R084003),SUM(R084004),'',sum(R084006),'' from r020411 where id='" & strUserNum & "' "
     strSql = "select '" & strTitle & "','',SUM(R084004),'',sum(R084006),'' from r020411 where R084009='000' AND id='" & strUserNum & "' "
     strSQL1 = "select '" & strTitle & "','',SUM(R084004),'',sum(R084006),'' from r020411 where R084009='020' AND id='" & strUserNum & "' "
Case Else
     Exit Sub
End Select
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 7
                StrTemp7(i) = ""
            Next i
'            For i = 0 To 6
            For i = 0 To 5
                StrTemp7(i) = CheckStr(.Fields(i))
                If Len(StrTemp7(i)) = 0 Then
                  StrTemp7(i) = "0"
                End If
            Next i
            If Rs.State <> adStateClosed Then Rs.Close
            Set Rs = Nothing
            Rs.CursorLocation = adUseClient
            Rs.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
            If Rs.RecordCount > 0 Then
               StrTemp7(6) = "" & Rs.Fields(2).Value
               StrTemp7(7) = "" & Rs.Fields(4).Value
            End If
            If Rs.State <> adStateClosed Then Rs.Close
            Set Rs = Nothing
                        
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1) + 600 - Printer.TextWidth(Format(StrTemp7(1), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print "" 'Format(StrTemp7(1), "####0")
            Printer.CurrentX = PLeft(2) + 600 - Printer.TextWidth(Format(StrTemp7(2), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(2), "####0")
            Printer.CurrentX = PLeft(3) + 100
            Printer.CurrentY = iPrint
            Printer.Print "" 'StrTemp7(3)
            Printer.CurrentX = PLeft(4) + 600 - Printer.TextWidth(Format(StrTemp7(4), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(4), "####0")
            Printer.CurrentX = PLeft(5) + 100
            Printer.CurrentY = iPrint
            Printer.Print "" 'StrTemp7(5)
            Printer.CurrentX = PLeft(6) + 600 - Printer.TextWidth(Format(StrTemp7(6), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(6), "####0")
            Printer.CurrentX = PLeft(7) + 600 - Printer.TextWidth(Format(StrTemp7(7), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(7), "####0")
            Printer.CurrentX = PLeft(6) - 200
            Printer.CurrentY = iPrint
            Printer.Line (PLeft(6) - 200, iPrint)-(PLeft(6) - 200, iPrint + 300)
            
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                iPrint = 0
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintTitle()
GetPleft
'Modify By Cheng 2002/02/22
'iPrint = 0
If m_Parts = 1 Then Printer.Orientation = 1
Printer.Font.Name = "細明體"
Printer.Font.Size = 12
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 4200
Printer.CurrentY = iPrint
If m_Parts = 1 Then Printer.Print GetTitleNick & "各區收/發文達成比較表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = 400
Printer.CurrentY = iPrint
If m_Parts = 1 Then Printer.Print "列印人：" & strUserName
'Printer.CurrentX = 8000
'Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 400
Printer.CurrentY = iPrint
'若列印上表
If m_Parts = 1 Then
   '若列印整個月
   If Len("" & Me.txt1(2).Text) <= 0 Then
      'Modify By Sindy 2011/2/1
      If Len(Me.txt1(1).Text) = 5 Then
         Printer.Print Mid(Me.txt1(1).Text, 1, 3) & "年" & Mid(Me.txt1(1).Text, 4, 2) & "月份各區商申及大陸商標收/發文情形"
      '2011/2/1 End
      Else
         Printer.Print Mid(Me.txt1(1).Text, 1, 2) & "年" & Mid(Me.txt1(1).Text, 3, 2) & "月份各區商申及大陸商標收/發文情形"
      End If
   '若列印日區間
   Else
      'Modify By Sindy 2011/2/1
      If Len(Me.txt1(1).Text) = 5 Then
         Printer.Print Mid(Me.txt1(1).Text, 1, 3) & "年" & Mid(Me.txt1(1).Text, 4, 2) & "月" & Me.txt1(2).Text & "日－" & Mid(Me.txt1(1).Text, 4, 2) & "月" & Me.txt1(3).Text & "日各區商申及大陸商標收/發文情形"
      '2011/2/1 End
      Else
         Printer.Print Mid(Me.txt1(1).Text, 1, 2) & "年" & Mid(Me.txt1(1).Text, 3, 2) & "月" & Me.txt1(2).Text & "日－" & Mid(Me.txt1(1).Text, 3, 2) & "月" & Me.txt1(3).Text & "日各區商申及大陸商標收/發文情形"
      End If
   End If
'若列印下表
Else
   '列印月份區間
   If Len("" & Me.txt1(2).Text) <= 0 Then
      'Modify By Cheng 2002/04/09
'      Printer.Print Mid(Me.txt1(1).Text, 1, 2) - 1 & "年1月－" & Val(Mid(Me.txt1(1).Text, 3, 2)) & "月份各區商申及大陸商標收/發文情形"
      'Modify By Sindy 2011/2/1
      If Len(Me.txt1(1).Text) = 5 Then
         Printer.Print Mid(Me.txt1(1).Text, 1, 3) & "年1月－" & Val(Mid(Me.txt1(1).Text, 4, 2)) & "月份各區商申及大陸商標收/發文情形"
      Else
         Printer.Print Mid(Me.txt1(1).Text, 1, 2) & "年1月－" & Val(Mid(Me.txt1(1).Text, 3, 2)) & "月份各區商申及大陸商標收/發文情形"
      End If
   End If
End If
Printer.CurrentX = 8400
Printer.CurrentY = iPrint
'Printer.Print "頁    次：" & str(Page)
Printer.CurrentX = 9000 '9400
Printer.CurrentY = iPrint
If m_Parts = 1 Then Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 400
Printer.CurrentY = iPrint
Printer.Line (300, iPrint - 30)-(11600, iPrint - 30)
'iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    iPrint = 0
    PrintTitle
    Exit Sub
End If
'Add By Cheng 2002/02/21
Printer.CurrentX = PLeft(4) - 800
Printer.CurrentY = iPrint
Printer.Print "國內"
Printer.CurrentX = PLeft(7) - 500
Printer.CurrentY = iPrint
Printer.Print "大陸"
Printer.CurrentX = PLeft(6) - 200
Printer.CurrentY = iPrint
Printer.Line (PLeft(6) - 200, iPrint)-(PLeft(6) - 200, iPrint + 300)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Line (300, iPrint - 30)-(11600, iPrint - 30)
'iPrint = iPrint + 300

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'Modify By Sindy 2012/12/24
'Printer.Print "目標件數"
Printer.Print "目標(類)"
'2012/12/24 End
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
'Modify By Sindy 2012/12/24
'Printer.Print "收文件數"
Printer.Print "收文(類)"
'2012/12/24 End
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print ""
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
'Modify By Sindy 2012/12/24
'Printer.Print "發文件數"
Printer.Print "發文(類)"
'2012/12/24 End
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print ""
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "收文件數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "發文件數"

Printer.CurrentX = PLeft(6) - 200
Printer.CurrentY = iPrint
Printer.Line (PLeft(6) - 200, iPrint)-(PLeft(6) - 200, iPrint + 300)

iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    iPrint = 0
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 400
Printer.CurrentY = iPrint
'Modify By Cheng 2002/02/20
'Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
Printer.Line (300, iPrint - 30)-(11600, iPrint - 30)
'iPrint = iPrint + 220
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    iPrint = 0
    PrintTitle
    Exit Sub
End If
End Sub

Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
'Modify By Cheng 2003/03/03
'Printer.CurrentX = PLeft(1) + 600 - Printer.TextWidth(Format(strTemp(1), "####0"))
Printer.CurrentX = PLeft(1) + 600 - Printer.TextWidth(Format(Val("0" & strTemp(1)), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(Val("0" & strTemp(1)), "####0")
'Modify By Cheng 2003/03/03
'Printer.CurrentX = PLeft(2) + 600 - Printer.TextWidth(Format(strTemp(2), "####0"))
Printer.CurrentX = PLeft(2) + 600 - Printer.TextWidth(Format(Val("0" & strTemp(2)), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(Val("0" & strTemp(2)), "####0")
Printer.CurrentX = PLeft(3) + 100
Printer.CurrentY = iPrint
Printer.Print strTemp(3)
'Modify By Cheng 2003/03/03
'Printer.CurrentX = PLeft(4) + 600 - Printer.TextWidth(Format(strTemp(4), "####0"))
Printer.CurrentX = PLeft(4) + 600 - Printer.TextWidth(Format(Val("0" & strTemp(4)), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(Val("0" & strTemp(4)), "####0")
Printer.CurrentX = PLeft(5) + 100
Printer.CurrentY = iPrint
Printer.Print strTemp(5)
'Modify By Cheng 2003/03/03
'Printer.CurrentX = PLeft(6) + 600 - Printer.TextWidth(Format(strTemp(6), "####0"))
Printer.CurrentX = PLeft(6) + 600 - Printer.TextWidth(Format(Val("0" & strTemp(6)), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(Val("0" & strTemp(6)), "####0")
'Modify By Cheng 2003/03/03
'Printer.CurrentX = PLeft(7) + 600 - Printer.TextWidth(Format(strTemp(7), "####0"))
Printer.CurrentX = PLeft(7) + 600 - Printer.TextWidth(Format(Val("0" & strTemp(7)), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(Val("0" & strTemp(7)), "####0")
Printer.CurrentX = PLeft(6) - 200
Printer.CurrentY = iPrint
Printer.Line (PLeft(6) - 200, iPrint)-(PLeft(6) - 200, iPrint + 300)
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 400
PLeft(1) = 2400
PLeft(2) = 3900
PLeft(3) = 5100
PLeft(4) = 6400
PLeft(5) = 7600
PLeft(6) = 8900
PLeft(7) = 10400
End Sub

Sub ShowLine()
Printer.CurrentX = 400
Printer.CurrentY = iPrint
'Modify By Cheng 2002/02/20
'Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
Printer.Line (300, iPrint - 30)-(11600, iPrint - 30)
'iPrint = iPrint + 220
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    iPrint = 0
    PrintTitle
End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'txt1(0) = GetSystemKindByNickT
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2015/7/3
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set Printer = Printers(SeekPrint)
'Printer.Orientation = SeekPrintL
   'Add By Sindy 2015/7/3
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2015/7/3 END
   
   Set frm020411 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Remove by Morgan 2007/4/14 點確定再檢查就好，否則會無法離開
   'CheckKeyIn Index
End Sub

Private Function CheckKeyIn(Index As Integer) As Integer
CheckKeyIn = -1
Select Case Index
Case 1 '收/發文年月
   If Len("" & Me.txt1(Index).Text) <= 0 Then
      s = MsgBox("未輸入收/發文年月!!!", , "USER 輸入錯誤")
      txt1(1).SetFocus
      txt1(1).SelStart = 0
      txt1(1).SelLength = Len(txt1(1))
      Exit Function
   Else
      If IsDate(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01"))) = False Then
         s = MsgBox("收/發文年月輸入錯誤!!!", , "USER 輸入錯誤")
         txt1(1).SetFocus
         txt1(1).SelStart = 0
         txt1(1).SelLength = Len(txt1(1))
         Exit Function
      End If
   End If
Case 2 '收發文日區間(起)
   If Len(txt1(1)) <> 0 And Len(txt1(2)) <> 0 Then
      If IsDate(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & Format(txt1(2), "00")))) Then
      
      Else
          s = MsgBox("收發文日錯誤, " & txt1(1) & " 沒有 " & txt1(2) & "這天!!", , "USER 輸入錯誤")
          txt1(2).SetFocus
          txt1(2).SelStart = 0
          txt1(2).SelLength = Len(txt1(2))
          Exit Function
      End If
   End If
Case 3 '收發文日區間(迄)
   If Len(txt1(1)) <> 0 And Len(txt1(3)) <> 0 Then
      If IsDate(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & Format(txt1(3), "00")))) Then
      
      Else
          s = MsgBox("收發文日錯誤, " & txt1(1) & " 沒有 " & txt1(3) & "這天!!", , "USER 輸入錯誤")
          txt1(3).SetFocus
          txt1(3).SelStart = 0
          txt1(3).SelLength = Len(txt1(3))
          Exit Function
      End If
   End If
   If RunNick(txt1(Index - 1), txt1(Index)) Then
       txt1(Index - 1).SetFocus
       txt1_GotFocus (Index - 1)
       Exit Function
   End If
Case Else
'Case 2, 3
'   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
'      Me.txt1(Index).SetFocus
'      txt1_GotFocus Index
'      Exit Function
'   End If
'   If Index = 3 Then
'     If RunNick(txt1(Index - 1), txt1(Index)) Then
'         txt1(Index - 1).SetFocus
'         txt1_GotFocus (Index - 1)
'         Exit Function
'      End If
'    End If
'Case Else
End Select
CheckKeyIn = 0
End Function
