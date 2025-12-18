VERSION 5.00
Begin VB.Form frm050311 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人准駁明細表"
   ClientHeight    =   2640
   ClientLeft      =   4500
   ClientTop       =   2400
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3105
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "Y"
      Top             =   1980
      Width           =   255
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2055
      TabIndex        =   10
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1260
      TabIndex        =   9
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1065
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2316
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2250
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1680
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   945
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1680
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   945
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1380
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2265
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   945
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   945
      MaxLength       =   1
      TabIndex        =   1
      Top             =   780
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   945
      TabIndex        =   0
      Top             =   480
      Width           =   2040
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "(Y列印)"
      Height          =   180
      Left            =   1665
      TabIndex        =   21
      Top             =   2010
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否列印明細："
      Height          =   180
      Left            =   90
      TabIndex        =   20
      Top             =   2010
      Width           =   1260
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   180
      Left            =   1845
      TabIndex        =   19
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1. 准 2. 駁 )"
      Height          =   180
      Left            =   1260
      TabIndex        =   18
      Top             =   840
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(1. 承辦人 2. 准駁日)"
      Height          =   180
      Left            =   1380
      TabIndex        =   17
      Top             =   2340
      Width           =   1605
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "列印順序："
      Height          =   180
      Left            =   90
      TabIndex        =   16
      Top             =   2340
      Width           =   900
   End
   Begin VB.Line Line4 
      X1              =   1905
      X2              =   2145
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   90
      TabIndex        =   15
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   1410
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   1905
      X2              =   2145
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "准駁日期："
      Height          =   180
      Left            =   90
      TabIndex        =   13
      Top             =   1110
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "准駁代碼："
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   510
      Width           =   900
   End
End
Attribute VB_Name = "frm050311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, i As Integer, j As Integer, s As Integer, k As Integer
Dim strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String
Dim Page As Integer, iPrint As Integer, PLeft(0 To 8) As Integer, strTemp(0 To 8) As String
Dim StrTemp6 As String, StrTemp4(0 To 4) As String, StrTemp5(0 To 5) As String, St As String, StrTemp7(0 To 5) As String
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
   'Add By Cheng 2002/09/16
   blnClkSure = False
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1(0).SelStart = 0
        txt1(0).SelLength = Len(txt1(0))
        Exit Sub
     Else
        If Len(txt1(1)) = 0 Then
            s = MsgBox("准駁代碼不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1(1).SelStart = 0
            txt1(1).SelLength = Len(txt1(1))
            Exit Sub
        Else
            If Len(txt1(3)) = 0 Then
                s = MsgBox("准駁日期不可空白!!", , "USER 輸入錯誤")
                txt1(2).SetFocus: txt1(2).SelStart = 0: txt1(2).SelLength = Len(txt1(2))
                Exit Sub
            Else
               'Add By Cheng 2002/03/20
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
               'Add By Cheng 2002/09/16
               If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
                  If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
                     MsgBox "准駁日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(2).SetFocus
                     txt1_GotFocus 2
                     Exit Sub
                   End If
               End If
               If txt1(4) <> "" Then
                  lbl1 = GetPrjSales(txt1(4))
                  If Me.txt1(4).Text <> "" Then
                     If Me.txt1(4).Text = Me.lbl1.Caption Then
                        Me.lbl1.Caption = ""
                        Me.txt1(4).SetFocus
                        txt1_GotFocus 4
                        Exit Sub
                     End If
                  End If
               End If
               If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
                  If Me.txt1(5).Text > Me.txt1(6).Text Then
                     MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(5).SetFocus
                     txt1_GotFocus 5
                     Exit Sub
                   End If
               End If
               If Me.txt1(8).Text = "" Then
                  MsgBox "請輸入列印順序!!!", vbExclamation + vbOKOnly
                  Me.txt1(8).SetFocus
                  txt1_GotFocus 8
                  Exit Sub
               End If
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
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

Sub Process()            '處理主程式
ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/3 清除查詢印表記錄檔欄位
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050311 WHERE ID='" & strUserNum & "' "
'系統類別
strSQL1 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " and PA01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/12/3
End If
'准駁日期
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 + " AND PA20>=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(3))) <> 0 Then
   strSQL1 = strSQL1 & " AND PA20<=" & Val(ChangeTStringToWString(txt1(3))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/3
End If
'承辦人
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP14='" & txt1(4) & "' "
    pub_QL05 = pub_QL05 & ";" & Label4 & txt1(4) 'Add By Sindy 2010/12/3
End If
'申請國家
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label5 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/3
End If
If txt1(7) = "Y" Or txt1(7) = "y" Then
   pub_QL05 = pub_QL05 & ";" & Label10 & txt1(7) 'Add By Sindy 2010/12/3
End If
'組合
'91.8.9 modify by sonia
'strSQL = "SELECT ST02," & SQLDate("PA20") & ",PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS B," & SQLDate("CP27") & ",NVL(PA05,NVL(PA06,PA07)),DECODE(PA09,'000',PTM03,PTM04),NVL(NA03,NA04),DECODE(PA16,'1','准','2','駁',''),DECODE(TO_NUMBER(CP10),107,'再審','初審'),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF,NATION,PATENTTRADEMARKMAP WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND pa09=na01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=ST01(+)  " & strSQL
'strSQL = strSQL + " ORDER BY CP14," & SQLDate("PA20") & ",B "
'92.03.27 nick add left join
'strSQL = "SELECT CP14,PA20,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS B," & SQLDate("CP27") & ",NVL(PA05,NVL(PA06,PA07)),PTM03,NVL(NA03,NA04),PA16,DECODE(TO_NUMBER(CP10),107,'再審','初審'),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF,NATION,PATENTTRADEMARKMAP WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND pa09=na01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=ST01(+) AND PA16=CP24 AND PA20=CP25 " & strSQL1
strSql = "SELECT CP14,PA20,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS B," & SQLDate("CP27") & ",NVL(PA05,NVL(PA06,PA07)),PTM03,NVL(NA03,NA04),PA16,DECODE(TO_NUMBER(CP10),107,'再審','初審'),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF,NATION,PATENTTRADEMARKMAP WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND pa09=na01(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=ST01(+) AND pa16=cp24(+) AND pa20=CP25(+) " & strSQL1
'91.8.9 end

CheckOC
k = 0
cnnConnection.Execute " INSERT INTO R050311 " & strSql
strSql = "SELECT * FROM R050311 WHERE ID='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/3
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/3
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
'是否列印明細
Page = 1
If txt1(7) = "Y" Or txt1(7) = "y" Then
    PrintData
Else
    PrintData1
End If
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

'91.8.9 modify by sonia
Sub PrintData()      '印明細
'Add By Cheng 2002/08/20
Dim strField0 As String

'Add By Cheng 2002/08/20
strField0 = ""
strSQL2 = ""
'准駁代碼
If Len(txt1(1)) <> 0 Then
    strSQL2 = strSQL2 + " AND R010008='" & txt1(1) & "' "
End If
PrintTitle
'列印順序
If txt1(8) = "1" Then
   strSql = "SELECT DISTINCT R010001,ST02 FROM R050311,STAFF WHERE ID='" & strUserNum & "'" & strSQL2 & " AND R010001=ST01(+)"
Else
   strSql = "SELECT DISTINCT R010002," & SQLDate("R010002") & " FROM R050311 WHERE ID='" & strUserNum & "'" & strSQL2
End If

CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    adoRecordset1.MoveFirst
    Do While adoRecordset1.EOF = False
        '列印順序
        If txt1(8) = "1" Then
            strSql = "select ST02," & SQLDate("R010002") & ",R010003,R010004,R010005,R010006,R010007,DECODE(R010008,'1','准','2','駁',''),R010009,ID from R050311,STAFF WHERE R010001='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL2 & " AND ID='" & strUserNum & "' AND R010001=ST01(+) ORDER BY R010001,R010002,R010003"
        Else
            strSql = "select " & SQLDate("R010002") & ",ST02,R010003,R010004,R010005,R010006,R010007,DECODE(R010008,'1','准','2','駁',''),R010009,ID from R050311,STAFF WHERE R010002='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL2 & " AND ID='" & strUserNum & "' AND R010001=ST01(+) ORDER BY R010002,R010001,R010003"
        End If
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            With adoRecordset
                .MoveFirst
                Do While .EOF = False
                    For i = 0 To 8
                        strTemp(i) = CheckStr(.Fields(i))
                    Next i
                    strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 32), vbUnicode)
                    strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 8), vbUnicode)
                     If strField0 <> strTemp(0) Then
                        strField0 = strTemp(0)
                     Else
                        strTemp(0) = ""
                     End If
                    PrintDatil
                    .MoveNext
                     'Add By Cheng 2002/08/20
                     If Not .EOF Then
                        If iPrint > 10000 Then
                           strTemp(0) = ""
                            PrintEnd
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle
                        End If
                     End If
                Loop
            End With
        End If
        CheckOC
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        StrTemp4(1) = "0.00"
        StrTemp4(2) = "0"
        StrTemp4(3) = "0"
    
        If txt1(8) = "1" Then
            '92.04.04 nick add left join
            'strSQL = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24 AND PA20=CP25 AND CP14='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL1 & " GROUP BY PA16"
            strSql = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24(+) AND PA20=CP25(+) AND CP14='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL1 & " GROUP BY PA16"
        Else
            '92.04.04 nick add left join
            'strSQL = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24 AND PA20=CP25 AND PA20='" & adoRecordset1.Fields(0) & "'" & strSQL1 & " GROUP BY PA16"
            strSql = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24(+) AND PA20=CP25(+) AND PA20='" & adoRecordset1.Fields(0) & "'" & strSQL1 & " GROUP BY PA16"
        End If
        
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                StrTemp5(2) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(3) = CheckStr(adoRecordset.Fields(1))
                If StrTemp5(2) = "1" Then
                    StrTemp4(2) = StrTemp5(3)
                Else
                    If StrTemp5(2) = "2" Then
                        StrTemp4(3) = StrTemp5(3)
                    End If
                End If
                adoRecordset.MoveNext
            Loop
        End If
        CheckOC
        'Modify By Cheng 2003/04/04
'        StrTemp4(0) = adoRecordset1.Fields(1)
        StrTemp4(0) = "" & adoRecordset1.Fields(1)
        StrTemp4(2) = Trim(str(Val(StrTemp4(2))))
        StrTemp4(3) = Trim(str(Val(StrTemp4(3))))
        StrTemp4(1) = Trim(str(Val(StrTemp4(2)) + Val(StrTemp4(3))))
        If Val(StrTemp4(1)) <> 0 Then
'            StrTemp4(1) = Format(Trim(str(Val(StrTemp4(2)) / Val(StrTemp4(1)) * 100)), "##.##")
            StrTemp4(1) = Format(Trim(str(Val(StrTemp4(2)) / Val(StrTemp4(1)) * 100)), "#0.00")
        Else
            StrTemp4(1) = "0.00"
        End If
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        PrintTotil
         'Add By Cheng 2002/08/20
        iPrint = iPrint + 300
        
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        adoRecordset1.MoveNext
    Loop
End If

CheckOC2
StrTemp4(0) = "ALL"
StrTemp4(2) = "0"
StrTemp4(3) = "0"
'92.04.04 nick add left join
'strSQL = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24 AND PA20=CP25 " & strSQL1 & "GROUP BY PA16"
strSql = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24(+) AND PA20=CP25(+) " & strSQL1 & "GROUP BY PA16"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        StrTemp5(2) = CheckStr(adoRecordset.Fields(0))
        StrTemp5(3) = CheckStr(adoRecordset.Fields(1))
        If StrTemp5(2) = "1" Then
            StrTemp4(2) = StrTemp5(3)
        Else
            If StrTemp5(2) = "2" Then
                StrTemp4(3) = StrTemp5(3)
            End If
        End If
        adoRecordset.MoveNext
    Loop
End If
CheckOC
'StrTemp4(0) = StrTemp5(0)
StrTemp4(2) = Trim(str(Val(StrTemp4(2))))
StrTemp4(3) = Trim(str(Val(StrTemp4(3))))
StrTemp4(1) = Trim(str(Val(StrTemp4(2)) + Val(StrTemp4(3))))
If Val(StrTemp4(1)) <> 0 Then
    StrTemp4(1) = Format(Trim(str(Val(StrTemp4(2)) / Val(StrTemp4(1)) * 100)), "##.##")
Else
    StrTemp4(1) = "0.00"
End If
PrintTotil
If iPrint > 10000 Then
    PrintEnd
    Printer.NewPage
    Page = Page + 1
    PrintTitle
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint > 10000 Then
    PrintEnd
    Printer.NewPage
    Page = Page + 1
    PrintTitle
End If
PrintEnd
Printer.EndDoc
End Sub

Sub PrintData1()        '不印明細
strSQL2 = ""
'准駁代碼
If Len(txt1(1)) <> 0 Then
    strSQL2 = strSQL2 + " AND R010008='" & txt1(1) & "' "
End If
PrintTitle
'列印順序
If txt1(8) = "1" Then
   strSql = "SELECT DISTINCT R010001,ST02 FROM R050311,STAFF WHERE ID='" & strUserNum & "'" & strSQL2 & " AND R010001=ST01(+)"
Else
   strSql = "SELECT DISTINCT R010002," & SQLDate("R010002") & " FROM R050311 WHERE ID='" & strUserNum & "'" & strSQL2
End If

CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    adoRecordset1.MoveFirst
    Do While adoRecordset1.EOF = False
        '列印順序
        If txt1(8) = "1" Then
            strSql = "select ST02," & SQLDate("R010002") & ",R010003,R010004,R010005,R010006,R010007,DECODE(R010008,'1','准','2','駁',''),R010009,ID from R050311,STAFF WHERE R010001='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL2 & " AND ID='" & strUserNum & "' AND R010001=ST01(+) ORDER BY R010001,R010002,R010003"
        Else
            strSql = "select " & SQLDate("R010002") & ",ST02,R010003,R010004,R010005,R010006,R010007,DECODE(R010008,'1','准','2','駁',''),R010009,ID from R050311,STAFF WHERE R010002='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL2 & " AND ID='" & strUserNum & "' AND R010001=ST01(+) ORDER BY R010002,R010001,R010003"
        End If
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            With adoRecordset
                .MoveFirst
                Do While .EOF = False
                    For i = 0 To 8
                        strTemp(i) = CheckStr(.Fields(i))
                    Next i
                    strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 32), vbUnicode)
                    strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 8), vbUnicode)
                    'PrintDatil  不印明細
                    .MoveNext
                Loop
            End With
        End If
        CheckOC
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        StrTemp4(1) = "0.00"
        StrTemp4(2) = "0"
        StrTemp4(3) = "0"
    
        If txt1(8) = "1" Then
            '92.04.04 nick add left join
            'strSQL = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24 AND PA20=CP25 AND CP14='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL1 & " GROUP BY PA16"
            strSql = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24(+) AND PA20=CP25(+) AND CP14='" & CheckStr(adoRecordset1.Fields(0)) & "'" & strSQL1 & " GROUP BY PA16"
        Else
            '92.04.04 nick add left join
            'strSQL = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24 AND PA20=CP25 AND PA20='" & adoRecordset1.Fields(0) & "'" & strSQL1 & " GROUP BY PA16"
            strSql = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24(+) AND PA20=CP25(+) AND PA20='" & adoRecordset1.Fields(0) & "'" & strSQL1 & " GROUP BY PA16"
        End If
        
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                StrTemp5(2) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(3) = CheckStr(adoRecordset.Fields(1))
                If StrTemp5(2) = "1" Then
                    StrTemp4(2) = StrTemp5(3)
                Else
                    If StrTemp5(2) = "2" Then
                        StrTemp4(3) = StrTemp5(3)
                    End If
                End If
                adoRecordset.MoveNext
            Loop
        End If
        CheckOC
        StrTemp4(0) = adoRecordset1.Fields(1)
        StrTemp4(2) = Trim(str(Val(StrTemp4(2))))
        StrTemp4(3) = Trim(str(Val(StrTemp4(3))))
        StrTemp4(1) = Trim(str(Val(StrTemp4(2)) + Val(StrTemp4(3))))
        If Val(StrTemp4(1)) <> 0 Then
            StrTemp4(1) = Format(Trim(str(Val(StrTemp4(2)) / Val(StrTemp4(1)) * 100)), "##.##")
        Else
            StrTemp4(1) = "0.00"
        End If
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        PrintTotil
        
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        adoRecordset1.MoveNext
    Loop
End If

CheckOC2
StrTemp4(0) = "ALL"
StrTemp4(2) = "0"
StrTemp4(3) = "0"
'92.04.04 nick add left join
'strSQL = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24 AND PA20=CP25 " & strSQL1 & "GROUP BY PA16"
strSql = "SELECT PA16,COUNT(PA16) FROM CASEPROGRESS,PATENT WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND PA16=CP24(+) AND PA20=CP25(+) " & strSQL1 & "GROUP BY PA16"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        StrTemp5(2) = CheckStr(adoRecordset.Fields(0))
        StrTemp5(3) = CheckStr(adoRecordset.Fields(1))
        If StrTemp5(2) = "1" Then
            StrTemp4(2) = StrTemp5(3)
        Else
            If StrTemp5(2) = "2" Then
                StrTemp4(3) = StrTemp5(3)
            End If
        End If
        adoRecordset.MoveNext
    Loop
End If
CheckOC
'StrTemp4(0) = StrTemp5(0)
StrTemp4(2) = Trim(str(Val(StrTemp4(2))))
StrTemp4(3) = Trim(str(Val(StrTemp4(3))))
StrTemp4(1) = Trim(str(Val(StrTemp4(2)) + Val(StrTemp4(3))))
If Val(StrTemp4(1)) <> 0 Then
    StrTemp4(1) = Format(Trim(str(Val(StrTemp4(2)) / Val(StrTemp4(1)) * 100)), "##.##")
Else
    StrTemp4(1) = "0.00"
End If
PrintTotil
If iPrint > 10000 Then
    PrintEnd
    Printer.NewPage
    Page = Page + 1
    PrintTitle
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint > 10000 Then
    PrintEnd
    Printer.NewPage
    Page = Page + 1
    PrintTitle
End If
PrintEnd
Printer.EndDoc
End Sub
'91.8.9 end

Sub PrintTitle()        '印抬頭
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "承辦人准駁明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "准駁日期：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人　：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
'Add By Cheng 2002/08/20
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "准駁條件：" & IIf(Me.txt1(1).Text = "1", "准", "駁")

Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
If txt1(8) = "1" Then
    Printer.Print "承辦人"
Else
    Printer.Print "准駁日"
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
If txt1(8) = "1" Then
    Printer.Print "准駁日"
Else
    Printer.Print "承辦人"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "專利種類"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "准/駁"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "初(再)審"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300

End Sub

Sub PrintEnd()          '印結尾
End Sub

Sub PrintTotil()        '印小計
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print StrConv(MidB(StrConv(StrTemp4(0), vbFromUnicode), 1, 12), vbUnicode)
'Modify By Cheng 2002/12/27
'Printer.CurrentX = 10000
Printer.CurrentX = 10000 - 125
Printer.CurrentY = iPrint
Printer.Print "核准率 "
'Modify By Cheng 2002/12/27
'Printer.CurrentX = 11500 - Printer.TextWidth(Format(CDbl(StrTemp4(1)), "##0.00"))
Printer.CurrentX = 11500 - 125 - Printer.TextWidth(Format(CDbl(StrTemp4(1)), "##0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(CDbl(StrTemp4(1)), "##0.00")
Printer.CurrentX = 11400
Printer.CurrentY = iPrint
Printer.Print " % 小計：准"
Printer.CurrentX = 13300 - Printer.TextWidth(StrTemp4(2))
Printer.CurrentY = iPrint
Printer.Print StrTemp4(2)
Printer.CurrentX = 13500
Printer.CurrentY = iPrint
Printer.Print "駁"
Printer.CurrentX = 14600 - Printer.TextWidth(StrTemp4(3))
Printer.CurrentY = iPrint
Printer.Print StrTemp4(3)
iPrint = iPrint + 300
End Sub

Sub PrintDatil()         '印內容
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()          '讀取位置陣列
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1500
PLeft(2) = 2600
PLeft(3) = 4800
PLeft(4) = 5800
PLeft(5) = 10300
PLeft(6) = 11600
PLeft(7) = 12600
PLeft(8) = 13500
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050311 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/16
   Select Case Index
   Case 1
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 7
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 8
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
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
Case 1
   'Modify By Cheng 2002/09/26
   If Me.txt1(1).Text <> "" Then
     Select Case Val(txt1(Index))
     Case 1, 2
     Case Else
          s = MsgBox("准駁代碼只能 1 或 2 !!", , "USER 輸入錯誤")
          txt1(Index).SetFocus
          txt1(Index).SelStart = 0
          txt1(Index).SelLength = Len(txt1(Index))
          Exit Sub
     End Select
   End If
Case 3, 6
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 4
   lbl1 = GetPrjSales(txt1(Index))
   If txt1(Index) <> "" Then
      'Add By Cheng 2002/09/26
      If Me.txt1(4).Text = Me.lbl1.Caption Then
         Me.lbl1.Caption = ""
         Me.txt1(4).SetFocus
         txt1_GotFocus 4
         Exit Sub
      End If
   End If
Case 7
   'Modify By Cheng 2002/09/26
   If Me.txt1(7).Text <> "" Then
     Select Case txt1(Index)
     Case "Y", "y", ""
     Case Else
          s = MsgBox("是否列印明細只能 Y 或空白 !!", , "USER 輸入錯誤")
          txt1(Index).SetFocus
          txt1(Index).SelStart = 0
          txt1(Index).SelLength = Len(txt1(Index))
          Exit Sub
     End Select
   End If
Case 8
   'Modify By Cheng 2002/09/26
   If Me.txt1(8).Text <> "" Then
     Select Case Val(txt1(Index))
     Case 1, 2
     Case Else
          s = MsgBox("列印順序只能 1 或 2 !!", , "USER 輸入錯誤")
          txt1(Index).SetFocus
          txt1(Index).SelStart = 0
          txt1(Index).SelLength = Len(txt1(Index))
          Exit Sub
     End Select
   End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 2, 3 '准駁日期
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
