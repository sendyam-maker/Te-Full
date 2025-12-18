VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060402 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人發文統計表"
   ClientHeight    =   2655
   ClientLeft      =   2880
   ClientTop       =   1860
   ClientWidth     =   4425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4425
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3204
      TabIndex        =   9
      Top             =   24
      Width           =   972
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2412
      TabIndex        =   8
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1830
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2175
      Width           =   585
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2175
      Width           =   585
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1140
      TabIndex        =   5
      Top             =   1845
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1530
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1230
      Width           =   210
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2190
      MaxLength       =   7
      TabIndex        =   2
      Top             =   915
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   1
      Top             =   915
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   585
      Width           =   2595
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Index           =   0
      Left            =   2130
      TabIndex        =   18
      Top             =   1560
      Width           =   1395
      VariousPropertyBits=   27
      Size            =   "2461;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Index           =   1
      Left            =   2130
      TabIndex        =   17
      Top             =   1890
      Width           =   1395
      VariousPropertyBits=   27
      Size            =   "2461;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      X1              =   1725
      X2              =   1830
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2175
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "(1.工程師 2.程序)"
      Height          =   180
      Index           =   6
      Left            =   1425
      TabIndex        =   16
      Top             =   1275
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   15
      Top             =   2190
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "程    序："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   13
      Top             =   1575
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   1245
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   11
      Top             =   930
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   10
      Top             =   630
      Width           =   915
   End
End
Attribute VB_Name = "frm060402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/15 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, StrSQL6 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 29) As String
Dim PLeft(0 To 29) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQLCP09 As String, StrSQL3 As String, StrSQL4 As String
'Add By Cheng 2002/09/17
Dim blnClkSure As Boolean '判斷是否按下確定按鈕


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
   'Add By Cheng 2002/09/17
   blnClkSure = False
   
     Printer.Orientation = 2
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(2)) = 0 Then
             s = MsgBox("發文日區間不可空白!!", , "USER 輸入錯誤")
             'If Len(txt1(2)) = 0 Then txt1(2).SetFocus
             If Len(txt1(1)) = 0 Then txt1(1).SetFocus
             Exit Sub
         Else
            'Add By Cheng 2002/03/21
            If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            'Add By Cheng 2002/09/17
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "發文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
             If Len(txt1(3)) = 0 Then
                 s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                 txt1(3).SetFocus
                 Exit Sub
             Else
               'Add By Cheng 2002/09/17
               lbl1(0).Caption = GetPrjSales(txt1(4))
               If Me.txt1(4).Text <> "" Then
                  If Me.txt1(4).Text = Me.lbl1(0).Caption Then
                     Me.lbl1(0).Caption = ""
                     Me.txt1(4).SetFocus
                     txt1_GotFocus 4
                     Exit Sub
                  End If
               End If
               'Modify By Cheng 2002/09/27
'               lbl1(1).Caption = GetPrjSales(txt1(5))
               lbl1(1).Caption = GetPrjSales(txt1(5), "程序")
               If Me.txt1(5).Text <> "" Then
                  If Me.txt1(5).Text = Me.lbl1(1).Caption Then
                     Me.lbl1(1).Caption = ""
                     Me.txt1(5).SetFocus
                     txt1_GotFocus 5
                     Exit Sub
                  End If
               End If
               If Me.txt1(6).Text <> "" And Me.txt1(7).Text <> "" Then
                  If Me.txt1(6).Text > Me.txt1(7).Text Then
                     MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(6).SetFocus
                     txt1_GotFocus 6
                     Exit Sub
                  End If
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

Sub Process()
ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
cnnConnection.Execute "DELETE FROM R060402 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   StrSQL3 = StrSQL3 + " and cp01 in (" & GetAddStr(txt1(0)) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/13
End If
StrSQL6 = ""
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1)))
    StrSQL3 = StrSQL3 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1)))
End If
If Len(txt1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(2)))
    StrSQL3 = StrSQL3 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(2)))
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/13
End If
If txt1(3) = "1" Then
    If Len(txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " AND CP14='" & txt1(4) & "' "
    End If
    pub_QL05 = pub_QL05 & ";" & Label1(2) & "1.工程師" 'Add By Sindy 2010/12/13
Else
    If Len(txt1(5)) <> 0 Then
        StrSQL6 = StrSQL6 + " AND CP14='" & txt1(5) & "' "
        pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(1) 'Add By Sindy 2010/12/13
    End If
    pub_QL05 = pub_QL05 & ";" & Label1(2) & "2.程序" 'Add By Sindy 2010/12/13
End If
If Len(txt1(4)) <> 0 Then
   StrSQL4 = " AND EP04='" & txt1(4) & "' "
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & lbl1(0) 'Add By Sindy 2010/12/13
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(6) & "' "
    StrSQL3 = StrSQL3 + " and cp10>='" & txt1(6) & "' "
End If
If Len(txt1(7)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(7) & "' "
    StrSQL3 = StrSQL3 + " AND CP10<='" & txt1(7) & "' "
End If
If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/13
End If

CheckOC
StrSQL6 = StrSQL6 & " AND ( (CP09<'C' AND CP10 NOT IN ('907','913')) OR ( CP09>'C' AND (CP10 IN ('1002','1401','1502') OR (CP10>='1201' AND CP10<='1203') OR (CP10>='1301' AND CP10<='1307') OR (CP10>='1504' AND CP10<='1506') OR (CP10>='1801' AND CP10<='1802') OR (CP10>='1805' AND CP10<='1808')))) "
''只抓承辦人為外專的人員
strSql = "SELECT cp14,ST15,cp13,CP10,CP09 FROM CASEPROGRESS,staff WHERE CP14=ST01 and substr(ST15,1,1)='F' " & strSQL1 & StrSQL6
strSql = strSql + "union all select cp14,ST15,cp13,CP10,CP09 FROM CASEPROGRESS,STAFF WHERE CP14=ST01 and substr(ST15,1,1)='F' " & strSQL2 & StrSQL6
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 29
                strTemp(i) = ""
            Next i
            strTemp(0) = CheckStr(.Fields(0))
            strTemp(1) = UCase(CheckStr(.Fields(1)))
            strTemp(2) = CheckStr(.Fields(2))
            Select Case Val(CheckStr(.Fields(3)))
            '************
            '***工程師***
            '************
            Case 201, 210
                 strTemp(3) = "1"
            'Modified by Morgan 2013/11/6 +235核對中說格式
            Case 209, 235
                 strTemp(4) = "1"
            Case 107
                 strTemp(5) = "1"
            Case 501
                 strTemp(6) = "1"
            Case 503, 504, 507
                 strTemp(7) = "1"
            Case 300 To 399
                 strTemp(8) = "1"
            Case 205
                 strTemp(9) = "1"
            Case 203, 204
                 strTemp(10) = "1"
            Case 206
                 strTemp(11) = "1"
            Case 801
                 strTemp(12) = "1"
            Case 803
                 strTemp(13) = "1"
            Case 802
                 strTemp(14) = "1"
            Case 804
                 strTemp(15) = "1"
            Case 1002
                 strTemp(16) = "1"
            Case 1201 To 1203, 1301 To 1307, 1401, 1502, 1504 To 1506, 1801 To 1802, 1805 To 1808
                 strTemp(17) = "1"
            Case 903
                 strTemp(18) = "1"
            '************
            '***承辦人***
            '************
            Case 101
                 strTemp(20) = "1"
            Case 102
                 strTemp(21) = "1"
            Case 103
                 strTemp(22) = "1"
            Case 104
                 strTemp(23) = "1"
            Case 105
                 strTemp(24) = "1"
            Case 601 To 603
                 strTemp(25) = "1"
            Case 605
                 strTemp(26) = "1"
            Case 401
                 strTemp(27) = "1"
            Case 701 To 799
                 strTemp(28) = "1"
            Case Else
                 If strTemp(1) <> "F22" Then
                     strTemp(19) = "1"
                 Else
                     strTemp(29) = "1"
                 End If
            End Select
            'Add By Cheng 2002/05/09
            If Len(strTemp(0)) <= 0 And (Val(.Fields(3)) = 101 Or Val(.Fields(3)) = 102 Or Val(.Fields(3)) = 103 Or _
            Val(.Fields(3)) = 104 Or Val(.Fields(3)) = 105 Or (Val(.Fields(3)) >= 601 And Val(.Fields(3)) <= 603) Or Val(.Fields(3)) = 605 Or _
            Val(.Fields(3)) = 401 Or (Val(.Fields(3)) = 700 And Val(.Fields(3)) <= 799)) Then
               strTemp(1) = "F22"
            End If
            
            strSql = "INSERT INTO R060402 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & "," & Val(strTemp(27)) & "," & Val(strTemp(28)) & "," & Val(strTemp(29)) & ",'" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/13
    ShowNoData
    Exit Sub
End If
CheckOC
'StrSQLCP09 = ""
'strSQL = "SELECT CP09 FROM CASEPROGRESS WHERE (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B' OR (SUBSTR(CP09,1,1)='C' AND (CP10>='101' AND CP10<='105') OR (CP10>='601' AND CP10<='603') OR (CP10 LIKE '7%') OR CP10 IN ('605','401','201','210','209','107','503','504','507','205','203','204','206','801','803','802','804','1002','903','1401','1502') OR CP10 LIKE '3%' OR (CP10>='1201' AND CP10<='1203') OR (CP10>='1301' AND CP10<='1307') OR (CP10>='1504' AND CP10<='1506') OR (CP10>='1801' AND CP10<='1802') OR (CP10>='1805' AND CP10<='1808'))) " & StrSQL3
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    With adoRecordset
'      .MoveFirst
'      Do While .EOF = False
'         StrSQLCP09 = StrSQLCP09 + "'" & chgsql(CheckStr(.Fields(0))) & "'"
'         .MoveNext
'         If .EOF = False Then
'            StrSQLCP09 = StrSQLCP09 + ","
'         End If
'      Loop
'   End With
'End If
CheckOC
'If Len(StrSQLCP09) <> 0 Then
   'Modify By Cheng 2002/04/15
'   strSQL = "SELECT EP04 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE cp09=EP02(+) AND EP04 IS NOT NULL " & StrSQL4 & " and (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B' OR (SUBSTR(CP09,1,1)='C' AND (CP10>='101' AND CP10<='105') OR (CP10>='601' AND CP10<='603') OR (CP10 LIKE '7%') OR CP10 IN ('605','401','201','210','209','107','503','504','507','205','203','204','206','801','803','802','804','1002','903','1401','1502') OR CP10 LIKE '3%' OR (CP10>='1201' AND CP10<='1203') OR (CP10>='1301' AND CP10<='1307') OR (CP10>='1504' AND CP10<='1506') OR (CP10>='1801' AND CP10<='1802') OR (CP10>='1805' AND CP10<='1808'))) " & StrSQL3

'Modify by Morgan 2004/1/29
'只抓承辦人為外專的人員
'   strSQL = "SELECT EP04 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE cp09=EP02(+) AND EP04 IS NOT NULL " & StrSQL4 & " and ( CP09<'C' OR ( CP09>'C' AND (CP10>='101' AND CP10<='105') OR (CP10>='601' AND CP10<='603') OR (CP10 LIKE '7%') OR CP10 IN ('605','401','201','210','209','107','503','504','507','205','203','204','206','801','803','802','804','1002','903','1401','1502') OR CP10 LIKE '3%' OR (CP10>='1201' AND CP10<='1203') OR (CP10>='1301' AND CP10<='1307') OR (CP10>='1504' AND CP10<='1506') OR (CP10>='1801' AND CP10<='1802') OR (CP10>='1805' AND CP10<='1808'))) " & StrSQL3
   strSql = "SELECT EP04 FROM ENGINEERPROGRESS,CASEPROGRESS, STAFF WHERE  CP14=ST01 and substr(ST15,1,1)='F' AND cp09=EP02(+) AND EP04 IS NOT NULL " & StrSQL4 & " and ( CP09<'C' OR ( CP09>'C' AND (CP10>='101' AND CP10<='105') OR (CP10>='601' AND CP10<='603') OR (CP10 LIKE '7%') OR CP10 IN ('605','401','201','210','209','107','503','504','507','205','203','204','206','801','803','802','804','1002','903','1401','1502') OR CP10 LIKE '3%' OR (CP10>='1201' AND CP10<='1203') OR (CP10>='1301' AND CP10<='1307') OR (CP10>='1504' AND CP10<='1506') OR (CP10>='1801' AND CP10<='1802') OR (CP10>='1805' AND CP10<='1808'))) " & StrSQL3
'Modify end -----

   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      With adoRecordset
         .MoveFirst
         Do While .EOF = False
            strSql = "INSERT INTO R060402 VALUES('" & ChgSQL(CheckStr(.Fields(0))) & "','F00',NULL,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
         Loop
      End With
   End If
'End If
CheckOC
'PrintData2
If txt1(3) = "1" Then
'    Process1
    PrintData1
Else
'    Process2
   PrintData2
End If
Printer.EndDoc
ShowPrintOk
End Sub

'Sub PrintData()
'strSQL = "SELECT st02,SUM(R052002),SUM(R052003),SUM(R052004),SUM(R052005),SUM(R052006),SUm(R052007),SUM(R052008),SUM(R052009),SUM(R052010),SUM(R052011),SUM(R052012),SUM(R052013),SUM(R052014),SUM(R052015),SUM(R052016),SUM(R052017),SUM(R052018),SUM(R052019),SUM(R052002)+SUM(R052003)+SUM(R052004)+SUM(R052005)+SUM(R052006)+SUm(R052007)+SUM(R052008)+SUM(R052009)+SUM(R052010)+SUM(R052011)+SUM(R052012)+SUM(R052013)+SUM(R052014)+SUM(R052015)+SUM(R052016)+SUM(R052017)+SUM(R052018)+SUM(R052019),R052001 FROM R060402,staff WHERE R052001=st01(+) and ID='" & strUserNum & "' GROUP BY st02,R052001 order by decode(R052001,null,'0',R052001) "
'CheckOC
'Page = 1
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    With adoRecordset
'        .MoveFirst
'        PrintTitle
 '       Do While .EOF = False
'            For i = 0 To 19
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strTemp(0) = StrToStr(strTemp(0), 3)
'            PrintDatil
'            If iPrint >= 10000 Then
'                Page = Page + 1
'                Printer.NewPage
'                PrintTitle
'            End If
'            .MoveNext
'        Loop
'    End With
'End If
'CheckOC
'Printer.CurrentX = 200
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
'iPrint = iPrint + 300
'PrintEnd
'Printer.EndDoc
'End Sub

'Sub PrintDatil()
'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = iPrint
'Printer.Print strTemp(0)
'For i = 1 To 19
'    Printer.CurrentX = PLeft(i) + 250 - Printer.TextWidth(strTemp(i))
'    Printer.CurrentY = iPrint
'    Printer.Print strTemp(i)
'Next i
'iPrint = iPrint + 300
'End Sub'

'Sub PrintTitle()
'GetPleft
'iPrint = 200
'p rinter.FontName = "細明體"
'p rinter.Font.Size = 22
'Printer.Font.Bold = True
'Printer.Font.Underline = True
'Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "承辦人發文統計表") / 2)
'Printer.CurrentY = iPrint
'Printer.Print GetTitleNick & "承辦人發文統計表"
'iPrint = iPrint + 500
'Printer.Font.Size = 12
'Printer.Font.Bold = False
'Printer.Font.Underline = False
'Printer.CurrentX = 7500 - (Printer.TextWidth("發文日：" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))) / 2)
'Printer.CurrentY = iPrint
'Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
'iPrint = iPrint + 300
'Printer.CurrentX = 200
'Printer.CurrentY = iPrint
'Printer.Print "列印人：" & strUserName
'Printer.CurrentX = 13000
'Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
'iPrint = iPrint + 300
'Printer.CurrentX = 13000
'Printer.CurrentY = iPrint
'Printer.Print "頁    次：" & str(Page)
'iPrint = iPrint + 300
'Printer.Font.Size = 8
'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = iPrint
'If txt1(3) = "1" Then
'    Printer.Print "工程師"
'Else
'    Printer.Print "程  序"
'End If
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iPrint
'Printer.Print "發明"
'Printer.CurrentX = PLeft(2)
'Printer.CurrentY = iPrint
'Printer.Print "新型"
'Printer.CurrentX = PLeft(3)
'Printer.CurrentY = iPrint
'Printer.Print "設計"
'Printer.CurrentX = PLeft(4)
'Printer.CurrentY = iPrint
'Printer.Print "再審"
'Printer.CurrentX = PLeft(5)
'Printer.CurrentY = iPrint
'Printer.Print "訴願"
'Printer.CurrentX = PLeft(6)
'Printer.CurrentY = iPrint
'Printer.Print "再訴願"
'Printer.CurrentX = PLeft(7)
'Printer.CurrentY = iPrint
'Printer.Print "行政"
'Printer.CurrentX = PLeft(8)
'Printer.CurrentY = iPrint
'Printer.Print "修正"
'Printer.CurrentX = PLeft(9)
'Printer.CurrentY = iPrint
'Printer.Print "改請"
'Printer.CurrentX = PLeft(10)
'Printer.CurrentY = iPrint
'Printer.Print "核駁"
'Printer.CurrentX = PLeft(11)
'Printer.CurrentY = iPrint
'Printer.Print "告知"
'Printer.CurrentX = PLeft(12)
'Printer.CurrentY = iPrint
'Printer.Print "專利"
'Printer.CurrentX = PLeft(13)
'Printer.CurrentY = iPrint
'Printer.Print "補充"
'Printer.CurrentX = PLeft(14)
'Printer.CurrentY = iPrint
'Printer.Print "異議"
'Printer.CurrentX = PLeft(15)
'Printer.CurrentY = iPrint
'Printer.Print "舉發"
'Printer.CurrentX = PLeft(16)
'Printer.CurrentY = iPrint
'Printer.Print "異答"
'Printer.CurrentX = PLeft(17)
'Printer.CurrentY = iPrint
'Printer.Print "舉答"
'Printer.CurrentX = PLeft(18)
'Printer.CurrentY = iPrint
'Printer.Print "其他"
'Printer.CurrentX = PLeft(19)
'Printer.CurrentY = iPrint
'Printer.Print "總計"
'iPrint = iPrint + 300
'Printer.CurrentX = PLeft(7)
'Printer.CurrentY = iPrint
'Printer.Print "訴訟"
'Printer.CurrentX = PLeft(9)
'Printer.CurrentY = iPrint
'Printer.Print "新型"
'Printer.CurrentX = PLeft(10)
'Printer.CurrentY = iPrint
'Printer.Print "分析"
'Printer.CurrentX = PLeft(11)
'Printer.CurrentY = iPrint
'Printer.Print "代理人"
'Printer.CurrentX = PLeft(12)
'Printer.CurrentY = iPrint
'Printer.Print "調查"
'Printer.CurrentX = PLeft(13)
'Printer.CurrentY = iPrint
'Printer.Print "說明"
'iPrint = iPrint + 300
'If iPrint >= 10000 Then
'    Page = Page + 1
'    Printer.NewPage
'    PrintTitle
'End If
'Printer.CurrentX = 200
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
'iPrint = iPrint + 300
'If iPrint >= 10000 Then
'    Page = Page + 1
'    Printer.NewPage
'    PrintTitle
'End If
'End Sub
Sub PrintData1()
'89/11/16 nick 改
'印非f22
strSql = "select DECODE(st02,NULL,R049001,ST02),sum(r049004),sum(r049005),sum(r049006),sum(r049007)" & _
         ",sum(r049008),sum(r049009),sum(r049010),sum(r049011),sum(r049012),sum(r049013),sum(r049014)" & _
         ",sum(r049015),sum(r049016),sum(r049017),sum(r049018),sum(r049019),sum(r049020)" & _
         ",sum(R049004+R049005+R049006+R049007+R049008+R049009+R049010+R049011+R049012+R049013+R049014+R049015+R049016+R049017+R049018+R049019+R049020)" & _
         ",r049001 from r060402,staff where r049001=st01(+) and (r049002 <> 'F22' OR R049002 IS NULL) and id='" & strUserNum & "' " & _
         " GROUP BY DECODE(st02,NULL,R049001,ST02),R049001 order by decode(r049001,null,'0',r049001) "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
        .MoveFirst
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 18
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = StrToStr(strTemp(0), 3)
            PrintDatil1
            If iPrint >= 9000 Then
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "備註：1.翻譯含製作中說"
               iPrint = iPrint + 300
               Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
               Printer.CurrentY = iPrint
               Printer.Print "2.核稿含檢視中說"
               iPrint = iPrint + 300
               Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
               Printer.CurrentY = iPrint
               Printer.Print "3.改請含所有改請程序"

                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            .MoveNext
        Loop
    End With
Else
    Exit Sub
End If
CheckOC
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
If iPrint >= 9000 Then
      iPrint = iPrint + 300
      Printer.CurrentX = 0
      Printer.CurrentY = iPrint
      Printer.Print "備註：1.翻譯含製作中說"
      iPrint = iPrint + 300
      Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
      Printer.CurrentY = iPrint
      Printer.Print "2.核稿含檢視中說"
      iPrint = iPrint + 300
      Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
      Printer.CurrentY = iPrint
      Printer.Print "3.改請含所有改請程序"

    Page = Page + 1
    Printer.NewPage
    PrintTitle1
End If
strSql = "select '合  計',sum(r049004),sum(r049005),sum(r049006),sum(r049007)" & _
         ",sum(r049008),sum(r049009),sum(r049010),sum(r049011),sum(r049012),sum(r049013),sum(r049014)" & _
         ",sum(r049015),sum(r049016),sum(r049017),sum(r049018),sum(r049019),sum(r049020)" & _
         ",sum(R049004+R049005+R049006+R049007+R049008+R049009+R049010+R049011+R049012+R049013+R049014+R049015+R049016+R049017+R049018+R049019+R049020)" & _
         " from r060402 where (r049002 <> 'F22' OR R049002 IS NULL) and id='" & strUserNum & "' "
         

'strSQL = "select '合  計',sum(r049004),sum(r049005),sum(r049006),sum(r049007),sum(r049008),sum(r049009),sum(r049010),sum(r049011),sum(r049012),sum(r049013),sum(r049014),sum(r049015),sum(r049016),sum(r049017),sum(r049018),sum(r049019),sum(r049020),sum(R049004+R049005+R049006+R049007+R049008+R049009+R049010+R049011+R049012+R049013+R049014+R049015+R049016+R049017+R049018+R049019+R049020) from r060401 where (r049002 <> 'F22' OR R049002 IS NULL) and id='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 18
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            PrintDatil1
            If iPrint >= 9000 Then
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "備註：1.翻譯含製作中說"
               iPrint = iPrint + 300
               Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
               Printer.CurrentY = iPrint
               Printer.Print "2.核稿含檢視中說"
               iPrint = iPrint + 300
               Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
               Printer.CurrentY = iPrint
               Printer.Print "3.改請含所有改請程序"
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            .MoveNext
        Loop
    End With
End If
CheckOC
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "備註：1.翻譯含製作中說"
iPrint = iPrint + 300
Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
Printer.CurrentY = iPrint
Printer.Print "2.核稿含檢視中說"
iPrint = iPrint + 300
Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
Printer.CurrentY = iPrint
Printer.Print "3.改請含所有改請程序"
'Printer.EndDoc
Printer.NewPage
End Sub

Sub PrintData2()
'89/11/16 nick 改
'印f22
strSql = "select DECODE(st02,NULL,R049001,ST02),sum(r049021),sum(r049022),sum(r049023),sum(r049024)" & _
         ",sum(r049025),sum(r049026),sum(r049027),sum(r049028),sum(r049029),sum(r049030)" & _
         ",sum(R049021+R049022+R049023+R049024+R049025+R049026+R049027+R049028+R049029+R049030)" & _
         ",r049001 from r060402,staff where r049001=st01(+) and r049002 = 'F22' and id='" & strUserNum & "' " & _
         " GROUP BY DECODE(st02,NULL,R049001,ST02),R049001 order by decode(r049001,null,'0',r049001) "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
        .MoveFirst
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = StrToStr(strTemp(0), 3)
            PrintDatil2
            If iPrint >= 9000 Then
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "備註：1.領證含加註追加、加註聯合"
               iPrint = iPrint + 300
               Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
               Printer.CurrentY = iPrint
               Printer.Print "2.讓與含合併、繼承、授權、設定質權"
               iPrint = iPrint + 300
                Page = Page + 1
                Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    End With
Else
    Exit Sub
End If
CheckOC
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
If iPrint >= 9000 Then
      iPrint = iPrint + 300
      Printer.CurrentX = 0
      Printer.CurrentY = iPrint
      Printer.Print "備註：1.領證含加註追加、加註聯合"
      iPrint = iPrint + 300
      Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
      Printer.CurrentY = iPrint
      Printer.Print "2.讓與含合併、繼承、授權、設定質權"

    Page = Page + 1
    Printer.NewPage
    PrintTitle2
End If
strSql = "select '合  計',sum(r049021),sum(r049022),sum(r049023),sum(r049024)" & _
         ",sum(r049025),sum(r049026),sum(r049027),sum(r049028),sum(r049029),sum(r049030)" & _
         ",sum(R049021+R049022+R049023+R049024+R049025+R049026+R049027+R049028+R049029+R049030)" & _
         " from r060402 where r049002 = 'F22' and id='" & strUserNum & "' "
         
CheckOC
'strSQL = "select '合  計',sum(r049021),sum(r049022),sum(r049023),sum(r049024),sum(r049025),sum(r049026),sum(r049027),sum(r049028),sum(r049029),sum(r049030),sum(R049021+R049022+R049023+R049024+R049025+R049026+R049027+R049028+R049029+R049030) from r060401 where r049002 = 'F22' and id='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            PrintDatil2
            If iPrint >= 9000 Then
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "備註：1.領證含加註追加、加註聯合"
               iPrint = iPrint + 300
               Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
               Printer.CurrentY = iPrint
               Printer.Print "2.讓與含合併、繼承、授權、設定質權"
                Page = Page + 1
                Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    End With
End If
CheckOC
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "備註：1.領證含加註追加、加註聯合"
iPrint = iPrint + 300
Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
Printer.CurrentY = iPrint
Printer.Print "2.讓與含合併、繼承、授權、設定質權"
'Printer.EndDoc
End Sub
Sub PrintTitle1()
GetPleft1
iPrint = 200
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "承辦人發文統計表") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "承辦人發文統計表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 7500 - (Printer.TextWidth("發文日：" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))) / 2)
Printer.CurrentY = iPrint
Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 200
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
Printer.Font.Size = 8
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "工程師"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "翻譯"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "核稿"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "再審"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "訴願"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "行政訴訟"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "改請"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申復"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "修正"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "補充說明"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "異議"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "舉發"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "異議答辯"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "舉發答辯"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "核駁分析"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "主管機關"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "專利調查"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "總計"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "　以上"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "　來文"
iPrint = iPrint + 300
If iPrint >= 9000 Then
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "備註：1.翻譯含製作中說"
   iPrint = iPrint + 300
   Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
   Printer.CurrentY = iPrint
   Printer.Print "2.核稿含檢視中說"
   iPrint = iPrint + 300
   Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
   Printer.CurrentY = iPrint
   Printer.Print "3.改請含所有改請程序"
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
If iPrint >= 9000 Then
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "備註：1.翻譯含製作中說"
   iPrint = iPrint + 300
   Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
   Printer.CurrentY = iPrint
   Printer.Print "2.核稿含檢視中說"
   iPrint = iPrint + 300
   Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
   Printer.CurrentY = iPrint
   Printer.Print "3.改請含所有改請程序"
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
End If
End Sub

Sub PrintDatil1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 1 To 18
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(strTemp(i))
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub


Sub GetPleft1()
Erase PLeft
PLeft(0) = 0
For i = 1 To 18
   PLeft(i) = 150 + (i * 800)
Next i
'PLeft(1) = 750
'PLeft(2) = 1300
'PLeft(3) = 1850
'p 'Left(4) = 2400
'p 'Left(5) = 2950
'PLeft(6) = 3500
'PLeft(7) = 4050
'PLeft(8) = 4600
'PLeft(9) = 5150
'PLeft(10) = 5700
'PLeft(11) = 6250
'PLeft(12) = 6800
'PLeft(13) = 7350
'PLeft(14) = 7900'
'PLeft(15) = 8450
'PLeft(16) = 9000
'PLeft(17) = 9550
'PLeft(18) = 10100
'PLeft(19) = 10650
'PLeft(20) = 11200
'PLeft(21) = 11750
'PLeft(22) = 12300
'PLeft(23) = 12850
'PLeft(24) = 13400
'PLeft(25) = 13950
'PLeft(26) = 14500
'PLeft(27) = 15050
'PLeft(28) = 15600
End Sub


Sub PrintTitle2()
GetPleft2
iPrint = 200
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "承辦人發文統計表") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "承辦人發文統計表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 7500 - (Printer.TextWidth("發文日：" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))) / 2)
Printer.CurrentY = iPrint
Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 200
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
Printer.Font.Size = 8
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "程序"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "發明"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "新型"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "設計"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "追加"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "聯合"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "領證"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "年費"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "變更"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "讓與"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "其他"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "總計"
iPrint = iPrint + 300
If iPrint >= 9000 Then
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "備註：1.領證含加註追加、加註聯合"
   iPrint = iPrint + 300
   Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
   Printer.CurrentY = iPrint
   Printer.Print "2.讓與含合併、繼承、授權、設定質權"
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
If iPrint >= 9000 Then
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "備註：1.領證含加註追加、加註聯合"
   iPrint = iPrint + 300
   Printer.CurrentX = 0 + (Printer.TextWidth("備註："))
   Printer.CurrentY = iPrint
   Printer.Print "2.讓與含合併、繼承、授權、設定質權"
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
End If
End Sub

Sub PrintDatil2()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 1 To 11
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(strTemp(i))
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft2()
Erase PLeft
PLeft(0) = 0
For i = 1 To 11
   PLeft(i) = 150 + (i * 800)
Next i

'PLeft(1) = 900
'PLeft(2) = 1600
'PLeft(3) = 2300
'PLeft(4) = 3000
'PLeft(5) = 3700
'PLeft(6) = 4400
'PLeft(7) = 5100
'PLeft(8) = 5800
'PLeft(9) = 6500
'PLeft(10) = 7200
'PLeft(11) = 7900
'PLeft(12) = 8600
'PLeft(13) = 9300
'PLeft(14) = 10000
'PLeft(15) = 10700
'PLeft(16) = 11400
'PLeft(17) = 12100
'PLeft(18) = 12800
'PLeft(19) = 13500
End Sub

'Sub GetPleft()
'Erase PLeft
'PLeft(0) = 200
'PLeft(1) = 900
'PLeft(2) = 1600
'PLeft(3) = 2300
'PLeft(4) = 3000
'PLeft(5) = 3700
'PLeft(6) = 4400
'PLeft(7) = 5100
'PLeft(8) = 5800
'PLeft(9) = 6500
'PLeft(10) = 7200
'PLeft(11) = 7900
'PLeft(12) = 8600
'PLeft(13) = 9300
'PLeft(14) = 10000
'PLeft(15) = 10700
'PLeft(16) = 11400
'PLeft(17) = 12100
'PLeft(18) = 12800
'PLeft(19) = 13500
'End Sub


'Sub PrintEnd()
'strSQL = "SELECT '合  計',SUM(R052002),SUM(R052003),SUM(R052004),SUM(R052005),SUM(R052006),SUm(R052007),SUM(R052008),SUM(R052009),SUM(R052010),SUM(R052011),SUM(R052012),SUM(R052013),SUM(R052014),SUM(R052015),SUM(R052016),SUM(R052017),SUM(R052018),SUM(R052019),SUM(R052002)+SUM(R052003)+SUM(R052004)+SUM(R052005)+SUM(R052006)+SUm(R052007)+SUM(R052008)+SUM(R052009)+SUM(R052010)+SUM(R052011)+SUM(R052012)+SUM(R052013)+SUM(R052014)+SUM(R052015)+SUM(R052016)+SUM(R052017)+SUM(R052018)+SUM(R052019) FROM R060402 WHERE ID='" & strUserNum & "' "
'CheckOC2
'Page = 1
'adoRecordset1.CursorLocation = adUseClient
'adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'    With adoRecordset1
'        .MoveFirst
'        Do While .EOF = False
'            For i = 0 To 19
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strTemp(0) = StrToStr(strTemp(0), 3)
'            PrintDatil
'            .MoveNext
'        Loop
'    End With
'End If
'CheckOC2
'End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'Modify by Morgan 2004/1/29
'固定預設
'txt1(0) = GetSystemKindByNick
txt1(0) = "FCP,FG,CFP,CPS,P,PS"
'Modify End -----
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm060402 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/17
   Select Case Index
   Case 3
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
   
End Sub

Private Sub txt1_LostFocus(Index As Integer)

Select Case Index
Case 0
     txt1(0) = UCase(txt1(0))
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        '93.12.29 MODIFY BY SONIA
        'If s = 0 Then
        If s = 0 And strTemp2(i) <> "CFP" And strTemp2(i) <> "CPS" And strTemp2(i) <> "P" And strTemp2(i) <> "PS" Then
        '93.12.29 END
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    Next i
Case 2
   'Modify By Cheng 2002/09/17
   If blnClkSure = False Then
     If RunNick(txt1(1), txt1(2)) Then
         txt1(1).SetFocus
         txt1_GotFocus (1)
     End If
   Else
      blnClkSure = False
   End If
Case 3
   'Modify By Cheng 2002/09/26
   'If Me.txt1(3).Text <> "" Then
     Select Case Val(txt1(3))
     Case 1, 2
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(3).SetFocus
          txt1(3).SelStart = 0
          txt1(3).SelLength = Len(txt1(3))
          Exit Sub
     End Select
   'End If
Case 4
     lbl1(0).Caption = GetPrjSales(txt1(4))
     'Add By Cheng 2002/09/26
     If Me.txt1(4).Text <> "" Then
         If Me.txt1(4).Text = Me.lbl1(0).Caption Then
            Me.lbl1(0).Caption = ""
            Me.txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
     End If
Case 5
      'Modify By Cheng 2002/09/27
'     lbl1(1).Caption = GetPrjSales(txt1(5))
     lbl1(1).Caption = GetPrjSales(txt1(5), "程序")
     'Add By Cheng 2002/09/26
     If Me.txt1(5).Text <> "" Then
         If Me.txt1(5).Text = Me.lbl1(1).Caption Then
            Me.lbl1(1).Caption = ""
            Me.txt1(5).SetFocus
            txt1_GotFocus 5
            Exit Sub
         End If
     End If
Case 7
   'Modify By Cheng 2002/09/17
   If blnClkSure = False Then
     If RunNick(txt1(6), txt1(7)) Then
         txt1(6).SetFocus
         txt1_GotFocus (6)
      End If
   Else
      blnClkSure = False
   End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '發文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
