VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050404 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人新案案件統計表"
   ClientHeight    =   3360
   ClientLeft      =   720
   ClientTop       =   4365
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form23"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6405
   Begin VB.CheckBox Chk1 
      Caption         =   "含非新申請案"
      Height          =   285
      Index           =   1
      Left            =   210
      TabIndex        =   7
      Top             =   1680
      Width           =   1875
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2700
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1350
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1350
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5325
      TabIndex        =   9
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4530
      TabIndex        =   8
      Top             =   90
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1290
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1020
      Width           =   1410
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2715
      MaxLength       =   7
      TabIndex        =   3
      Top             =   675
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   2
      Top             =   675
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2715
      MaxLength       =   3
      TabIndex        =   1
      Top             =   345
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   0
      Top             =   345
      Width           =   885
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Index           =   2
      Left            =   210
      TabIndex        =   17
      Top             =   2730
      Width           =   6135
      VariousPropertyBits=   27
      Size            =   "10821;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   16
      Top             =   2400
      Width           =   6135
      VariousPropertyBits=   27
      Size            =   "10821;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   15
      Top             =   2070
      Width           =   6135
      VariousPropertyBits=   27
      Size            =   "10821;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   14
      Top             =   1380
      Width           =   900
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   2310
      X2              =   2550
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   2310
      X2              =   2550
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Line Line1 
      X1              =   2310
      X2              =   2550
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
      Height          =   180
      Left            =   2790
      TabIndex        =   13
      Top             =   1050
      Width           =   1380
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收發文日期："
      Height          =   180
      Left            =   210
      TabIndex        =   11
      Top             =   675
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人國籍："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   10
      Top             =   345
      Width           =   1080
   End
End
Attribute VB_Name = "frm050404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/15 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSql As String, strTemp(0 To 9) As String, i As Integer, j As Integer, s As Integer, PLeft(0 To 9) As Integer, iPrint As Integer, Page As Integer, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, StrSQL6 As String
Dim strSQL5 As String
Dim strTemp1 As String, SeekPrint As Integer, SeekPrintL As Integer


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
      If Len(txt1(1)) = 0 And Len(txt1(4)) = 0 Then
         s = MsgBox("代理人國籍區間 或 代理人編號 不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Sub
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
      
      If Len(txt1(3)) = 0 Then
         s = MsgBox("收發文日期區間不可空白!!", , "USER 輸入錯誤")
         txt1(2).SetFocus
         txt1_GotFocus (2)
         Exit Sub
      End If
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
     Process
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Process()
''Add By Cheng 2002/01/29
'Dim strSQL_1 As String
'Dim strSQL_2 As String
'Dim strSQL_3 As String
'Dim strSQL_4 As String

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050404 WHERE ID='" & strUserNum & "' "
strSQL1 = "": strSQL2 = "": StrSQL3 = "": StrSQL4 = "": strSQL5 = "": StrSQL6 = ""
''Add By Cheng 2002/01/29
'strSQL_1 = "": strSQL_2 = "": strSQL_3 = "": strSQL_4 = ""
'若有輸入代理人編號
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA75='" & GetNewFagent(txt1(4)) & "' "
    strSQL2 = strSQL2 + " AND TM44='" & GetNewFagent(txt1(4)) & "' "
    strSQL5 = strSQL5 + " AND cp44='" & GetNewFagent(txt1(4)) & "' "
    StrSQL6 = StrSQL6 + " AND cp44='" & GetNewFagent(txt1(4)) & "' "
    pub_QL05 = pub_QL05 & ";" & Label7 & txt1(4) 'Add By Sindy 2010/10/19
End If
'若未輸入代理人編號
If Len(txt1(4)) = 0 Then
    If Trim(txt1(0)) <> "" Then
      strSQL1 = strSQL1 + " AND FA10>='" & Mid(txt1(0), 1, 3) & "' "
      strSQL2 = strSQL2 + " AND FA10>='" & Mid(txt1(0), 1, 3) & "' "
      strSQL5 = strSQL5 + " AND FA10>='" & Mid(txt1(0), 1, 3) & "' "
      StrSQL6 = StrSQL6 + " AND FA10>='" & Mid(txt1(0), 1, 3) & "' "
    End If
    strSQL1 = strSQL1 + " AND FA10<='" & Mid(txt1(1), 1, 3) & "z' "
    strSQL2 = strSQL2 + " AND FA10<='" & Mid(txt1(1), 1, 3) & "z' "
    strSQL5 = strSQL5 + " AND FA10<='" & Mid(txt1(1), 1, 3) & "z' "
    StrSQL6 = StrSQL6 + " AND FA10<='" & Mid(txt1(1), 1, 3) & "z' "
    pub_QL05 = pub_QL05 & ";" & Label1(0) & Mid(txt1(0), 1, 3) & "-" & Mid(txt1(1), 1, 3) & "z" 'Add By Sindy 2010/10/19
End If
strTemp1 = ""
If Len(txt1(2)) <> 0 Then
   StrSQL3 = StrSQL3 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & " "
   StrSQL4 = StrSQL4 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(txt1(3)) <> 0 Then
   StrSQL3 = StrSQL3 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(3))) & " "
   StrSQL4 = StrSQL4 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(3))) & " "
End If
If Len(txt1(2)) <> 0 Or Len(txt1(3)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/19
End If
'Add By Sindy 2009/10/27
If Chk1(1).Value = 1 Then '1.含非新申請案
   strSQL1 = strSQL1 + " AND cp31='Y' "
   strSQL2 = strSQL2 + " AND cp31='Y' "
   strSQL5 = strSQL5 + " AND cp31='Y' "
   StrSQL6 = StrSQL6 + " AND cp31='Y' "
   pub_QL05 = pub_QL05 & ";" & Chk1(1).Caption 'Add By Sindy 2010/10/19
Else
   strSQL1 = strSQL1 + " AND instr('" & NewCasePtyList & "',cp10)>0 "
   strSQL2 = strSQL2 + " AND cp10='101' "
   strSQL5 = strSQL5 + " AND instr('" & NewCasePtyList & "',cp10)>0 "
   StrSQL6 = StrSQL6 + " AND cp10='101' "
End If
'2009/10/27 End

'Add By Cheng 2002/01/29
'申請國家
If Len(Me.txt1(5).Text) > 0 Then
    'Modify By Cheng 2003/02/26
'   strSQL_1 = " And PA09>='" & Me.txt1(5).Text & "' "
'   strSQL_2 = " And TM10>='" & Me.txt1(5).Text & "' "
'   strSQL_3 = strSQL_1
'   strSQL_4 = strSQL_2
   strSQL1 = strSQL1 & " AND PA09>='" & Me.txt1(5).Text & "' "
   strSQL2 = strSQL2 & " AND TM10>='" & Me.txt1(5).Text & "' "
   strSQL5 = strSQL5 & " AND PA09>='" & Me.txt1(5).Text & "' "
   StrSQL6 = StrSQL6 & " AND TM10>='" & Me.txt1(5).Text & "' "
'   StrSQL3 = StrSQL3 & " And PA09>='" & Me.txt1(5).Text & "' "
'   StrSQL4 = StrSQL4 & " And TM10>='" & Me.txt1(5).Text & "' "
End If
If Len(Me.txt1(6).Text) > 0 Then
    'Modify By Cheng 2003/02/26
'   strSQL_1 = strSQL_1 & " And PA09<='" & Me.txt1(6).Text & "' "
'   strSQL_2 = strSQL_2 & " And TM10<='" & Me.txt1(6).Text & "' "
'   strSQL_3 = strSQL_1
'   strSQL_4 = strSQL_2
   strSQL1 = strSQL1 & " AND PA09<='" & Me.txt1(6).Text & "' "
   strSQL2 = strSQL2 & " AND TM10<='" & Me.txt1(6).Text & "' "
   strSQL5 = strSQL5 & " AND PA09<='" & Me.txt1(6).Text & "' "
   StrSQL6 = StrSQL6 & " AND TM10<='" & Me.txt1(6).Text & "' "
'   StrSQL3 = StrSQL3 & " And PA09<='" & Me.txt1(6).Text & "' "
'   StrSQL4 = StrSQL4 & " And TM10<='" & Me.txt1(6).Text & "' "
End If
If Len(Me.txt1(5).Text) > 0 Or Len(Me.txt1(6).Text) > 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(5) & "-" & txt1(6)  'Add By Sindy 2010/10/19
End If

    '1.FCP,P
    'Modify By Cheng 2003/02/26
'    strSQL = "SELECT FA10,PA75 AS A,CP09," & SQLDate("CP05") & ",pa01||'-'||pa02||'-'||pa03||'-'||pa04 FROM CASEPROGRESS,PATENT,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND cp31='Y' and CP04=PA04(+) AND CP01 IN ('FCP','P') AND CP09<'B' AND CP57 IS NULL AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+)  " & StrSQL3 & strSQL1 & strSQL_1 & " ORDER BY 5,3,2,1 "
    '2006/7/18 MODIFY BY SONIA CP09<'B'->CP09<'C'因為舊系統剛上線時案件無A類FCP-009911
    'strSQL = "SELECT FA10,PA75 AS A,CP09," & SQLDate("CP05") & ",pa01||'-'||pa02||'-'||pa03||'-'||pa04 FROM CASEPROGRESS,PATENT,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND cp31='Y' and CP04=PA04(+) AND CP01 IN ('FCP','P') AND CP09<'B' AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL)) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+)  " & StrSQL3 & strSQL1 & " ORDER BY 5,3,2,1 "
    'Modify By Sindy 2009/10/27 拿掉Where條件(AND cp31='Y')
    'MODIFY BY SONIA 2015/12/1 +CP04='00'否則EPC子案也會算到
    strSql = "SELECT FA10,PA75 AS A,CP09," & SQLDate("CP05") & ",pa01||'-'||pa02||'-'||pa03||'-'||pa04" & _
                   " FROM CASEPROGRESS,PATENT,FAGENT" & _
                   " WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & _
                   " AND CP01 IN ('FCP','P')" & _
                   " AND CP09<'C' AND CP04='00'" & _
                   " AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL))" & _
                   " AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+)" & _
                   StrSQL3 & strSQL1 & _
                   " ORDER BY 5,3,2,1"
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            For i = 0 To 4
                strTemp(i) = CheckStr(adoRecordset.Fields(i))
            Next i
            If strTemp(4) <> strTemp1 Then
                strTemp1 = strTemp(4)
                strSql = "INSERT INTO R050404(R018001,R018002,R018005,R018006,id) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',1,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If
            DoEvents
            adoRecordset.MoveNext
        Loop
    End If
    CheckOC
    '2.FCT,T
    'Modify By Cheng 2003/12/05
'    strSQL = "SELECT FA10,TM44 AS A,CP09," & SQLDate("CP05") & ",tm01||'-'||tm02||'-'||tm03||'-'||tm04 FROM CASEPROGRESS,TRADEMARK,FAGENT WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp31='Y' AND CP01 IN ('FCT','T') AND CP09<'B' AND CP57 IS NULL  AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),'','0',SUBSTR(TM44,9,1))=FA02(+) " & StrSQL4 & strSQL2 & " ORDER BY 5,3,2,1 "
    '2006/7/18 MODIFY BY SONIA CP09<'B'->CP09<'C'因為舊系統剛上線時案件無A類FCP-009911
    'strSQL = "SELECT FA10,TM44 AS A,CP09," & SQLDate("CP05") & ",tm01||'-'||tm02||'-'||tm03||'-'||tm04 FROM CASEPROGRESS,TRADEMARK,FAGENT WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp31='Y' AND CP01 IN ('FCT','T') AND CP09<'B' AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL))  AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),'','0',SUBSTR(TM44,9,1))=FA02(+) " & StrSQL3 & strSQL2 & " ORDER BY 5,3,2,1 "
    'Modify By Sindy 2009/10/27 拿掉Where條件(AND cp31='Y')
    'MODIFY BY SONIA 2015/12/1 +CP04='00'否則EPC子案也會算到
    strSql = "SELECT FA10,TM44 AS A,CP09," & SQLDate("CP05") & ",tm01||'-'||tm02||'-'||tm03||'-'||tm04" & _
                   " FROM CASEPROGRESS,TRADEMARK,FAGENT" & _
                   " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
                   " AND CP01 IN ('FCT','T')" & _
                   " AND CP09<'C' AND CP04='00'" & _
                   " AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL))" & _
                   " AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),'','0',SUBSTR(TM44,9,1))=FA02(+)" & _
                   StrSQL3 & strSQL2 & _
                   " ORDER BY 5,3,2,1"
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            For i = 0 To 4
                strTemp(i) = CheckStr(adoRecordset.Fields(i))
            Next i
            If strTemp(4) <> strTemp1 Then
                strTemp1 = strTemp(4)
                strSql = "INSERT INTO R050404(R018001,R018002,R018009,R018010,id) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',1,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If

            DoEvents
            adoRecordset.MoveNext
        Loop
    End If
    CheckOC
    '3.CFP,P
    'Modify By Cheng 2002/01/29
'    strSQL = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04 FROM CASEPROGRESS,FAGENT WHERE CP01 IN ('CFP','P') AND CP09<'C'  and cp31='Y' AND SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+) and cp57 is null " & StrSQL4 & StrSQL6 & " ORDER BY 5,3,2,1 "
    'Modify By Cheng 2003/02/26
'    strSQL = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04 FROM CASEPROGRESS,FAGENT,Patent WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01 IN ('CFP','P') AND CP09<'C'  and cp31='Y' AND SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+) and cp57 is null " & StrSQL4 & StrSQL6 & strSQL_3 & " ORDER BY 5,3,2,1 "
'    strSQL = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04 FROM CASEPROGRESS,FAGENT,Patent WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01 IN ('CFP','P') AND CP09<'C'  and cp31='Y' AND SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+) and cp57 is null " & StrSQL3 & StrSQL6 & " ORDER BY 5,3,2,1 "
    'Modify By Sindy 2009/10/27 拿掉Where條件(AND cp31='Y')
    'MODIFY BY SONIA 2015/12/1 +CP04='00'否則EPC子案也會算到
    strSql = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
                   " FROM CASEPROGRESS,FAGENT,Patent" & _
                   " WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & _
                   " AND CP01 IN ('CFP')" & _
                   " AND CP09<'C' AND CP04='00'" & _
                   " AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL))" & _
                   " AND SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+)" & _
                   StrSQL4 & strSQL5
    strSql = strSql & " Union SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
                   " FROM CASEPROGRESS,FAGENT,Patent" & _
                   " WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)" & _
                   " AND CP01 IN ('P')" & _
                   " AND CP09<'C' AND CP04='00'" & _
                   " AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL))" & _
                   " AND SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+)" & _
                   StrSQL4 & strSQL5 & _
                   " ORDER BY 5,3,2,1"
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            For i = 0 To 4
                strTemp(i) = CheckStr(adoRecordset.Fields(i))
            Next i
            If strTemp(4) <> strTemp1 Then
                strTemp1 = strTemp(4)
                strSql = "INSERT INTO R050404(R018001,R018002,R018003,R018004,id) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',1,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If

            DoEvents
            adoRecordset.MoveNext
        Loop
    End If
    CheckOC
    '4.CFT,T
      'Modify By Cheng 2002/01/29
'    strSQL = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04 FROM CASEPROGRESS,FAGENT WHERE CP01 IN ('CFT','T') AND CP09<'C' AND cp31='Y' and SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+) and cp57 is null " & StrSQL4 & StrSQL6 & " ORDER BY 5,3,2,1 "
    'Modify By Cheng 2003/02/26
'    strSQL = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04 FROM CASEPROGRESS,FAGENT,TradeMark WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01 IN ('CFT','T') AND CP09<'C' AND cp31='Y' and SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+) and cp57 is null " & StrSQL4 & StrSQL6 & strSQL_4 & " ORDER BY 5,3,2,1 "
'    strSQL = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04 FROM CASEPROGRESS,FAGENT,TradeMark WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01 IN ('CFT','T') AND CP09<'C' AND cp31='Y' and SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+) and cp57 is null " & StrSQL4 & StrSQL6 & " ORDER BY 5,3,2,1 "
    'Modify By Sindy 2009/10/27 拿掉Where條件(AND cp31='Y')
    'MODIFY BY SONIA 2015/12/1 +CP04='00'否則EPC子案也會算到
    strSql = "SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
                   " FROM CASEPROGRESS,FAGENT,TradeMark" & _
                   " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
                   " AND CP01 IN ('CFT')" & _
                   " AND CP09<'C' AND CP04='00'" & _
                   " AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL))" & _
                   " AND SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+)" & _
                   StrSQL4 & StrSQL6
    strSql = strSql & " Union SELECT FA10,CP44 AS A,CP09," & SQLDate("CP27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
                   " FROM CASEPROGRESS,FAGENT,TradeMark" & _
                   " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
                   " AND CP01 IN ('T')" & _
                   " AND CP09<'C' AND CP04='00'" & _
                   " AND (CP57 IS NULL OR (CP57 IS NOT NULL AND CP27 IS NOT NULL))" & _
                   " AND SUBSTR(CP44,1,8)=FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1))=FA02(+)" & _
                   StrSQL4 & StrSQL6 & _
                   " ORDER BY 5,3,2,1"
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            For i = 0 To 4
                strTemp(i) = CheckStr(adoRecordset.Fields(i))
            Next i
            If strTemp(4) <> strTemp1 Then
                strTemp1 = strTemp(4)
                strSql = "INSERT INTO R050404(R018001,R018002,R018007,R018008,id) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',1,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If

            DoEvents
            adoRecordset.MoveNext
        Loop
    End If
    CheckOC
    '總計
    strSql = "SELECT R018001,R018002,SUM(R018003),MIN(R018004)||'-'||MAX(R018004),SUM(R018005),MIN(R018006)||'-'||MAX(R018006),SUM(R018007),MIN(R018008)||'-'||MAX(R018008),SUM(R018009),MIN(R018010)||'-'||MAX(R018010) FROM R050404 WHERE ID='" & strUserNum & "' GROUP BY R018001,R018002 "
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/10/19
        adoRecordset.MoveFirst
        cnnConnection.Execute "DELETE FROM R050404 WHERE ID='" & strUserNum & "' "
        Do While adoRecordset.EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(adoRecordset.Fields(i))
            Next i
            strSql = "INSERT INTO R050404 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(strTemp(2)) & ",'" & ChgSQL(strTemp(3)) & "'," & Val(strTemp(4)) & ",'" & ChgSQL(strTemp(5)) & "'," & Val(strTemp(6)) & ",'" & ChgSQL(strTemp(7)) & "'," & Val(strTemp(8)) & ",'" & ChgSQL(strTemp(9)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            adoRecordset.MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    PrintData
CheckOC
Screen.MousePointer = vbDefault
End Sub

Private Sub PrintData()
'NVL(DECODE(FA10,'020',NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(FA05||' '||FA63||' '||FA64||' '||FA65,NVL(FA01,FA06))),PA75)
'NVL(DECODE(FA10,'020',NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(FA05||' '||FA63||' '||FA64||' '||FA65,NVL(FA01,FA06))),TM44)
'NVL(DECODE(FA10,'020',NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(FA05||' '||FA63||' '||FA64||' '||FA65,NVL(FA01,FA06))),CP44)
'NVL(DECODE(FA10,'020',NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(FA05||' '||FA63||' '||FA64||' '||FA65,NVL(FA01,FA06))),CP44)
strSql = "SELECT nvl(na03,r018001),NVL(DECODE(FA10,'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),r018002),r018003,r018004,r018005,r018006,r018007,r018008,r018009,r018010,r018001,R018002 FROM R050404,NATION,FAGENT WHERE ID='" & strUserNum & "' AND R018002 IS NOT NULL AND R018001=NA01(+) AND SUBSTR(R018002,1,8)=FA01(+) AND DECODE(SUBSTR(R018002,9,1),NULL,'0',SUBSTR(R018002,9,1))=FA02(+) ORDER BY R018001,NVL(DECODE(FA10,'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),r018002),r018002 "
CheckOC
strTemp1 = ""
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(0) = strTemp1 Then
                strTemp(0) = ""
            Else
                strTemp1 = strTemp(0)
            End If
            strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 4), vbUnicode)
            strTemp(1) = StrToStr(strTemp(1), 8)
            
            If strTemp(3) = "-" Then
                strTemp(3) = ""
            End If
            If strTemp(5) = "-" Then
                strTemp(5) = ""
            End If
            If strTemp(7) = "-" Then
                strTemp(7) = ""
            End If
            If strTemp(9) = "-" Then
                strTemp(9) = ""
            End If
            PrintDatil
            If iPrint > 10000 Then
                Page = Page + 1
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = iPrint
                Printer.Print String(204, "-")
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
Else
   ShowNoData
   Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print String(204, "-")
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentY = iPrint
'Modify By Sindy 2009/10/27
If Chk1(1).Value = 1 Then '含非新申請案
   Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "代理人案件統計表") / 2)
   Printer.Print GetTitleNick & "代理人案件統計表"
Else
   Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "代理人新案案件統計表") / 2)
   Printer.Print GetTitleNick & "代理人新案案件統計表"
End If
'2009/10/27 End
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 7500 - (Printer.TextWidth("收/發文日期：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))) / 2)
Printer.CurrentY = iPrint
Printer.Print "收/發文日期：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName

'Add By Sindy 2009/10/27
If Chk1(1).Value = 1 Then '1.含非新申請案
   Printer.CurrentX = 7500 - (Printer.TextWidth("含非新申請案") / 2)
   Printer.CurrentY = iPrint
   Printer.Print "含非新申請案"
End If

Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "國籍代號：" & Format(txt1(0) & " ", "@@@@") & "－" & txt1(1)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print String(204, "-")
iPrint = iPrint + 300

Printer.CurrentX = PLeft(3) + 800
Printer.CurrentY = iPrint
Printer.Print "ＣＦＰ"
Printer.CurrentX = PLeft(5) + 800
Printer.CurrentY = iPrint
Printer.Print "ＦＣＰ"
Printer.CurrentX = PLeft(7) + 800
Printer.CurrentY = iPrint
Printer.Print "ＣＦＴ"
Printer.CurrentX = PLeft(9) + 800
Printer.CurrentY = iPrint
Printer.Print "ＦＣＴ"
iPrint = iPrint + 300
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "國籍"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "代理人名稱"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "首件日期－末件日期"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "首件日期－末件日期"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "首件日期－末件日期"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "首件日期－末件日期"
iPrint = iPrint + 300
Printer.Font.Underline = False
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print String(204, "-")
iPrint = iPrint + 300
End Sub

Private Sub PrintDatil()
For i = 0 To 9
   Select Case i
   Case 2, 4, 6, 8
      Printer.CurrentX = PLeft(i) + 500 - (Printer.TextWidth(strTemp(i)))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Case Else
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   End Select
Next i
iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 500 - 500
PLeft(1) = 1300 - 500
PLeft(2) = 3300 + 700 - 500
PLeft(3) = 4000 + 700 - 500
PLeft(4) = 6450 + 700 - 500
PLeft(5) = 7150 + 700 - 500
PLeft(6) = 9600 + 700 - 500
PLeft(7) = 10300 + 700 - 500
PLeft(8) = 12750 + 700 - 500
PLeft(9) = 13450 + 700 - 500
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050404 = Nothing
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
Case 1, 6
     If RunNick(txt1(Index - 1), txt1(Index)) Then
      txt1(Index - 1).SetFocus
      txt1_GotFocus (Index - 1)
      Exit Sub
   End If
Case 2, 3
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 3 Then
    If RunNick(txt1(2), txt1(3)) Then
      txt1(2).SetFocus
      txt1_GotFocus (2)
      Exit Sub
    End If
  End If
Case 4
    If Trim(txt1(4)) <> "" Then
         CheckOC
         strSql = "select '中'||'  '||fa04,'英'||'  '||fa05||' '||fa63,'日'||'  '||fa06 from fagent where fa01='" & Left(GetNewFagent(txt1(4)), 8) & "' "
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 Then
              lbl1(0).Caption = CheckStr(adoRecordset.Fields(0))
              lbl1(1).Caption = CheckStr(adoRecordset.Fields(1))
              lbl1(2).Caption = CheckStr(adoRecordset.Fields(2))
         Else
              s = MsgBox("代理人編號錯誤!!", , "無此代理人")
              lbl1(0).Caption = ""
              lbl1(1).Caption = ""
              lbl1(2).Caption = ""
              txt1(4).SetFocus
              txt1_GotFocus (4)
              Exit Sub
         End If
         CheckOC
    End If
Case Else
End Select
End Sub
