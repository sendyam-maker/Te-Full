VERSION 5.00
Begin VB.Form frm050315 
   BorderStyle     =   1  '單線固定
   Caption         =   "延期明細表"
   ClientHeight    =   1800
   ClientLeft      =   2985
   ClientTop       =   1770
   ClientWidth     =   3120
   ControlBox      =   0   'False
   LinkTopic       =   "Form14"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3120
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   960
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1152
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2244
      TabIndex        =   7
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1416
      TabIndex        =   6
      Top             =   20
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1488
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   960
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1488
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   2
      Top             =   816
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   960
      MaxLength       =   7
      TabIndex        =   1
      Top             =   816
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2130
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(次以上)"
      Height          =   180
      Left            =   1920
      TabIndex        =   12
      Top             =   1152
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "延期次數："
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   1155
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   2160
      Y1              =   1608
      Y2              =   1608
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   2160
      Y1              =   936
      Y2              =   936
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申請國家："
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "延期日期："
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm050315"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, k As Integer
Dim strTemp(0 To 10) As String, PLeft(0 To 10) As Integer, iPrint As Integer
Dim strTemp1 As Variant, strTemp2 As Variant, Page As Integer, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim SeekTmp(0 To 1) As String
Dim strSQL1N As String, strSQL2N As String, strSQL3N As String, strSQL4N As String, strSQL5N As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
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
        
        If Len(txt1(2)) = 0 Then
            s = MsgBox("延期日期不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1_GotFocus (1)
            Exit Sub
        Else
            If Len(Trim(txt1(3))) = 0 Then
                s = MsgBox("延期次數不可空白!!", , "USER 輸入錯誤")
                txt1(3).SetFocus
                Exit Sub
            Else
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
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

Private Sub Process()
Dim strDateLimit1 As String '延期資料本所期限
Dim strDateLimit2 As String '延期資料法定期限

cnnConnection.Execute "DELETE FROM R050315_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R050315_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
strSQL1N = ""
strSQL2N = ""
strSQL3N = ""
strSQL4N = ""
strSQL5N = ""

If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   StrSQL3 = StrSQL3 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") "
   StrSQL4 = StrSQL4 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") "
   strSQL5 = strSQL5 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   strSQL1N = strSQL1N + " AND NP02 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2N = strSQL2N + " AND NP02 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL3N = strSQL3N + " AND NP02 IN (" & SQLGrpStr(txt1(0), 3) & ") "
   strSQL4N = strSQL4N + " AND NP02 IN (" & SQLGrpStr(txt1(0), 4) & ") "
   strSQL5N = strSQL5N + " AND NP02 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/10/4
End If
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & txt1(4) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(4) & "' "
    strSQL5 = strSQL5 + " AND SP09>='" & txt1(4) & "' "
    strSQL1N = strSQL1N + " AND PA09>='" & txt1(4) & "' "
    strSQL2N = strSQL2N + " AND TM10>='" & txt1(4) & "' "
    strSQL3N = strSQL3N + " AND LC15>='" & txt1(4) & "' "
    strSQL5N = strSQL5N + " AND SP09>='" & txt1(4) & "' "
End If
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & txt1(5) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(5) & "' "
    strSQL5 = strSQL5 + " AND SP09<='" & txt1(5) & "' "
    strSQL1N = strSQL1N + " AND PA09<='" & txt1(5) & "' "
    strSQL2N = strSQL2N + " AND TM10<='" & txt1(5) & "' "
    strSQL3N = strSQL3N + " AND LC15<='" & txt1(5) & "' "
    strSQL5N = strSQL5N + " AND SP09<='" & txt1(5) & "' "
End If
If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label4 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/4
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/4
End If
If Len(txt1(3)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(3) & Label5 'Add By Sindy 2010/10/4
End If
'先取得符合條件的延期記錄資料
CheckOC
strSql = "SELECT DL01,COUNT(DL01) FROM DATELIMIT WHERE DL02>=" & Val(ChangeTStringToWString(txt1(1))) & " AND DL02<=" & Val(ChangeTStringToWString(txt1(2))) & " GROUP BY DL01 "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 1
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        If Val(strTemp(1)) >= Val(txt1(3)) Then
            strSql = "SELECT DL02 FROM DATELIMIT WHERE DL02>=" & Val(ChangeTStringToWString(txt1(1))) & " AND DL02<=" & Val(ChangeTStringToWString(txt1(2))) & " AND DL01='" & ChgSQL(strTemp(0)) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                Do While adoRecordset1.EOF = False
                    strSql = "insert into R050315_1 values('" & ChgSQL(strTemp(0)) & "','" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0)))) & "','" & strUserNum & "') "
                    cnnConnection.Execute strSql
                    adoRecordset1.MoveNext
                Loop
            End If
            CheckOC2
        End If
        adoRecordset.MoveNext
    Loop
Else
   InsertQueryLog (0) 'Add By Sindy 2010/10/4
   ShowNoData
   Exit Sub
End If
CheckOC
strSql = "SELECT * FROM R050315_1,DateLimit WHERE R014001=DL01(+) AND (REPLACE(R014002,'/','') + 19110000)=DL02(+) AND ID='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
         If IsNull(adoRecordset("DL03")) = False Then
            strDateLimit1 = IIf(IsNull(adoRecordset("DL03")), "", ChangeTStringToTDateString(adoRecordset("DL03") - 19110000))
         Else
            strDateLimit1 = ""
         End If
         If IsNull(adoRecordset("DL04")) = False Then
            strDateLimit2 = IIf(IsNull(adoRecordset("DL04")), "", ChangeTStringToTDateString(adoRecordset("DL04") - 19110000))
         Else
            strDateLimit2 = ""
         End If
         '若資料來源為案件進度檔, 業務區別為CP12
         If adoRecordset("DL05") = "1" Then
           strSql = "SELECT CP12,CP13,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),S2.ST02,CP64,'" & strUserNum & "',NVL(NA03,NA04) FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,NATION WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND PA09=NA01(+) " & strSQL1
           strSql = strSql + " union all select CP12,CP13,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),S2.ST02,CP64,'" & strUserNum & "',NVL(NA03,NA04) FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,NATION WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND SP09=NA01(+) " & strSQL5
           strSql = strSql + " union all select CP12,CP13,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),S2.ST02,CP64,'" & strUserNum & "',NVL(NA03,NA04) FROM CASEPROGRESS,LAWCASE,STAFF S1,STAFF S2,CASEPROPERTYMAP,NATION WHERE cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) AND CP13=S1.ST01(+)  AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND LC15=NA01(+) " & StrSQL3
           strSql = strSql + " union all select CP12,CP13,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),S2.ST02,CP64,'" & strUserNum & "',NVL(NA03,NA04) FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,NATION WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  AND CP13=S1.ST01(+)  AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND TM10=NA01(+) " & strSQL2
           strSql = strSql + " union all select CP12,CP13,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),S2.ST02,CP64,'" & strUserNum & "',NVL(NA03,NA04) FROM CASEPROGRESS,HIRECASE,STAFF S1,STAFF S2,CASEPROPERTYMAP,NATION WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) AND CP04=HC04(+) AND CP13=S1.ST01(+)  AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND '000'=NA01(+) " & StrSQL4
           cnnConnection.Execute "INSERT INTO R050315_2 " & strSql
         '若資料來源為下一程序檔, 業務區別為ST15
         Else
            If Not IsNull(adoRecordset("DL06")) Then
               strSql = "SELECT S1.ST15,NP10,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(PA05,NVL(PA06,PA07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),'',NP15,'" & strUserNum & "',NVL(NA03,NA04) FROM NEXTPROGRESS,PATENT,STAFF S1,CASEPROPERTYMAP,NATION WHERE NP02=pa01(+) and Np03=pa02(+) and Np04=pa03(+) and Np05=pa04(+)  AND NP10=S1.ST01(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND NP01='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND PA09=NA01(+) AND NP22=" & adoRecordset("DL06") & " " & strSQL1N
               strSql = strSql + " union all select S1.ST15,NP10,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),'',NP15,'" & strUserNum & "',NVL(NA03,NA04) FROM NEXTPROGRESS,SERVICEPRACTICE,STAFF S1,CASEPROPERTYMAP,NATION WHERE Np02=sp01(+) and Np03=sp02(+) and Np04=sp03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND NP01='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND SP09=NA01(+) AND NP22=" & adoRecordset("DL06") & " " & strSQL5N
               strSql = strSql + " union all select S1.ST15,NP10,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(LC05,NVL(LC06,LC07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),'',NP15,'" & strUserNum & "',NVL(NA03,NA04) FROM NEXTPROGRESS,LAWCASE,STAFF S1,CASEPROPERTYMAP,NATION WHERE Np02=lc01(+) and Np03=lc02(+) and Np04=lc03(+) and Np05=lc04(+) AND NP10=S1.ST01(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND NP01='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND LC15=NA01(+) AND NP22=" & adoRecordset("DL06") & " " & strSQL3N
               strSql = strSql + " union all select S1.ST15,NP10,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),'',NP15,'" & strUserNum & "',NVL(NA03,NA04) FROM NEXTPROGRESS,TRADEMARK,STAFF S1,CASEPROPERTYMAP,NATION WHERE Np02=tm01(+) and Np03=tm02(+) and Np04=tm03(+) and Np05=tm04(+)  AND NP10=S1.ST01(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND NP01='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND TM10=NA01(+) AND NP22=" & adoRecordset("DL06") & " " & strSQL2N
               strSql = strSql + " union all select S1.ST15,NP10,'" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',NP02||'-'||NP03||'-'||NP04||'-'||NP05,HC06,'" & strDateLimit1 & "','" & strDateLimit2 & "',NVL(CPM03,CPM04),'',NP15,'" & strUserNum & "',NVL(NA03,NA04) FROM NEXTPROGRESS,HIRECASE,STAFF S1,CASEPROPERTYMAP,NATION WHERE Np02=hc01(+) and Np03=hc02(+) and Np04=hc03(+) AND NP05=HC04(+) AND NP10=S1.ST01(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND NP01='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND '000'=NA01(+) AND NP22=" & adoRecordset("DL06") & " " & strSQL4N
               cnnConnection.Execute "INSERT INTO R050315_2 " & strSql
            End If
         End If
        DoEvents
        adoRecordset.MoveNext
    Loop
    strSql = "SELECT * FROM R050315_2 WHERE ID='" & strUserNum & "' "
    CheckOC2
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 Then
      InsertQueryLog (adoRecordset1.RecordCount) 'Add By Sindy 2010/10/4
      PrintData
    Else
      InsertQueryLog (0) 'Add By Sindy 2010/10/4
      ShowNoData
      CheckOC2
      Exit Sub
    End If
Else
    InsertQueryLog (0) 'Add By Sindy 2010/10/4
    ShowNoData
    Exit Sub
End If
CheckOC
End Sub

Private Sub PrintData()
strSql = "SELECT NVL(A0902,A0903),ST02,R015003,R015004,R015005,R015006,R015007,R015008,R015009,R015011,R015010,R015001,R015002 FROM R050315_2,STAFF,ACC090 WHERE R015001=A0901(+) AND R015002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R015001,R015002,R015003 "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        SeekTmp(0) = ""
        SeekTmp(1) = ""
        Do While .EOF = False
'            For i = 0 To 9
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 18), vbUnicode)
            strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 10), vbUnicode)
            strTemp(8) = StrConv(MidB(StrConv(strTemp(8), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(9) = StrConv(MidB(StrConv(strTemp(9), vbFromUnicode), 1, 4), vbUnicode)
            strTemp(10) = StrConv(MidB(StrConv(strTemp(10), vbFromUnicode), 1, 16), vbUnicode)
            If SeekTmp(0) = "" And SeekTmp(1) = "" Then
               SeekTmp(0) = strTemp(0)
               SeekTmp(1) = strTemp(1)
            Else
               If SeekTmp(0) = strTemp(0) Then
                  strTemp(0) = ""
                  If SeekTmp(1) = strTemp(1) Then
                     strTemp(1) = ""
                  Else
                     SeekTmp(1) = strTemp(1)
                  End If
               Else
                  SeekTmp(0) = strTemp(0)
                  SeekTmp(1) = strTemp(1)
               End If
            End If
            If iPrint > 10000 Then
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 1
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
            .MoveNext
        Loop
    End With
End If
CheckOC
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "延期明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "延期日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "延期 " & txt1(3) & " 次以上 "
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "延期日"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "法定期限"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
'Add By Cheng 2002/02/19
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "申請國家"

Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "備註"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300

End Sub

Private Sub PrintDatil()
'For i = 0 To 9
For i = 0 To 10
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300

End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 500 - 500
PLeft(1) = 1800 - 500
PLeft(2) = 2950 - 500
PLeft(3) = 4100 - 500
PLeft(4) = 5800
PLeft(5) = 8300
PLeft(6) = 9500
PLeft(7) = 10700
PLeft(8) = 12200
PLeft(9) = 13000
PLeft(10) = 14000
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050315 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
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
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 1, 2
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
Case 5
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 3
     For i = 1 To Len(Trim(txt1(3)))
        Select Case Mid(Trim(txt1(3)), i, 1)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
        Case Else
             s = MsgBox("延期次數請輸入數字!!", , "USER 輸入錯誤")
             txt1(3).SetFocus
             txt1(3).SelStart = 0
             txt1(3).SelLength = Len(txt1(3))
             Exit Sub
        End Select
     Next i
Case Else
End Select
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub


