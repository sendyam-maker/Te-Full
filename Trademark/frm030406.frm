VERSION 5.00
Begin VB.Form frm030406 
   BorderStyle     =   1  '單線固定
   Caption         =   "催審表"
   ClientHeight    =   2535
   ClientLeft      =   4500
   ClientTop       =   2955
   ClientWidth     =   3795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3795
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   9
      Left            =   1350
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1500
      Width           =   435
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   10
      Left            =   1860
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1500
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   11
      Left            =   2910
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1500
      Width           =   225
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   12
      Left            =   3210
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1500
      Width           =   360
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2925
      TabIndex        =   13
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2130
      TabIndex        =   12
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   7
      Left            =   2460
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1860
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   6
      Left            =   1350
      MaxLength       =   4
      TabIndex        =   9
      Top             =   1860
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   8
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "1"
      Top             =   2205
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   5
      Left            =   2460
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1215
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   1350
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1200
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   2460
      MaxLength       =   7
      TabIndex        =   2
      Top             =   900
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   1350
      MaxLength       =   7
      TabIndex        =   1
      Top             =   900
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   525
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "發文日期："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   1260
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "催審期限："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   975
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   1530
      Width           =   1200
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1620
      X2              =   3440
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1890
      X2              =   3090
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1845
      X2              =   3045
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1830
      X2              =   3030
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   19
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   18
      Top             =   2250
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Top             =   585
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "(1. 管制表  2.定稿)"
      Height          =   180
      Index           =   5
      Left            =   1830
      TabIndex        =   16
      Top             =   2250
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frm030406"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay(0 To 1) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, BolProcess As Boolean
'Add By Cheng 2002/09/17
Dim m_strUserRight As String '使用者系統類別使用權限
Dim m_arrUserRight '使用者系統類別使用權限陣列
Dim ii As Integer '回圈序號
Dim blnUserRight As Boolean '是否有此系統類別權限

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
         '選擇催審期限
         If Option1(0).Value = True Then
           'Add By Cheng 2002/03/21
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
                 s = MsgBox("催審期限區間不可空白!!", , "USER 輸入錯誤")
                 txt1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             End If
         '選擇發文日期
         ElseIf Me.Option1(1).Value Then
           'Add By Cheng 2002/03/21
           If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
              Me.txt1(4).SetFocus
              txt1_GotFocus 4
              Exit Sub
           End If
           If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
              Me.txt1(5).SetFocus
              txt1_GotFocus 5
              Exit Sub
           End If
             
             If Len(txt1(5)) = 0 Then
                 s = MsgBox("發文日期區間不可空白!!", , "USER 輸入錯誤")
                 txt1(4).SetFocus
                 txt1_GotFocus (4)
                 Exit Sub
             End If
         Else
            'Add By Cheng 2002/09/17
            If Me.txt1(9).Text = "" Then
               MsgBox "請輸入本所案號的系統類別!!!", vbExclamation + vbOKOnly
               Me.txt1(9).SetFocus
               txt1_GotFocus 9
               Exit Sub
            End If
            blnUserRight = False
            If Me.txt1(9).Text <> "" Then
               If m_strUserRight <> "" Then
                  For ii = LBound(m_arrUserRight) To UBound(m_arrUserRight)
                     If m_arrUserRight(ii) = Me.txt1(9).Text Then
                        blnUserRight = True
                     End If
                  Next ii
                  If blnUserRight = False Then
                     MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
                     Me.txt1(9).SetFocus
                     txt1_GotFocus 9
                     Exit Sub
                  End If
               Else
                  MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.txt1(9).SetFocus
                  txt1_GotFocus 9
                  Exit Sub
               End If
            End If
         End If
         If Len(txt1(8)) = 0 Then
             s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
             txt1(8).SetFocus
             Exit Sub
         Else
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
             If txt1(8) = "1" Then
                pub_QL05 = pub_QL05 & ";" & Label1(2) & "管制表"  'Add By Sindy 2010/10/22
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                Process
                Me.Enabled = True
                Screen.MousePointer = vbDefault
             Else
                'Memo by Lydia 2018/03/16 隱藏列印別,預設為管制表
                pub_QL05 = pub_QL05 & ";" & Label1(2) & "定稿"  'Add By Sindy 2010/10/22
                s = MsgBox("定稿部分尚未完成!!")
                Exit Sub
             End If
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R030406 WHERE ID='" & strUserNum & "' "
'選擇催審期限
If Option1(0).Value = True Then
    strSQL1 = ""
    strSQL2 = ""
    StrSQL6 = ""
    If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 2) & ") "
      strSQL2 = strSQL2 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0).Caption & txt1(0) 'Add By Sindy 2010/10/22
    End If
    StrSQL6 = ""
    If Len(Trim(txt1(2))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(2))) & " "
    End If
    If Len(Trim(txt1(3))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND NP08<=" & Val(ChangeTStringToWString(txt1(3))) & " "
    End If
    If Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/22
    End If
    StrSQL6 = StrSQL6 & " AND NP07=305 AND (NP06 IS NULL OR NP06='') "
    strSQL1 = strSQL1 + " AND (TM29 IS NULL OR TM29='') "
    strSQL2 = strSQL2 + " AND (SP15 IS NULL OR SP15='') "
    If Len(txt1(6)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10>='" & txt1(6) & "' "
        strSQL2 = strSQL2 + " AND SP09>='" & txt1(6) & "' "
    End If
    If Len(txt1(7)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10<='" & txt1(7) & "' "
        strSQL2 = strSQL2 + " AND SP09<='" & txt1(7) & "' "
    End If
    If Len(Trim(txt1(6))) <> 0 Or Len(Trim(txt1(7))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7)  'Add By Sindy 2010/10/22
    End If
    Process1
'選擇發文日期或本所案號
Else
    strSQL1 = ""
    strSQL2 = ""
    StrSQL6 = ""
    If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0).Caption & txt1(0) 'Add By Sindy 2010/10/22
    End If
    StrSQL6 = ""
   'Modify By Cheng 2002/09/17
   If Me.Option1(1).Value Then
      If Len(Trim(txt1(4))) <> 0 Then
        StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(4))) & " "
      End If
      If Len(Trim(txt1(5))) <> 0 Then
        StrSQL6 = StrSQL6 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(5))) & " "
      End If
      If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(4) & "-" & txt1(5)  'Add By Sindy 2010/10/22
      End If
   Else
      StrSQL6 = StrSQL6 & " AND " & ChgCaseprogress(Me.txt1(9).Text & Me.txt1(10).Text & Me.txt1(11).Text & Me.txt1(12).Text) & " "
      pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & txt1(9) & "-" & txt1(10) & "-" & txt1(11) & "-" & txt1(12) 'Add By Sindy 2010/10/22
   End If
   'Modify By Cheng 2002/09/17
'    strsql6 = strsql6 & " AND NP07=305 AND (NP06 IS NULL OR NP06='') "
    strSQL1 = strSQL1 + " AND (TM29 IS NULL OR TM29='') "
    strSQL2 = strSQL2 + " AND (SP15 IS NULL OR SP15='') "
    If Len(txt1(6)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10>='" & txt1(6) & "' "
        strSQL2 = strSQL2 + " AND SP09>='" & txt1(6) & "' "
    End If
    If Len(txt1(7)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10<='" & txt1(7) & "' "
        strSQL2 = strSQL2 + " AND SP09<='" & txt1(7) & "' "
    End If
    If Len(Trim(txt1(6))) <> 0 Or Len(Trim(txt1(7))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7)  'Add By Sindy 2010/10/22
    End If
   'Add By Cheng 2002/09/17
   '2009/11/23 MODIFY BY SONIA CFT-008417
   'StrSQL6 = StrSQL6 & " AND CP27 IS NOT NULL AND CP24 IS NULL AND CP57 IS NULL AND CP09 <'C' "
   StrSQL6 = StrSQL6 & " AND CP27 IS NOT NULL AND (CP24 IS NULL or cp24<>'1') AND CP57 IS NULL AND CP09 <'C' "
   '2009/11/23 end
       
    Process2
End If
If BolProcess = False Then
   Exit Sub
End If
PrintData
End Sub

Sub Process1()
'Add By Cheng 2002/09/17
Dim blnUpdate As Boolean '是否要更新下一程序的期限
'add  by nickc 2006/11/14
'加入印催審的暫緩或取消
Dim Is1705or310 As Boolean
BolProcess = True
'組字串    NP--CP AND (TM OR SP)
'strSQL = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL1 & StrSQL6
'strSQL = strSQL + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL2 & StrSQL6
strSql = "SELECT NP08,CP27,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND TM10=NA01(+) " & strSQL1 & StrSQL6
strSql = strSql + " union all select NP08,CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SP09=NA01(+)" & strSQL2 & StrSQL6
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
         'Add By Cheng 2002/09/17
         '若選擇催審期限, 加詢問使用者是否要更新
         If MsgBox("是否再管制三個月?", vbExclamation + vbYesNo) = vbYes Then
            blnUpdate = True
         Else
            blnUpdate = False
         End If
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/22
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            TestOk = True
            'add by nickc 2006/11/14 檢查有無暫緩或取消催審
            Is1705or310 = False
            strSql = "SELECT * FROM CASEPROGRESS WHERE CP43='" & CheckStr(.Fields("np01")) & "' and cp10 in ('1705','310') "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                Is1705or310 = True
            End If
'2009/11/23 modify by sonia 2009/06/22於審查報告來函時,以該法定期限加1年更新至申請的催審期限,故此處不必再更新
'            strSQL = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='1202' or cp10='1201') AND CP09>'C' "
'            CheckOC2
'            adoRecordset1.CursorLocation = adUseClient
'            adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            '假如   c類 且案件性質 為    1201,1202
'            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                strSQL = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201' or cp10='203' or cp10='301' or cp10='302') AND CP09<'C'  "
'                CheckOC2
'                adoRecordset1.CursorLocation = adUseClient
'                adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    strSQL = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201' or cp10='203' or cp10='301' or cp10='302') AND CP09<'C'  AND (CP27='' OR CP27 IS NULL) "
'                     CheckOC2
'                    adoRecordset1.CursorLocation = adUseClient
'                    adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                        '印資料  UPDATE 本所期限，法定期限
'                        strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
'                        strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
'                        'edit by nickc 2006/11/14
'                        'strSQL = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "') "
'                        strSQL = "INSERT INTO R030406 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "') "
'                        cnnConnection.Execute strSQL
'                        If Len(CheckStr(.Fields(1))) <> 8 Then
'                            SavDay(0) = CheckStr(.Fields(1))
'                        Else
'                            SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                        End If
'                        If Len(CheckStr(.Fields(13))) <> 8 Then
'                            SavDay(1) = CheckStr(.Fields(13))
'                        Else
'                            SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                        End If
'                        'Modify By Cheng 2002/09/17
'                        If blnUpdate Then
'                           cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
'                        End If
'                    Else
'                        'UPDATE 本所期限，法定期限
'                        If Len(CheckStr(.Fields(1))) <> 8 Then
'                            SavDay(0) = CheckStr(.Fields(1))
'                        Else
'                            SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                        End If
'                        If Len(CheckStr(.Fields(13))) <> 8 Then
'                            SavDay(1) = CheckStr(.Fields(13))
'                        Else
'                            SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                        End If
'                        'Modify By Cheng 2002/09/17
'                        If blnUpdate Then
'                           cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
'                        End If
'                    End If
'                Else
'                'UPDATE 本所期限，法定期限
'                    If Len(CheckStr(.Fields(1))) <> 8 Then
'                       SavDay(0) = CheckStr(.Fields(1))
'                     Else
'                        SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                    End If
'                    If Len(CheckStr(.Fields(13))) <> 8 Then
'                        SavDay(1) = CheckStr(.Fields(13))
'                    Else
'                        SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                    End If
'                     'Modify By Cheng 2002/09/17
'                     If blnUpdate Then
'                       cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
'                     End If
'                End If
'            Else
                '印資料  UPDATE 本所期限，法定期限
                strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                'edit by nickc 2006/11/14
                'strSQL = "INSERT INTO R030406 VALUES ('" & ChgSQL(CheckStr(.Fields(14))) & "','" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & strUserNum & "') "
                strSql = "INSERT INTO R030406 VALUES ('" & ChgSQL(CheckStr(.Fields(14))) & "','" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
                If Len(CheckStr(.Fields(0))) <> 8 Then
                    SavDay(0) = CheckStr(.Fields(0))
                Else
                    SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(0)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                End If
                If Len(CheckStr(.Fields(13))) <> 8 Then
                    SavDay(1) = CheckStr(.Fields(13))
                Else
                    SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                End If
               'Modify By Cheng 2002/09/17
               If blnUpdate Then
                   cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
               End If
'            End If
'2009/11/23 end
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/22
        ShowNoData
        BolProcess = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
End Sub

Sub Process2()
'Add By Cheng 2002/09/17
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'add  by nickc 2006/11/14
'加入印催審的暫緩或取消
Dim Is1705or310 As Boolean
BolProcess = True
'組字串    CP--NP AND (TM OR SP)
'strSQL = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL1 & StrSQL6
'strSQL = strSQL + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL2 & StrSQL6
'Modify By Cheng 2002/09/17
'strSQL = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CPM03,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND TM10=NA01(+)  " & strsql1 & strsql6
'strSQL = strSQL + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CPM03,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SP09=NA01(+)  " & strsql2 & strsql6
strSql = "SELECT CP27,'',NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CPM03,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,CP09,'','','',NVL(NA03,NA04),CP01,TM10,CP10 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND TM10=NA01(+)  " & strSQL1 & StrSQL6
strSql = strSql + " union all select CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CPM03,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,CP09,'','','',NVL(NA03,NA04),CP01,SP09,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SP09=NA01(+)  " & strSQL2 & StrSQL6
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/22
        .MoveFirst
        DoEvents
        Do While .EOF = False
            'Add By Cheng 2002/09/17
            '若選擇發文日期或本所期限時, 檢查案件國家檔的CF05, 若無資料或CF05 IS NULL,則不管制
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            StrSQLa = "Select CF05 FROM CASEFEE WHERE CF01='" & .Fields(15).Value & "' AND CF02='" & .Fields(16).Value & "' AND CF03='" & .Fields(17) & "' AND CF05 IS NOT NULL "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount <= 0 Then
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               GoTo NextRecord
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            TestOk = True
            'add by nickc 2006/11/14 檢查有無暫緩或取消催審
            Is1705or310 = False
            strSql = "SELECT * FROM CASEPROGRESS WHERE CP43='" & CheckStr(.Fields("cp09")) & "' and cp10 in ('1705','310') "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                Is1705or310 = True
            End If
            '假如   c類 且案件性質 為    1201,1202
            strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='1202' or cp01='1201') AND CP09>'C' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201' or cp10='202' or cp10='301' or cp10='302') AND CP09<'C'  "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201' or cp10='202' or cp10='301' or cp10='302') AND CP09<'C' AND (CP27='' OR CP27 IS NULL) "
                    CheckOC2
                    adoRecordset1.CursorLocation = adUseClient
                    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                        '印資料  UPDATE 本所期限，法定期限
                            strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                            strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                            'edit by nickc 2006/11/14
                            'strSQL = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "') "
                            strSql = "INSERT INTO R030406 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "') "
                            cnnConnection.Execute strSql
                           'Modify By Cheng 2002/09/17
                           '取消更新下一程序期限
'                            If Len(CheckStr(.Fields(1))) <> 8 Then
'                                SavDay(0) = CheckStr(.Fields(1))
'                            Else
'                                SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                            End If
'                            If Len(CheckStr(.Fields(13))) <> 8 Then
'                                SavDay(1) = CheckStr(.Fields(13))
'                            Else
'                                SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                            End If
'                            cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                    Else
                    'UPDATE 本所期限，法定期限
                        'Modify By Cheng 2002/09/17
                        '取消更新下一程序期限
'                        If Len(CheckStr(.Fields(1))) <> 8 Then
'                            SavDay(0) = CheckStr(.Fields(1))
'                        Else
'                            SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                        End If
'                        If Len(CheckStr(.Fields(13))) <> 8 Then
'                            SavDay(1) = CheckStr(.Fields(13))
'                        Else
'                            SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                        End If
'                        cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                    End If
                Else
                'UPDATE 本所期限，法定期限
                  'Modify By Cheng 2002/09/17
                  '取消更新下一程序期限
'                    If Len(CheckStr(.Fields(1))) <> 8 Then
'                        SavDay(0) = CheckStr(.Fields(1))
'                    Else
'                        SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                    End If
'                    If Len(CheckStr(.Fields(13))) <> 8 Then
'                        SavDay(1) = CheckStr(.Fields(13))
'                    Else
'                        SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                    End If
'                    cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                End If
            Else
            '印資料  UPDATE 本所期限，法定期限
                strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                'edit by nickc 2006/11/14
                'strSQL = "INSERT INTO R030406 VALUES ('" & ChgSQL(CheckStr(.Fields(14))) & "','" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & strUserNum & "') "
                strSql = "INSERT INTO R030406 VALUES ('" & ChgSQL(CheckStr(.Fields(14))) & "','" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
               'Modify By Cheng 2002/09/17
               '取消更新下一程序期限
'                If Len(CheckStr(.Fields(1))) <> 8 Then
'                    SavDay(0) = CheckStr(.Fields(1))
'                Else
'                    SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                End If
'                If Len(CheckStr(.Fields(13))) <> 8 Then
'                    SavDay(1) = CheckStr(.Fields(13))
'                Else
'                    SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                End If
'                cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
            End If


            CheckOC2
NextRecord:
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/22
        ShowNoData
        BolProcess = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "外商催審表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "催審期限：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Else
    Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(4)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & txt1(6) & "－" & txt1(7)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "催審期限"
Else
    Printer.Print "發文日"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "發文日"
Else
    Printer.Print "催審期限"
End If
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "申請案號/審定號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil()
For i = 1 To 9
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 2000
PLeft(3) = 3200
PLeft(4) = 6000
PLeft(5) = 7700
PLeft(6) = 11000
PLeft(7) = 12000
PLeft(8) = 14000
PLeft(9) = 15000
End Sub

Sub PrintData()
strSql = "SELECT * FROM R030406 WHERE ID='" & strUserNum & "' ORDER BY R095002,R095003,R095005 "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(3) = StrToStr(strTemp(3), 14)
            strTemp(5) = StrToStr(strTemp(5), 16)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = StrToStr(strTemp(7), 9)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(9) = StrToStr(strTemp(9), 4)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    Else
        Exit Sub
    End If
End With
Printer.EndDoc
CheckOC
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
'Add By Cheng 2002/09/17
m_strUserRight = Me.txt1(0).Text
If m_strUserRight <> "" Then
   m_arrUserRight = Split(m_strUserRight, ",")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm030406 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
     txt1(2).SetFocus
     txt1_GotFocus (2)
Case 1
     txt1(4).SetFocus
     txt1_GotFocus (4)
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)

Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 8
     Select Case Trim(txt1(8))
     Case "1", "2", ""
     Case Else
        s = MsgBox("列印別只能輸入1或2！", , "錯誤！")
        txt1(Index).SetFocus
        txt1_GotFocus (Index)
        Exit Sub
     End Select
Case 7
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 2, 3, 4, 5
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 3 Or Index = 5 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
'Add By Cheng 2002/09/17
Case 9
   blnUserRight = False
   If Me.txt1(9).Text <> "" Then
      If m_strUserRight <> "" Then
         For ii = LBound(m_arrUserRight) To UBound(m_arrUserRight)
            If m_arrUserRight(ii) = Me.txt1(9).Text Then
               blnUserRight = True
            End If
         Next ii
         If blnUserRight = False Then
            MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txt1(9).SetFocus
            txt1_GotFocus 9
            Exit Sub
         End If
      Else
         MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(9).SetFocus
         txt1_GotFocus 9
         Exit Sub
      End If
   End If
Case Else
End Select
End Sub

