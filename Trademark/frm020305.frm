VERSION 5.00
Begin VB.Form frm020305 
   BorderStyle     =   1  '單線固定
   Caption         =   "催審函/催審表"
   ClientHeight    =   3780
   ClientLeft      =   1920
   ClientTop       =   2310
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3975
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   14
      Left            =   2430
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2670
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   13
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2670
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   12
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2340
      Width           =   360
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   11
      Left            =   2940
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2340
      Width           =   225
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   10
      Left            =   1890
      MaxLength       =   6
      TabIndex        =   6
      Top             =   2340
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   9
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2340
      Width           =   435
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   22
      Top             =   2370
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "催審期限："
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   17
      Top             =   1755
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "發文日期："
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   16
      Top             =   2070
      Width           =   1200
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   1380
      Width           =   2055
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1740
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   2445
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1740
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2040
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   5
      Left            =   2445
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2040
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   8
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3300
      Width           =   300
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   6
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2985
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   7
      Left            =   2445
      MaxLength       =   4
      TabIndex        =   12
      Top             =   2985
      Width           =   990
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1968
      TabIndex        =   14
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2760
      TabIndex        =   15
      Top             =   48
      Width           =   756
   End
   Begin VB.Label Label1 
      Caption         =   "1.管制表已改午夜通知承辦人2.定稿改至案件催審作業產生3.此程式程式有誤!!!!!!!!!"
      ForeColor       =   &H000000FF&
      Height          =   540
      Index           =   4
      Left            =   360
      TabIndex        =   24
      Top             =   600
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   1770
      X2              =   2970
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   2700
      Width           =   915
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1650
      X2              =   3470
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label1 
      Caption         =   "(1. 管制表  2.定稿)"
      Height          =   180
      Index           =   5
      Left            =   1770
      TabIndex        =   21
      Top             =   3330
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   20
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   19
      Top             =   3330
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   165
      TabIndex        =   18
      Top             =   2985
      Width           =   915
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1815
      X2              =   3015
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1830
      X2              =   3030
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1875
      X2              =   3075
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "frm020305"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay(0 To 1) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String
'Add By Cheng 2002/09/17
Dim m_strUserRight As String '使用者系統類別使用權限
Dim m_arrUserRight '使用者系統類別使用權限陣列
Dim ii As Integer '回圈序號
Dim blnUserRight As Boolean '是否有此系統類別權限
Dim boleFileSave As Boolean, m_TM01 As String 'Add By Sindy 2012/1/16


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
     Else
         '選擇催審期限區間
         If Option1(0).Value = True Then
              'Add By Cheng 2002/03/21
              If PUB_CheckKeyInDate(Me.Txt1(2)) = -1 Then
                 Me.Txt1(2).SetFocus
                 txt1_GotFocus 2
                 Exit Sub
              End If
              If PUB_CheckKeyInDate(Me.Txt1(3)) = -1 Then
                 Me.Txt1(3).SetFocus
                 txt1_GotFocus 3
                 Exit Sub
              End If
             
             If Len(Txt1(3)) = 0 Then
                 s = MsgBox("催審期限區間不可空白!!", , "USER 輸入錯誤")
                 Txt1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             End If
         '選擇發文日期區間
         ElseIf Me.Option1(1).Value Then
              'Add By Cheng 2002/03/21
              If PUB_CheckKeyInDate(Me.Txt1(4)) = -1 Then
                 Me.Txt1(4).SetFocus
                 txt1_GotFocus 4
                 Exit Sub
              End If
              If PUB_CheckKeyInDate(Me.Txt1(5)) = -1 Then
                 Me.Txt1(5).SetFocus
                 txt1_GotFocus 5
                 Exit Sub
              End If
             
             If Len(Txt1(5)) = 0 Then
                 s = MsgBox("發文日期區間不可空白!!", , "USER 輸入錯誤")
                 Txt1(4).SetFocus
                 txt1_GotFocus (4)
                 Exit Sub
             End If
         '選擇本所案號
         Else
            'Add By Cheng 2002/09/17
            If Me.Txt1(9).Text = "" Then
               MsgBox "請輸入本所案號的系統類別!!!", vbExclamation + vbOKOnly
               Me.Txt1(9).SetFocus
               txt1_GotFocus 9
               Exit Sub
            End If
            blnUserRight = False
            If Me.Txt1(9).Text <> "" Then
               If m_strUserRight <> "" Then
                  For ii = LBound(m_arrUserRight) To UBound(m_arrUserRight)
                     If m_arrUserRight(ii) = Me.Txt1(9).Text Then
                        blnUserRight = True
                     End If
                  Next ii
                  If blnUserRight = False Then
                     MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
                     Me.Txt1(9).SetFocus
                     txt1_GotFocus 9
                     Exit Sub
                  End If
               Else
                  MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.Txt1(9).SetFocus
                  txt1_GotFocus 9
                  Exit Sub
               End If
            End If
            If Me.Txt1(10).Text = "" Then
               MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
               Me.Txt1(10).SetFocus
               txt1_GotFocus 10
               Exit Sub
            End If
         End If
         
         If Len(Txt1(8)) = 0 Then
             s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
             Txt1(8).SetFocus
             Exit Sub
         Else
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
             If Txt1(8) = "1" Then '列印管制表
                pub_QL05 = pub_QL05 & ";" & Label1(2) & "管制表" 'Add By Sindy 2010/9/30
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                Process
                Me.Enabled = True
                Screen.MousePointer = vbDefault
             Else '列印定稿
                pub_QL05 = pub_QL05 & ";" & Label1(2) & "定稿" 'Add By Sindy 2010/9/30
                ProcessToWord
                Exit Sub
             End If
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub ProcessToWord()
'Add By Cheng 2002/09/17
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

boleFileSave = False 'Add By Sindy 2012/1/16
'選擇催審期限
If Option1(0).Value = True Then
    strSQL1 = ""
    strSQL2 = ""
    StrSQL6 = ""
    If Len(Txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " and NP02 in (" & SQLGrpStr(Txt1(0), 2) & ") "
      strSQL2 = strSQL2 + " and NP02 in (" & SQLGrpStr(Txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/9/30
    End If
    StrSQL6 = ""
    If Len(Txt1(2)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(Txt1(2))) & " "
    End If
    If Len(Txt1(3)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(Txt1(3))) & " "
    End If
    If Len(Txt1(2)) <> 0 Or Len(Txt1(3)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(2) & "-" & Txt1(3) 'Add By Sindy 2010/9/30
    End If
    StrSQL6 = StrSQL6 & " AND NP07=305 AND NP06 IS NULL "
    strSQL1 = strSQL1 + " AND TM29 IS NULL  "
    strSQL2 = strSQL2 + " AND SP15 IS NULL  "
    If Len(Txt1(6)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10>='" & Txt1(6) & "' "
        strSQL2 = strSQL2 + " AND SP09>='" & Txt1(6) & "' "
    End If
    If Len(Txt1(7)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10<='" & Txt1(7) & "' "
        strSQL2 = strSQL2 + " AND SP09<='" & Txt1(7) & "' "
    End If
    If Len(Txt1(6)) <> 0 Or Len(Txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(6) & "-" & Txt1(7) 'Add By Sindy 2010/9/30
    End If
    
    'add by nickc 2006/11/08 加入案件性質範圍
    If Len(Txt1(13)) <> 0 Then
        StrSQL6 = StrSQL6 & " and np07>=" & Txt1(13) & " "
    End If
    If Len(Txt1(14)) <> 0 Then
        StrSQL6 = StrSQL6 & " and np07<=" & Txt1(14) & " "
    End If
    If Len(Txt1(13)) <> 0 Or Len(Txt1(14)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(13) & "-" & Txt1(14) 'Add By Sindy 2010/9/30
    End If
    
   strSql = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,TM10,CP01,CP02,CP03,CP04 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
   strSql = strSql + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,SP09,CP01,CP02,CP03,CP04 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
           .MoveFirst
           DoEvents
           Do While .EOF = False
               For i = 0 To 9
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               m_TM01 = CheckStr(.Fields("CP01")) 'Add By Sindy 2012/1/16
               TestOk = True
               strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='1202' or cp10='1201') AND CP09>'C' "
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               '假如   c類 且案件性質 為    1201,1202
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201' or cp10='301' or cp10='302') AND CP09<'C'  "
                   CheckOC2
                   adoRecordset1.CursorLocation = adUseClient
                   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                       strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201'  or cp10='301' or cp10='302') AND CP09<'C' AND (CP27='' OR CP27 IS NULL) "
                        CheckOC2
                       adoRecordset1.CursorLocation = adUseClient
                       adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                       'MODIFY BY SONIA 2016/3/11 若有"通知修正"之C類收文但無"修正"之A、B類收文，或有"修正" 之A、B類收文但無發文日者不印
                       If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                           '印資料  UPDATE 本所期限，法定期限
                           'strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                           'strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                           'StrSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                           'cnnConnection.Execute StrSQL
                           'Modify By Sindy 2012/1/16
                           'PrintLetter "", "", CheckStr(.Fields(14)), "", "", CheckStr(.Fields(10))
                           PrintLetter "", CheckStr(.Fields("CP01")), CheckStr(.Fields(14)), "", "", CheckStr(.Fields(10)), CheckStr(.Fields("CP02")), CheckStr(.Fields("CP03")), CheckStr(.Fields("CP04"))
                           '2012/1/16 End
                           'Modify By Cheng 2002/09/17
                           '催審管制表才要更新下一期限
'                           '900803  邱小姐說  只有催審定稿才 UPDATE
'                           If Len(CheckStr(.Fields(1))) <> 8 Then
'                               SavDay(0) = CheckStr(.Fields(1))
'                           Else
'                               SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                           End If
'                           If Len(CheckStr(.Fields(13))) <> 8 Then
'                               SavDay(1) = CheckStr(.Fields(13))
'                           Else
'                               SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                           End If
'                           cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                       Else
                           'UPDATE 本所期限，法定期限
                           'Modify By Cheng 2002/09/17
                           '催審管制表才要更新下一期限
'                           '900803  邱小姐說  只有催審定稿才 UPDATE
'                           If Len(CheckStr(.Fields(1))) <> 8 Then
'                               SavDay(0) = CheckStr(.Fields(1))
'                           Else
'                               SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                           End If
'                           If Len(CheckStr(.Fields(13))) <> 8 Then
'                               SavDay(1) = CheckStr(.Fields(13))
'                           Else
'                               SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                           End If
'                           cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                       End If
                   Else
                   'UPDATE 本所期限，法定期限
                        'Modify By Cheng 2002/09/17
                        '催審管制表才要更新下一期限
'                       '900803  邱小姐說  只有催審定稿才 UPDATE
'                       If Len(CheckStr(.Fields(1))) <> 8 Then
'                          SavDay(0) = CheckStr(.Fields(1))
'                        Else
'                           SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                       End If
'                       If Len(CheckStr(.Fields(13))) <> 8 Then
'                           SavDay(1) = CheckStr(.Fields(13))
'                       Else
'                           SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                       End If
'                       cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                   End If
               Else
                   '印資料  UPDATE 本所期限，法定期限
                   'strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                   'strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                   'StrSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                   'cnnConnection.Execute StrSQL
                   'Modify By Sindy 2012/1/16
                   'PrintLetter "", "", CheckStr(.Fields(14)), "", "", CheckStr(.Fields(10))
                   PrintLetter "", CheckStr(.Fields("CP01")), CheckStr(.Fields(14)), "", "", CheckStr(.Fields(10)), CheckStr(.Fields("CP02")), CheckStr(.Fields("CP03")), CheckStr(.Fields("CP04"))
                   '2012/1/16 End
                  'Modify By Cheng 2002/09/17
                  '催審管制表才要更新下一期限
'                   '900803  邱小姐說  只有催審定稿才 UPDATE
'                   If Len(CheckStr(.Fields(1))) <> 8 Then
'                       SavDay(0) = CheckStr(.Fields(1))
'                   Else
'                       SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                   End If
'                   If Len(CheckStr(.Fields(13))) <> 8 Then
'                       SavDay(1) = CheckStr(.Fields(13))
'                   Else
'                       SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                   End If
'                   cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
               End If
               .MoveNext
               DoEvents
           Loop
       Else
           InsertQueryLog (0) 'Add By Sindy 2010/9/30
           ShowNoData
          Exit Sub
       End If
   End With
   CheckOC
   
   'Add By Sindy 2012/1/16
   If boleFileSave = True Then
      MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
   End If
   '2012/1/16 End
   
   Screen.MousePointer = vbDefault
   Exit Sub
'選擇發文日期或本所期限
Else
    strSQL1 = ""
    strSQL2 = ""
    StrSQL6 = ""
    If Len(Txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(Txt1(0), 2) & ") "
      strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(Txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/9/30
    End If
    StrSQL6 = ""
   'Modify By Cheng 2002/09/17
   If Me.Option1(1).Value Then
      If Len(Txt1(4)) <> 0 Then
           StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(Txt1(4))) & ""
      End If
      If Len(Trim(Txt1(5))) <> 0 Then
        StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(Txt1(5))) & " "
      End If
      If Len(Txt1(4)) <> 0 Or Len(Txt1(5)) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Txt1(4) & "-" & Txt1(5) 'Add By Sindy 2010/9/30
      End If
   Else
      StrSQL6 = StrSQL6 + " AND " & ChgCaseprogress(Me.Txt1(9).Text & Me.Txt1(10).Text & Me.Txt1(11).Text & Me.Txt1(12).Text) & " "
      pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & Txt1(9) & "-" & Txt1(10) & "-" & Txt1(11) & "-" & Txt1(12) 'Add By Sindy 2010/9/30
   End If
   'Modify By Cheng 2002/09/17
'    strsql6 = strsql6 & " AND NP07=305 AND NP06 IS NULL "
    strSQL1 = strSQL1 + " AND TM29 IS NULL "
    strSQL2 = strSQL2 + " AND SP15 IS NULL "
    If Len(Txt1(6)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10>='" & Txt1(6) & "' "
        strSQL2 = strSQL2 + " AND SP09>='" & Txt1(6) & "' "
    End If
    If Len(Txt1(7)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10<='" & Txt1(7) & "' "
        strSQL2 = strSQL2 + " AND SP09<='" & Txt1(7) & "' "
    End If
    If Len(Txt1(6)) <> 0 Or Len(Txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(6) & "-" & Txt1(7) 'Add By Sindy 2010/9/30
    End If
    
    'add by nickc 2006/11/08 加入案件性質範圍
    If Len(Txt1(13)) <> 0 Then
        StrSQL6 = StrSQL6 & " and cp10>='" & Txt1(13) & "' "
    End If
    If Len(Txt1(14)) <> 0 Then
        StrSQL6 = StrSQL6 & " and cp10<='" & Txt1(14) & "' "
    End If
    If Len(Txt1(13)) <> 0 Or Len(Txt1(14)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(13) & "-" & Txt1(14) 'Add By Sindy 2010/9/30
    End If
    
   'Add By Cheng 2002/09/17
   StrSQL6 = StrSQL6 & " AND CP27 IS NOT NULL AND CP24 IS NULL AND CP57 IS NULL AND CP09 <'C' "
    
    '組字串    CP--NP AND (TM OR SP)
   'Modify By Cheng 2002/09/17
'      strSQL = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,TM10 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strsql1 & strsql6
'      strSQL = strSQL + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,SP09 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strsql2 & strsql6
      strSql = "SELECT CP27,'',NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,CP09,'','','',TM10,CP01,TM10,CP10,CP02,CP03,CP04 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
      strSql = strSql + " union all select CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,CP09,'','','',SP09,CP01,SP09,CP10,CP02,CP03,CP04 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
          .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
              InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
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
                  m_TM01 = CheckStr(.Fields("CP01")) 'Add By Sindy 2012/1/16
                  TestOk = True
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
                                  'strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                                  'strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                                  'StrSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                                  'cnnConnection.Execute StrSQL
                                  PrintLetter "", CheckStr(.Fields("CP01")), CheckStr(.Fields(14)), "", "", CheckStr(.Fields(10)), CheckStr(.Fields("CP02")), CheckStr(.Fields("CP03")), CheckStr(.Fields("CP04"))
                                 'Modify By Cheng 2002/09/17
                                 '催審管制表才要更新下一期限
'                                  '900803  邱小姐說  只有催審定稿才 UPDATE
'                                  If Len(CheckStr(.Fields(1))) <> 8 Then
'                                      SavDay(0) = CheckStr(.Fields(1))
'                                  Else
'                                      SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                                  End If
'                                  If Len(CheckStr(.Fields(13))) <> 8 Then
'                                      SavDay(1) = CheckStr(.Fields(13))
'                                  Else
'                                      SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                                  End If
'                                  cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                          Else
                          'UPDATE 本所期限，法定期限
                              'Modify By Cheng 2002/09/17
                              '催審管制表才要更新下一期限
'                              '900803  邱小姐說  只有催審定稿才 UPDATE
'                              If Len(CheckStr(.Fields(1))) <> 8 Then
'                                  SavDay(0) = CheckStr(.Fields(1))
'                              Else
'                                  SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                              End If
'                              If Len(CheckStr(.Fields(13))) <> 8 Then
'                                  SavDay(1) = CheckStr(.Fields(13))
'                              Else
'                                  SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                              End If
'                              cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                          End If
                      Else
                      'UPDATE 本所期限，法定期限
                           'Modify By Cheng 2002/09/17
                           '催審管制表才要更新下一期限
'                          '900803  邱小姐說  只有催審定稿才 UPDATE
'                          If Len(CheckStr(.Fields(1))) <> 8 Then
'                              SavDay(0) = CheckStr(.Fields(1))
'                          Else
'                              SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                          End If
'                          If Len(CheckStr(.Fields(13))) <> 8 Then
'                              SavDay(1) = CheckStr(.Fields(13))
'                          Else
'                              SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                          End If
'                          cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                      End If
                  Else
                  '印資料  UPDATE 本所期限，法定期限
                      'strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                      'strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                      'StrSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                      'cnnConnection.Execute StrSQL
                      PrintLetter "", CheckStr(.Fields("CP01")), CheckStr(.Fields(14)), "", "", CheckStr(.Fields(10)), CheckStr(.Fields("CP02")), CheckStr(.Fields("CP03")), CheckStr(.Fields("CP04"))
                     'Modify By Cheng 2002/09/17
                     '催審管制表才要更新下一期限
'                      '900803  邱小姐說  只有催審定稿才 UPDATE
'                      If Len(CheckStr(.Fields(1))) <> 8 Then
'                          SavDay(0) = CheckStr(.Fields(1))
'                      Else
'                          SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                      End If
'                      If Len(CheckStr(.Fields(13))) <> 8 Then
'                          SavDay(1) = CheckStr(.Fields(13))
'                      Else
'                          SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
'                      End If
'                      cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                  End If
                  CheckOC2
NextRecord:
                  .MoveNext
                  DoEvents
              Loop
          Else
              InsertQueryLog (0) 'Add By Sindy 2010/9/30
              ShowNoData
              Exit Sub
          End If
      End With
      CheckOC
      
      'Add By Sindy 2012/1/16
      If boleFileSave = True Then
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      End If
      '2012/1/16 End
   
      Screen.MousePointer = vbDefault
      Exit Sub
End If
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020305 WHERE ID='" & strUserNum & "' "
If Option1(0).Value = True Then '選擇催審期限
    strSQL1 = ""
    strSQL2 = ""
    StrSQL6 = ""
    If Len(Txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " and NP02 in (" & SQLGrpStr(Txt1(0), 2) & ") "
      strSQL2 = strSQL2 + " and NP02 in (" & SQLGrpStr(Txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/9/30
    End If
    StrSQL6 = ""
    If Len(Txt1(2)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(Txt1(2))) & " "
    End If
    If Len(Txt1(3)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(Txt1(3))) & " "
    End If
    If Len(Txt1(2)) <> 0 Or Len(Txt1(3)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(2) & "-" & Txt1(3) 'Add By Sindy 2010/9/30
    End If
    StrSQL6 = StrSQL6 & " AND NP07=305 AND NP06 IS NULL "
    strSQL1 = strSQL1 + " AND TM29 IS NULL  "
    strSQL2 = strSQL2 + " AND SP15 IS NULL  "
    If Len(Txt1(6)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10>='" & Txt1(6) & "' "
        strSQL2 = strSQL2 + " AND SP09>='" & Txt1(6) & "' "
    End If
    If Len(Txt1(7)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10<='" & Txt1(7) & "' "
        strSQL2 = strSQL2 + " AND SP09<='" & Txt1(7) & "' "
    End If
    If Len(Txt1(6)) <> 0 Or Len(Txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(6) & "-" & Txt1(7) 'Add By Sindy 2010/9/30
    End If
    
    'add by nickc 2006/11/08 加入案件性質範圍
    If Len(Txt1(13)) <> 0 Then
        StrSQL6 = StrSQL6 & " and cp10>=" & Txt1(13) & " "
    End If
    If Len(Txt1(14)) <> 0 Then
        StrSQL6 = StrSQL6 & " and cp10<=" & Txt1(14) & " "
    End If
    If Len(Txt1(13)) <> 0 Or Len(Txt1(14)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(13) & "-" & Txt1(14) 'Add By Sindy 2010/9/30
    End If
    
    'Process1
    If Process1 = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
Else '選擇發文日期
    strSQL1 = ""
    strSQL2 = ""
    StrSQL6 = ""
    If Len(Txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(Txt1(0), 2) & ") "
      strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(Txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/9/30
    End If
    StrSQL6 = ""
   'Modify By Cheng 2002/09/17
   If Me.Option1(1).Value Then
      If Len(Txt1(4)) <> 0 Then
           StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(Txt1(4))) & ""
      End If
      If Len(Trim(Txt1(5))) <> 0 Then
        StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(Txt1(5))) & " "
      End If
      If Len(Txt1(4)) <> 0 Or Len(Txt1(5)) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Txt1(4) & "-" & Txt1(5) 'Add By Sindy 2010/9/30
      End If
   Else
      StrSQL6 = StrSQL6 & " AND " & ChgCaseprogress(Me.Txt1(9).Text & Me.Txt1(10).Text & Me.Txt1(11).Text & Me.Txt1(12)) & " "
      pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & Txt1(9) & "-" & Txt1(10) & "-" & Txt1(11) & "-" & Txt1(12) 'Add By Sindy 2010/9/30
   End If
   'Modify By Cheng 2002/09/17
'    strsql6 = strsql6 & " AND NP07=305 AND NP06 IS NULL "
    strSQL1 = strSQL1 + " AND TM29 IS NULL "
    strSQL2 = strSQL2 + " AND SP15 IS NULL "
    If Len(Txt1(6)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10>='" & Txt1(6) & "' "
        strSQL2 = strSQL2 + " AND SP09>='" & Txt1(6) & "' "
    End If
    If Len(Txt1(7)) <> 0 Then
        strSQL1 = strSQL1 + " AND TM10<='" & Txt1(7) & "' "
        strSQL2 = strSQL2 + " AND SP09<='" & Txt1(7) & "' "
    End If
    If Len(Txt1(6)) <> 0 Or Len(Txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(6) & "-" & Txt1(7) 'Add By Sindy 2010/9/30
    End If
    'add by nickc 2006/11/08 加入案件性質範圍
    If Len(Txt1(13)) <> 0 Then
        StrSQL6 = StrSQL6 & " and cp10>='" & Txt1(13) & "' "
    End If
    If Len(Txt1(14)) <> 0 Then
        StrSQL6 = StrSQL6 & " and cp10<='" & Txt1(14) & "' "
    End If
    If Len(Txt1(13)) <> 0 Or Len(Txt1(14)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(13) & "-" & Txt1(14) 'Add By Sindy 2010/9/30
    End If
    
    'Add By Cheng 2002/09/17
    StrSQL6 = StrSQL6 & " AND CP27 IS NOT NULL AND CP24 IS NULL AND CP57 IS NULL AND CP09 <'C' "
    
    'Process2
    If Process2 = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
End If
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Function Process1() As Boolean
'Add By Cheng 2002/09/17
Dim blnUpdate As Boolean '是否要更新下一程序的期限

'add  by nickc 2006/11/14
'加入印催審的暫緩或取消
Dim Is1705or310 As Boolean
'組字串    NP--CP AND (TM OR SP)
'Modify By Cheng 2002/02/19
'strSQL = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
'Modify by Amy 2015//06/22 +只抓 NP10之ST03為 P2 開頭,智權改顯示 CP13
'strSql = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) AS NATIONNAME FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,Nation WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And TM10=NA01(+) " & strSQL1 & StrSQL6
'strSql = strSql + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) AS NATIONNAME FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And SP09=NA01(+) " & strSQL2 & StrSQL6
strSql = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S3.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) AS NATIONNAME FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,Staff S3,CASEPROPERTYMAP,CUSTOMER,Nation WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) And SubStr(S2.st03,1,2)='P2' And CP13=S3.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And TM10=NA01(+) " & strSQL1 & StrSQL6
strSql = strSql + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) AS NATIONNAME FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,Staff S3,CASEPROPERTYMAP,CUSTOMER,NATION WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP14=S1.ST01(+) AND NP10=S2.ST01(+) And SubStr(S2.st03,1,2)='P2' And CP13=S3.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And SP09=NA01(+) " & strSQL2 & StrSQL6
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
         'Add By Cheng 2002/09/17
         '若選擇催審期限, 加詢問使用者是否要更新
         If MsgBox("是否再管制三個月?", vbExclamation + vbYesNo) = vbYes Then
            blnUpdate = True
         Else
            blnUpdate = False
         End If
        
        .MoveFirst

        'frm100.Show
        'frm100.Tag = Trim(Str(.RecordCount)) & "=0"
        DoEvents
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'add by nickc 2006/11/14 檢查有無暫緩或取消催審
            Is1705or310 = False
            strSql = "SELECT * FROM CASEPROGRESS WHERE CP43='" & CheckStr(.Fields("np01")) & "' and cp10 in ('1705','310') "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                Is1705or310 = True
            End If
            'Add By Cheng 2002/02/19
            strTemp(10) = CheckStr(.Fields("NationName").Value)
            TestOk = True
            strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='1202' or cp10='1201') AND CP09>'C' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            '假如   c類 且案件性質 為    1201,1202
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201' or cp10='301' or cp10='302') AND CP09<'C'  "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201'  or cp10='301' or cp10='302') AND CP09<'C' AND (CP27='' OR CP27 IS NULL) "
                     CheckOC2
                    adoRecordset1.CursorLocation = adUseClient
                    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                        '印資料  UPDATE 本所期限，法定期限
                        strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                        strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                        'Modify By Cheng 2002/02/19
'                        strSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                        'edit by nickc 2006/11/14
                        'strSQL = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                        strSql = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                        cnnConnection.Execute strSql
                        'Modify By Cheng 2002/09/17
                        '只有催審管制表才要更新下一程序期限
'                        '900803  邱小姐說  只有催審定稿才 UPDATE
                        If Len(CheckStr(.Fields(1))) <> 8 Then
                            SavDay(0) = CheckStr(.Fields(1))
                        Else
                            SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
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
                    Else
                        'UPDATE 本所期限，法定期限
                        'Modify By Cheng 2002/09/17
                        '只有催審管制表才要更新下一程序期限
'                        '900803  邱小姐說  只有催審定稿才 UPDATE
                        If Len(CheckStr(.Fields(1))) <> 8 Then
                            SavDay(0) = CheckStr(.Fields(1))
                        Else
                            SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
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
                    End If
                Else
                'UPDATE 本所期限，法定期限
                     'Modify By Cheng 2002/09/17
                     '只有催審管制表才要更新下一程序期限
'                    '900803  邱小姐說  只有催審定稿才 UPDATE
                    If Len(CheckStr(.Fields(1))) <> 8 Then
                       SavDay(0) = CheckStr(.Fields(1))
                     Else
                        SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
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
                End If
            Else
                '印資料  UPDATE 本所期限，法定期限
                strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
               'Modify By Cheng 2002/02/19
'                strSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                'edit by nickc 2006/11/14
                'strSQL = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                strSql = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                cnnConnection.Execute strSql
                  'Modify By Cheng 2002/09/17
                  '只有催審管制表才要更新下一程序期限
'                '900803  邱小姐說  只有催審定稿才 UPDATE
                If Len(CheckStr(.Fields(1))) <> 8 Then
                    SavDay(0) = CheckStr(.Fields(1))
                Else
                    SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
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
            End If
            .MoveNext

            'frm100.Tag = Trim(Str(.RecordCount)) & "=" & Trim(Str(k))
            'frm100.StrMenu
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/9/30
        ShowNoData
        Process1 = False
        Exit Function
    End If
End With
Process1 = True
CheckOC
End Function

Function Process2() As Boolean
'Add By Cheng 2002/09/17
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'add  by nickc 2006/11/14
'加入印催審的暫緩或取消
Dim Is1705or310 As Boolean
'組字串    CP--NP AND (TM OR SP)
'Modify By Cheng 2002/02/19
'strSQL = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
'Modify By Cheng 2002/09/17
'strSQL = "SELECT CP27,NP08,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) AS NATIONNAME FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND TM10=NA01(+) " & strsql1 & strsql6
'strSQL = strSQL + " union all select CP27,NP08,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04) AS NATIONNAME FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SP09=NA01(+) " & strsql2 & strsql6
strSql = "SELECT CP27,'',NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,CP09,'','','',NVL(NA03,NA04) AS NATIONNAME,CP01,TM10,CP10 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND TM10=NA01(+) " & strSQL1 & StrSQL6
strSql = strSql + " union all select CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,CP09,'','','',NVL(NA03,NA04) AS NATIONNAME,CP01,SP09,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SP09=NA01(+) " & strSQL2 & StrSQL6
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
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
            'add by nickc 2006/11/14 檢查有無暫緩或取消催審
            Is1705or310 = False
            strSql = "SELECT * FROM CASEPROGRESS WHERE CP43='" & CheckStr(.Fields("cp09")) & "' and cp10 in ('1705','310') "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                Is1705or310 = True
            End If
            strTemp(10) = CheckStr(.Fields("NationName").Value)
            TestOk = True
            '假如   c類 且案件性質 為    1201,1202
            strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='1202' or cp01='1201') AND CP09>'C' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(3), 1) & "' AND CP02='" & SystemNumber(strTemp(3), 2) & "' AND CP03='" & SystemNumber(strTemp(3), 3) & "' AND CP04='" & SystemNumber(strTemp(3), 4) & "' AND (CP10='203' or cp10='201' or cp10='202' or cp10='301' or cp10='302') AND CP09<'C' "
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
                           'Modify By Cheng 2002/02/19
'                            strSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                            'edit by nickc 2006/11/14
                            'strSQL = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                            strSql = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                            cnnConnection.Execute strSql
                            '900803  邱小姐說  只有催審定稿才 UPDATE
                            'If Len(CheckStr(.Fields(1))) <> 8 Then
                            '    SavDay(0) = CheckStr(.Fields(1))
                            'Else
                            '    SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                            'End If
                            'If Len(CheckStr(.Fields(13))) <> 8 Then
                            '    SavDay(1) = CheckStr(.Fields(13))
                            'Else
                            '    SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                            'End If
                            'cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                    Else
                    'UPDATE 本所期限，法定期限
                        '900803  邱小姐說  只有催審定稿才 UPDATE
                        'If Len(CheckStr(.Fields(1))) <> 8 Then
                        '    SavDay(0) = CheckStr(.Fields(1))
                        'Else
                        '    SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                        'End If
                        'If Len(CheckStr(.Fields(13))) <> 8 Then
                        '    SavDay(1) = CheckStr(.Fields(13))
                        'Else
                        '    SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                        'End If
                        'cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                    End If
                Else
                'UPDATE 本所期限，法定期限
                    '900803  邱小姐說  只有催審定稿才 UPDATE
                    'If Len(CheckStr(.Fields(1))) <> 8 Then
                    '    SavDay(0) = CheckStr(.Fields(1))
                    'Else
                    '    SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                    'End If
                    'If Len(CheckStr(.Fields(13))) <> 8 Then
                    '    SavDay(1) = CheckStr(.Fields(13))
                    'Else
                    '    SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                    'End If
                    'cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
                End If
            Else
            '印資料  UPDATE 本所期限，法定期限
                strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
               'Modify By Cheng 2002/02/19
'                strSQL = "INSERT INTO R020305 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & strUserNum & "') "
                'edit by nickc 2006/11/14
                'strSQL = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                strSql = "INSERT INTO R020305 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & IIf(Is1705or310, "△", "") & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & ChgSQL(strTemp(10)) & "') "
                cnnConnection.Execute strSql
                '900803  邱小姐說  只有催審定稿才 UPDATE
                'If Len(CheckStr(.Fields(1))) <> 8 Then
                '    SavDay(0) = CheckStr(.Fields(1))
                'Else
                '    SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                'End If
                'If Len(CheckStr(.Fields(13))) <> 8 Then
                '    SavDay(1) = CheckStr(.Fields(13))
                'Else
                '    SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(13)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                'End If
                'cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(10)) & "' AND NP07=" & Val(CheckStr(.Fields(11))) & " AND NP22=" & Val(CheckStr(.Fields(12)))
            End If

            CheckOC2
NextRecord:
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/9/30
        ShowNoData
        Process2 = False
        Exit Function
    End If
End With
Process2 = True
CheckOC
End Function

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "內商催審函/催審表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "催審期限：" & Format(ChangeTStringToTDateString(Txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(3))
Else
    Printer.Print "發文日：" & Format(ChangeTStringToTDateString(Txt1(4)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(5))
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "催審期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請案號/審定號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "是否出名"
'Add By Cheng 2002/02/19
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil()
'Modify By Cheng 2002/02/19
'For i = 0 To 9
For i = 0 To 10
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1500
PLeft(2) = 2500
PLeft(3) = 4200
PLeft(4) = 6000
PLeft(5) = 8500
PLeft(6) = 9500
PLeft(7) = 12000
PLeft(8) = 13000
PLeft(9) = 14000
'Add By Cheng 2002/02/19
PLeft(10) = 15000
End Sub

Sub PrintData()
'Add by Amy 2015/06/29
Dim bolPage As Boolean '是否分頁
Dim CountRow As Integer

strSql = "SELECT * FROM R020305 WHERE ID='" & strUserNum & "' "
 'Modify by Amy 2015/06/29 +依承辦人分頁
If Option1(0).Value = True Then
    'strSql = strSql + " ORDER BY R056001,R056002,R056004 "
    strSql = strSql + " ORDER BY R056008,R056001,R056002,R056004 "
    bolPage = True
Else
    strSql = strSql + " ORDER BY R056002,R056001,R056004 "
    bolPage = False
End If
'end 2015/06/29
CheckOC
Page = 1: CountRow = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            'Add by Amy 2015/06/29 +承辦人分頁
            If bolPage = True And Page > 1 And (strTemp(7) <> .Fields("R056008") Or CountRow > 25) Then
                Printer.Font.Size = 12
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = iPrint
                Printer.Print "△表示有暫緩審理或取消催審"
                iPrint = iPrint + 300
                Page = Page + 1
                CountRow = 1
                Printer.NewPage
                PrintTitle
            End If
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 8)
            'Modify By Cheng 2002/02/19
'            strTemp(4) = StrToStr(strTemp(4), 14)
            strTemp(4) = StrToStr(strTemp(4), 10)
            strTemp(5) = StrToStr(strTemp(5), 4)
            'Modify By Cheng 2002/02/19
'            strTemp(6) = StrToStr(strTemp(6), 14)
            strTemp(6) = StrToStr(strTemp(6), 10)
            strTemp(7) = StrToStr(strTemp(7), 4)
            strTemp(8) = StrToStr(strTemp(8), 4)
            'Add By Cheng 2002/02/19
            strTemp(10) = StrToStr(.Fields("r056011").Value, 4)
            PrintDatil
            If iPrint >= 10000 Then
                'add by nickc 2007/03/14 加入最後印的說明
                Printer.Font.Size = 12
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = iPrint
                Printer.Print "△表示有暫緩審理或取消催審"
                iPrint = iPrint + 300
                Page = Page + 1
                CountRow = 1 'Add by Amy 2015/06/29
                Printer.NewPage
                PrintTitle
            End If
            CountRow = CountRow + 1  'Add by Amy 2015/06/29
            .MoveNext
        Loop
    Else
        Exit Sub
    End If
End With
'add by nickc 2007/03/14 加入最後印的說明
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "△表示有暫緩審理或取消催審"
iPrint = iPrint + 300


Printer.EndDoc
CheckOC
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Txt1(0) = GetSystemKindByNick
'Add By Cheng 2002/09/17
m_strUserRight = Me.Txt1(0).Text
If m_strUserRight <> "" Then
   m_arrUserRight = Split(m_strUserRight, ",")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm020305 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
      Txt1(2).SetFocus
      txt1_GotFocus (2)
Case 1
      Txt1(4).SetFocus
      txt1_GotFocus (4)
'Add By Cheng 2002/09/17
Case 2
      Txt1(9).SetFocus
      txt1_GotFocus (9)
Case Else
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
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
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(Txt1(0)), ",")
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
            Txt1(0).SetFocus
            Txt1(0).SelStart = 0
            Txt1(0).SelLength = Len(Txt1(0))
            Exit Sub
        End If
     Next i
Case 2, 3, 4, 5
   If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
      Me.Txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 3 Or Index = 5 Then
     If RunNick(Txt1(Index - 1), Txt1(Index)) Then
         Txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case 7
     If RunNick(Txt1(Index - 1), Txt1(Index)) Then
         Txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 8
     Select Case Trim(Txt1(8))
     Case "1", "2", ""
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(8).SetFocus
          Txt1(8).SelStart = 0
          Txt1(8).SelLength = Len(Txt1(8))
          Exit Sub
     End Select
'Add By Cheng 2002/09/17
Case 9
   blnUserRight = False
   If Me.Txt1(9).Text <> "" Then
      If m_strUserRight <> "" Then
         For ii = LBound(m_arrUserRight) To UBound(m_arrUserRight)
            If m_arrUserRight(ii) = Me.Txt1(9).Text Then
               blnUserRight = True
            End If
         Next ii
         If blnUserRight = False Then
            MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.Txt1(9).SetFocus
            txt1_GotFocus 9
            Exit Sub
         End If
      Else
         MsgBox "本所案號的系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.Txt1(9).SetFocus
         txt1_GotFocus 9
         Exit Sub
      End If
   End If
Case Else
End Select
End Sub



Private Sub InsExpField(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, ByVal strDate As String, ByVal strSysDate As String, ByVal strCP09 As String)
   'Dim strSQL As String
   ' 下一程序
   'Dim StrNP07 As Strubg
   ' 系統別
   'Dim strTM01 As String
   ' 申請國家
   'Dim strTM10 As String
   ' 專用期限止日
   'Dim strDate As String
   ' 系統日
   'Dim strSysDate As String
   ' 總收文號
   'Dim strCP09 As String
      ' 延展
    ' 申請國家為台灣
   If strTM10 < "010" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "11", strCP09, "01", strUserNum
   ' 申請國家為大陸
   ElseIf strTM10 = "020" Then
      ' 清除定稿例外欄位檔原有資料
      EndLetter "11", strCP09, "02", strUserNum
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, ByVal strData As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
'Add By Sindy 2012/1/16
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/16 End
   
   ' 下一程序
   'Dim StrNP07 As Strubg
   ' 系統別
   'Dim StrTM01 As String
   ' 申請國家
   'Dim StrTM10 As String
   ' 專用期限止日
   'Dim StrDate As String
   ' 系統日
   'Dim StrSysDate As String
   ' 總收文號
   'Dim StrCP09 As String
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField strNP07, strTM01, strTM10, strDate, strSysDate, strCP09
   
   'Add By Sindy 2012/1/16
   ET01 = "11"
   ET02 = strCP09
   bolEdit = False
   '2012/1/16 End
   
    ' 申請國家為台灣
   If strTM10 < "010" Then
       ' 列印定稿
'       NowPrint strCP09, "11", "01", False, strUserNum, 0
      ET03 = "01" 'Modify By Sindy 2012/1/16

   ' 申請國家為大陸
   ElseIf strTM10 = "020" Then
      ' 列印定稿
'      NowPrint strCP09, "11", "02", False, strUserNum, 0
      ET03 = "02" 'Modify By Sindy 2012/1/16
   End If
   
   'Add By Sindy 2012/1/16
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(strTM01 & strTM02 & strTM03 & strTM04, strNP07 = "102", , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
         boleFileSave = True
'         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      Else
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0
      End If
   End If
   '2012/1/16 End
End Sub
