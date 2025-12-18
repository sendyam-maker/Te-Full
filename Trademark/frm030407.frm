VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030407 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員案件明細表"
   ClientHeight    =   3060
   ClientLeft      =   5040
   ClientTop       =   3300
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3600
   Begin VB.OptionButton opt1 
      Caption         =   "發文"
      Height          =   225
      Index           =   1
      Left            =   1230
      TabIndex        =   18
      Top             =   1470
      Width           =   915
   End
   Begin VB.OptionButton opt1 
      Caption         =   "收文"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   17
      Top             =   1470
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1395
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2100
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1092
      TabIndex        =   0
      Top             =   468
      Width           =   1740
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1092
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1104
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1092
      MaxLength       =   3
      TabIndex        =   1
      Top             =   780
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1896
      MaxLength       =   3
      TabIndex        =   2
      Top             =   780
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1092
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1755
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1896
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1755
      Width           =   705
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1536
      TabIndex        =   7
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2328
      TabIndex        =   8
      Top             =   48
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2505
      Left            =   60
      TabIndex        =   19
      Top             =   3120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4419
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "是否列印明細："
      Height          =   180
      Index           =   8
      Left            =   90
      TabIndex        =   16
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:印)"
      Height          =   180
      Index           =   11
      Left            =   1710
      TabIndex        =   15
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "FC案件系統類別：FCT,CFT,CFC,T,S,L                          業務區：F10-F19"
      Height          =   420
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2490
      Width           =   3315
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   1920
      TabIndex        =   13
      Top             =   1155
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   492
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1152
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   1224
      X2              =   2484
      Y1              =   924
      Y2              =   924
   End
   Begin VB.Line Line2 
      X1              =   1530
      X2              =   2280
      Y1              =   1860
      Y2              =   1860
   End
End
Attribute VB_Name = "frm030407"
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
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, SavDay3 As String
'add by nickc 2007/12/18
Dim StrTemp4(0 To 4) As String, StrTemp5(0 To 4) As String
'Add By Sindy 2010/5/4
Dim intCompRow As Integer
Dim dblPointTot As Double
Dim m_strTemp1 As String, m_strTemp3 As String
Dim dblSum As Double
'2010/5/4 End
Dim dblPointTotSub As Double, dblPointRow As Integer 'Add By Sindy 2012/2/3


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
             'edit by nickc 2008/01/04
             's = MsgBox("收文日期區間不可空白!!", , "USER 輸入錯誤")
             s = MsgBox(IIf(opt1(0).Value = True, "收文", "發文") & "日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(4).SetFocus
             txt1_GotFocus (4)
             Exit Sub
         Else
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
             Process
             Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Dim strPointSql As String 'Add By Sindy 2010/5/4
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R030407 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
strPointSql = "" 'Add By Sindy 2010/5/4
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/22
End If
'edit by nickc 2008/01/04
If opt1(0).Value = True Then
    If Len(Trim(txt1(4))) <> 0 Then
       StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(4))) & " "
       strPointSql = strPointSql + " AND CP05>=" & Val(ChangeTStringToWString(txt1(4))) & " " 'Add By Sindy 2010/5/4
    End If
    If Len(Trim(txt1(5))) <> 0 Then
       StrSQL6 = StrSQL6 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(5))) & " "
       strPointSql = strPointSql + " AND CP05<=" & Val(ChangeTStringToWString(txt1(5))) & " " 'Add By Sindy 2010/5/4
    End If
    If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & opt1(0).Caption & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/22
    End If
Else
    If Len(Trim(txt1(4))) <> 0 Then
       StrSQL6 = StrSQL6 + " AND cp27>=" & Val(ChangeTStringToWString(txt1(4))) & " "
       strPointSql = strPointSql + " AND cp27>=" & Val(ChangeTStringToWString(txt1(4))) & " " 'Add By Sindy 2010/5/4
    End If
    If Len(Trim(txt1(5))) <> 0 Then
       StrSQL6 = StrSQL6 & " AND cp27<=" & Val(ChangeTStringToWString(txt1(5))) & " "
       strPointSql = strPointSql + " AND cp27<=" & Val(ChangeTStringToWString(txt1(5))) & " " 'Add By Sindy 2010/5/4
    End If
    If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & opt1(1).Caption & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/22
    End If
End If
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12>='" & txt1(1) & "' "
    strPointSql = strPointSql + " AND CP12>='" & txt1(1) & "' " 'Add By Sindy 2010/5/4
End If
If Len(txt1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12<='" & txt1(2) & "' "
    strPointSql = strPointSql + " AND CP12<='" & txt1(2) & "' " 'Add By Sindy 2010/5/4
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & "-" & txt1(2)  'Add By Sindy 2010/10/22
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP13='" & txt1(3) & "' "
    strPointSql = strPointSql + " AND a1n04='" & txt1(3) & "' " 'Add By Sindy 2010/5/4
    pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(3) & lbl1 'Add By Sindy 2010/10/22
End If
'91.5.7 MODIFY BY SONIA
'edit by nickc 2008/01/04  start
''strSQL = "SELECT NVL(A0902,A0903),S1.ST01," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM10,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' " & strSQL1 & StrSQL6
''strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),S1.ST01," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),DECODE(SP09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' " & strSQL2 & StrSQL6
''end
'strSQL = "SELECT NVL(A0902,A0903),S1.ST01," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM10,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'C' " & strSQL1 & StrSQL6
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),S1.ST01," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),DECODE(SP09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'C' " & strSQL2 & StrSQL6
'91.5.7 END

'Add By Sindy 2010/5/4
Call QueryPointData(strPointSql)
'2010/5/4 End

'Modify By Sindy 2010/4/27 點數改抓點數分配檔
'strSQL = "SELECT NVL(A0902,A0903),S1.ST01," & SQLDate(IIf(opt1(0).Value = True, "CP05", "CP27")) & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM10,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate(IIf(opt1(0).Value = True, "CP27", "CP05")) & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' " & strSQL1 & StrSQL6
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),S1.ST01," & SQLDate(IIf(opt1(0).Value = True, "CP05", "CP27")) & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),DECODE(SP09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate(IIf(opt1(0).Value = True, "CP27", "CP05")) & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' " & strSQL2 & StrSQL6
strSql = "SELECT NVL(A0902,A0903),S1.ST01," & SQLDate(IIf(opt1(0).Value = True, "CP05", "CP27")) & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM10,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate(IIf(opt1(0).Value = True, "CP27", "CP05")) & ",decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "' " & _
                "FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER,acc1n0,acc1k0 " & _
                "WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 AND a1k01(+)=cp60 AND CP09<'B' " & strSQL1 & StrSQL6
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),S1.ST01," & SQLDate(IIf(opt1(0).Value = True, "CP05", "CP27")) & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),DECODE(SP09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP06") & "," & SQLDate(IIf(opt1(0).Value = True, "CP27", "CP05")) & ",decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "' " & _
                "FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,CUSTOMER,acc1n0,acc1k0 " & _
                "WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP12=A0901(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 AND a1k01(+)=cp60 AND CP09<'B' " & strSQL2 & StrSQL6

cnnConnection.Execute "INSERT INTO R030407 " & strSql
strSql = "SELECT * FROM R030407 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
   'InsertQueryLog (adoRecordset.RecordCount + dblPointTot) 'Add By Sindy 2010/10/22
   InsertQueryLog (adoRecordset.RecordCount + dblPointRow) 'Modify By Sindy 2012/2/3
Else
   'Modify By Sindy 2010/5/6
   'If dblPointTot > 0 Then
   'Modify By Sindy 2012/2/3
   If dblPointRow > 0 Then
      'InsertQueryLog (dblPointTot) 'Add By Sindy 2010/10/22
      InsertQueryLog (dblPointRow) 'Add By Sindy 2012/2/3
      Page = 1
      PrintTitle
      Call PrintEndPoint("", "")
      ShowLine
      PrintEnd1
      Printer.EndDoc
      ShowPrintOk
   '2010/5/6 End
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/10/22
      ShowNoData
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "SELECT R096001,ST02,R096003,R096004,R096005,R096006,R096007,R096008,R096009,R096010,R096011,R096012,R096002 FROM R030407,STAFF WHERE R096002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R096002,R096001,R096003,R096004 "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        'SavDay1 = CheckStr(.Fields(0))
        'SavDay2 = CheckStr(.Fields(1))
        'SavDay3 = CheckStr(.Fields(12))
        'SavDay(0) = CheckStr(.Fields(2))
        'SavDay(1) = CheckStr(.Fields(3))
        PrintTitle
        'PrintTitle1
        SavDay1 = "              "
        SavDay2 = "              "
        SavDay3 = "              "
        SavDay(0) = "              "
        SavDay(1) = "              "
        Call PrintEndPoint("", CheckStr(.Fields(12))) 'Add By Sindy 2010/5/4
        Do While .EOF = False
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(0) <> SavDay1 Or CheckStr(.Fields(12)) <> SavDay3 Then
                If Len(Trim(SavDay1)) <> 0 And Len(Trim(SavDay2)) <> 0 Then
                   ShowLine
                   PrintEnd
                   Call PrintEndPoint(SavDay3, CheckStr(.Fields(12))) 'Add By Sindy 2010/5/4
                   ShowLine
                   iPrint = iPrint + 600
                End If
                SavDay1 = strTemp(0)
                SavDay2 = strTemp(1)
                SavDay3 = CheckStr(.Fields(12))
                SavDay(0) = strTemp(2)
                SavDay(1) = strTemp(3)
                If iPrint >= 10000 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitle
                    PrintTitle1
                Else
                    PrintTitle1
                End If
            Else
                If SavDay(0) = strTemp(2) Then
                    strTemp(2) = ""
                    If strTemp(3) = SavDay(1) Then
                        strTemp(3) = ""
                    End If
                Else
                    SavDay(0) = strTemp(2)
                    SavDay(1) = strTemp(3)
                End If
            End If
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle1
            End If
            'add by nickc 2007/12/18
            If txt1(6) = "Y" Then
                strTemp(4) = StrToStr(strTemp(4), 23)
                strTemp(5) = StrToStr(strTemp(5), 4)
                strTemp(6) = StrToStr(strTemp(6), 4)
                strTemp(7) = StrToStr(strTemp(7), 5)
                strTemp(8) = StrToStr(strTemp(8), 4)
                PrintDatil
            End If
            .MoveNext
        Loop
     End If
End With
ShowLine
PrintEnd
Call PrintEndPoint(SavDay3, "") 'Add By Sindy 2010/5/4
ShowLine
PrintEnd1
Printer.EndDoc
CheckOC
End Sub

Sub PrintEnd()
Call GetPointTotSub(SavDay3) 'Add By Sindy 2012/2/3 取得此人員的分配點數小計
If Len(SavDay1) = 0 Then
    strSql = "SELECT COUNT(*),SUM(R096012) FROM R030407 WHERE ID='" & strUserNum & "' AND (R096001='' or r096001 is null) AND R096002='" & SavDay3 & "' "
Else
    If Len(SavDay2) = 0 Then
        strSql = "SELECT COUNT(*),SUM(R096012) FROM R030407 WHERE ID='" & strUserNum & "' AND R096001='" & SavDay1 & "' AND (R096002='' or r096002 is null ) "
    Else
        If Len(SavDay1) = 0 And Len(SavDay2) = 0 Then
            strSql = "SELECT COUNT(*),SUM(R096012) FROM R030407 WHERE ID='" & strUserNum & "' AND (R096001='' or r096001 is null) AND (R096002='' or r096002 is null) "
        Else
            strSql = "SELECT COUNT(*),SUM(R096012) FROM R030407 WHERE ID='" & strUserNum & "' AND R096001='" & SavDay1 & "' AND R096002='" & SavDay3 & "' "
        End If
    End If
End If
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    'Add By Sindy 2010/5/6
    If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
    End If
    '2010/5/6 End
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "小計"
    'Modify By Sindy 2012/2/3
'    Printer.CurrentX = PLeft(8)
'    Printer.CurrentY = iPrint
'    Printer.Print "件數：" & Format(CheckStr(adoRecordset1.Fields(0)), "###,###,##0")
'    Printer.CurrentX = PLeft(10)
'    Printer.CurrentY = iPrint
'    Printer.Print "點數：" & Format(CheckStr(adoRecordset1.Fields(1)), "###,###,##0.00")
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "件數：" & Format(CheckStr(adoRecordset1.Fields(0)), "###,###,##0")
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = iPrint
    Printer.Print "點數：" & Format(CheckStr(adoRecordset1.Fields(1)), "###,###,##0.00")
    If dblPointTotSub > 0 Then
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "分配點數：" & Format(CheckStr(dblPointTotSub), "###,###,##0.00")
    End If
    '2012/2/3 End
    iPrint = iPrint + 300
End If
CheckOC2
'add by nickc 2007/12/18
ShowLine
If Len(SavDay1) = 0 Then
    strSql = "SELECT R096006,COUNT(*) FROM R030407 WHERE ID='" & strUserNum & "' AND (R096001='' or r096001 is null) AND R096002='" & SavDay3 & "' GROUP BY R096006 "
Else
    If Len(SavDay2) = 0 Then
        strSql = "SELECT R096006,COUNT(*) FROM R030407 WHERE ID='" & strUserNum & "' AND R096001='" & SavDay1 & "' AND (R096002='' or r096002 is null ) GROUP BY R096006 "
    Else
        If Len(SavDay1) = 0 And Len(SavDay2) = 0 Then
            strSql = "SELECT R096006,COUNT(*) FROM R030407 WHERE ID='" & strUserNum & "' AND (R096001='' or r096001 is null) AND (R096002='' or r096002 is null) GROUP BY R096006 "
        Else
            strSql = "SELECT R096006,COUNT(*) FROM R030407 WHERE ID='" & strUserNum & "' AND R096001='" & SavDay1 & "' AND R096002='" & SavDay3 & "' GROUP BY R096006 "
        End If
    End If
End If
   intI = 1
   Set adoRecordset1 = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With adoRecordset1
      .MoveFirst
      Do While .EOF = False
         For i = 0 To 4
            StrTemp4(i) = CheckStr(.Fields(0))
            StrTemp5(i) = CheckStr(.Fields(1))
            .MoveNext
            If .EOF = True Then
                For j = i + 1 To 4
                    StrTemp4(j) = ""
                    StrTemp5(j) = ""
                Next j
                Exit For
            End If
         Next i
         PrintSubTot1
      Loop
      End With
   End If
End Sub

Sub PrintEnd1()
strSql = "SELECT COUNT(*),SUM(R096012) FROM R030407 WHERE ID='" & strUserNum & "' "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    'Add By Sindy 2010/5/6
    If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
    End If
    '2010/5/6 End
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "總計"
    'Modify By Sindy 2012/2/3
'    Printer.CurrentX = PLeft(8)
'    Printer.CurrentY = iPrint
'    Printer.Print "件數：" & Format(CheckStr(adoRecordset1.Fields(0)), "###,###,##0")
'    Printer.CurrentX = PLeft(10)
'    Printer.CurrentY = iPrint
'    Printer.Print "點數：" & Format(CheckStr(adoRecordset1.Fields(1)), "###,###,##0.00")
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "件數：" & Format(CheckStr(adoRecordset1.Fields(0)), "###,###,##0")
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = iPrint
    Printer.Print "點數：" & Format(CheckStr(adoRecordset1.Fields(1)), "###,###,##0.00")
    If dblPointTot > 0 Then
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "分配點數：" & Format(CheckStr(dblPointTot), "###,###,##0.00")
    End If
    '2012/2/3 End
    iPrint = iPrint + 300
End If
CheckOC2
'add by nickc 2007/12/18
'ShowLine
strSql = "SELECT R096006,COUNT(*) FROM R030407 WHERE ID='" & strUserNum & "' GROUP BY R096006 "
   intI = 1
   Set adoRecordset1 = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ShowLine
      With adoRecordset1
      .MoveFirst
      Do While .EOF = False
         For i = 0 To 4
            StrTemp4(i) = CheckStr(.Fields(0))
            StrTemp5(i) = CheckStr(.Fields(1))
            .MoveNext
            If .EOF = True Then
                For j = i + 1 To 4
                    StrTemp4(j) = ""
                    StrTemp5(j) = ""
                Next j
                Exit For
            End If
         Next i
         PrintSubTot1
      Loop
      End With
   End If
   
'   'Add By Sindy 2010/5/5
'   If dblPointTot > 0 Then
'      ShowLine
'      If iPrint >= 10000 Then
'         Page = Page + 1
'         Printer.NewPage
'         PrintTitle
'         PrintTitle1
'      End If
'      Printer.CurrentX = PLeft(8)
'      Printer.CurrentY = iPrint
'      Printer.Print "分配點數"
'      Printer.CurrentX = PLeft(10)
'      Printer.CurrentY = iPrint
'      Printer.Print dblPointTot
'      iPrint = iPrint + 300
'   End If
'   '2010/5/5 End
End Sub

Sub PrintTitle()
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
'add by nickc 2007/12/18
If txt1(6) = "Y" Then
    'edit by nickc 2008/01/04
    'Printer.Print "外商智權人員收文明細表"
    Printer.Print "外商智權人員" & IIf(opt1(0).Value = True, "收文", "發文") & "明細表"
Else
    'edit by nickc 2008/01/04
    'Printer.Print "外商智權人員收文統計表"
    Printer.Print "外商智權人員" & IIf(opt1(0).Value = True, "收文", "發文") & "統計表"
End If
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
'edit by nickc 2008/01/04
'Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(4)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
Printer.Print IIf(opt1(0).Value = True, "收文", "發文") & "日：" & Format(ChangeTStringToTDateString(txt1(4)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
Printer.CurrentX = 0
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
Printer.Font.Size = 10
End Sub

Sub PrintTitle1()
GetPleft
If iPrint >= 10000 Then
   Page = Page + 1
   Printer.NewPage
   PrintTitle
   PrintTitle1
   Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "業務區：" & SavDay1
Printer.CurrentX = 5000
Printer.CurrentY = iPrint
Printer.Print "智權人員：" & SavDay2
iPrint = iPrint + 300
'add by nickc 2007/12/18
If txt1(6) = "Y" Then
    Printer.CurrentX = 0
    Printer.CurrentY = iPrint
    Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
    iPrint = iPrint + 300
    If iPrint >= 10000 Then
        Page = Page + 1
        Printer.NewPage
        PrintTitle
        PrintTitle1
        Exit Sub
    End If
    Printer.Font.Size = 10
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    'edit by nickc 2008/01/04
    'Printer.Print "收文日"
    Printer.Print IIf(opt1(0).Value = True, "收文", "發文") & "日"
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
    Printer.Print "申請國家"
    Printer.CurrentX = PLeft(7)
    Printer.CurrentY = iPrint
    Printer.Print "申請人國籍"
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = iPrint
    Printer.Print "承辦人員"
    Printer.CurrentX = PLeft(9)
    Printer.CurrentY = iPrint
    Printer.Print "本所期限"
    Printer.CurrentX = PLeft(10)
    Printer.CurrentY = iPrint
    'edit by nickc 2008/01/04
    'Printer.Print "發文日"
    Printer.Print IIf(opt1(0).Value = True, "發文", "收文") & "日"
    Printer.CurrentX = PLeft(11)
    Printer.CurrentY = iPrint
    Printer.Print "點數"
    iPrint = iPrint + 300
    Printer.CurrentX = 0
    Printer.CurrentY = iPrint
    Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
    iPrint = iPrint + 300
    If iPrint >= 10000 Then
        Page = Page + 1
        Printer.NewPage
        PrintTitle
        PrintTitle1
        Exit Sub
    End If
End If
End Sub

Sub PrintDatil()
For i = 2 To 10
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
Printer.CurrentX = PLeft(11) + 300 - Printer.TextWidth(strTemp(11))
Printer.CurrentY = iPrint
Printer.Print strTemp(11)
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = 0
PLeft(3) = 1000
PLeft(4) = 3000
PLeft(5) = 8500
PLeft(6) = 9500
PLeft(7) = 10500
PLeft(8) = 11800
PLeft(9) = 13000
PLeft(10) = 14000
PLeft(11) = 15000
End Sub

Sub ShowLine()
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
   iPrint = iPrint + 300
   'Add By Sindy 2010/5/6
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
      PrintTitle1
   End If
   '2010/5/6 End
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm030407 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CmdOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'add by nickc 2007/12/18
If Index = 6 Then
    If KeyAscii <> 89 And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
'edit by nickc 2008/03/05 秀玲說不要控制了
'     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
'     strTemp2 = Split(Replace(UCase(TXT1(0)), ",,", ""), ",")
'     For i = 0 To UBound(strTemp2)
'        s = 0
'        For j = 0 To UBound(strTemp1)
'            If strTemp2(i) = strTemp1(j) Then
'                s = 1
'                Exit For
'            End If
'        Next j
'        If s = 0 Then
'            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
'            TXT1(0).SetFocus
'            TXT1(0).SelStart = 0
'            TXT1(0).SelLength = Len(TXT1(0))
'            Exit Sub
'        End If
'     Next i
Case 2
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 4, 5
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 5 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case 3
     lbl1 = GetPrjSalesNM(txt1(3))
     If Trim(txt1(3)) <> "" Then
        If Trim(lbl1.Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            txt1(3).SetFocus
            txt1_GotFocus (3)
            Exit Sub
        End If
     End If
Case Else
End Select
End Sub

Sub PrintSubTot1()
Dim lngX0 As Long
   
   If txt1(6) = "Y" Then
      lngX0 = PLeft(1)
   Else
      lngX0 = 2000
   End If
   'Add By Sindy 2010/5/6
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
      PrintTitle1
   End If
   '2010/5/6 End
   For j = 0 To 4
      With Printer
         .CurrentX = lngX0 + (j * 2300)
         .CurrentY = iPrint
         Printer.Print StrConv(MidB(StrConv(StrTemp4(j), vbFromUnicode), 1, 14), vbUnicode)
         .CurrentX = lngX0 + ((j + 1) * 2300) - 400 - .TextWidth(StrTemp5(j))
         .CurrentY = iPrint
         Printer.Print StrTemp5(j)
      End With
   Next j
   iPrint = iPrint + 300
End Sub

'Add By Sindy 2010/4/30
'非自己辦理的案件但點數歸屬自己
Private Sub QueryPointData(strConSql As String)
Dim strEmp As String 'Add By Sindy 2012/2/3
   
   With grd1
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      .FormatString = "cp12|a1n04|a1n03|a1n05|cp01|cp02|cp03|cp04|cp10|a1k02"
   End With
   intCompRow = 0
   dblPointRow = 0 'Add By Sindy 2012/2/3
   dblPointTot = 0
   strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,DECODE(TM10,'000',CPM03,CPM04),a1k02,cp13," & IIf(opt1(0).Value = True, "cp05", "cp27") & ",a.st02,' ' 點數小計 " & _
                  "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                  "WHERE cp60>'X' AND cp13 is not null " & _
                  "AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)<>cp13 and a1n05>0 " & _
                  "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                  "AND a1k25 is null " & _
                  "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp13=b.st01(+) " & _
                  "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
                  "AND CP09<'B' AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") " & _
                  "AND cp01=tm01 AND cp02=tm02 AND cp03=tm03 AND cp04=tm04 " & strConSql
   strSql = strSql & " UNION ALL " & _
                  "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,DECODE(SP09,'000',CPM03,CPM04),a1k02,cp13," & IIf(opt1(0).Value = True, "cp05", "cp27") & ",a.st02,' ' 點數小計 " & _
                  "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                  "WHERE cp60>'X' AND cp13 is not null " & _
                  "AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)<>cp13 and a1n05>0 " & _
                  "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                  "AND a1k25 is null " & _
                  "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp13=b.st01(+) " & _
                  "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
                  "AND CP09<'B' AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") " & _
                  "AND cp01=sp01 AND cp02=sp02 AND cp03=sp03 AND cp04=sp04 " & strConSql & _
                  "Order By a1n04,a1k02,cp01,cp02,cp03,cp04 "
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      intCompRow = 1 '代表有分配點數資料,從第1筆開始讀取
      dblPointRow = adoRecordset.RecordCount 'Add By Sindy 2012/2/3
      Set grd1.Recordset = adoRecordset.Clone
      
      'Modify By Sindy 2012/2/3
'      strSql = "select sum(tt) from (" & _
'                     "SELECT sum(a1n05) tt " & _
'                     "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                     "WHERE cp60>'X' AND cp13 is not null " & _
'                     "AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)<>cp13 and a1n05>0 " & _
'                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                     "AND a1k25 is null " & _
'                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp13=b.st01(+) " & _
'                     "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
'                     "AND CP09<'B' AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") " & _
'                     "AND cp01=tm01 AND cp02=tm02 AND cp03=tm03 AND cp04=tm04 " & strConSql
'      strSql = strSql & " UNION ALL " & _
'                     "SELECT sum(a1n05) tt " & _
'                     "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                     "WHERE cp60>'X' AND cp13 is not null " & _
'                     "AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)<>cp13 and a1n05>0 " & _
'                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                     "AND a1k25 is null " & _
'                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp13=b.st01(+) " & _
'                     "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
'                     "AND CP09<'B' AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") " & _
'                     "AND cp01=sp01 AND cp02=sp02 AND cp03=sp03 AND cp04=sp04 " & strConSql & _
'                   ")"
'      intI = 1
'      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         dblPointTot = adoRecordset.Fields(0)
'      End If
      '計算分配點數的小計及總計
      For i = 1 To grd1.Rows - 1
         If strEmp <> grd1.TextMatrix(i, 1) Then
            If strEmp <> "" Then
               For j = 1 To grd1.Rows - 1
                  If grd1.TextMatrix(j, 1) = strEmp Then
                     grd1.TextMatrix(j, 13) = dblPointTotSub
                  End If
               Next j
            End If
            strEmp = grd1.TextMatrix(i, 1)
            dblPointTotSub = 0
         End If
         dblPointTotSub = dblPointTotSub + Val(grd1.TextMatrix(i, 3))
         dblPointTot = dblPointTot + Val(grd1.TextMatrix(i, 3))
      Next i
      If strEmp <> "" Then
         For j = 1 To grd1.Rows - 1
            If grd1.TextMatrix(j, 1) = strEmp Then
               grd1.TextMatrix(j, 13) = dblPointTotSub
            End If
         Next j
      End If
      '2012/2/3 End
   End If
End Sub

'Add By Sindy 2010/5/4
Sub PrintEndPoint(strComp1 As String, strNext1 As String)
Dim ii As Integer
Dim bNextTrue As Boolean
   
   If intCompRow = 0 Or intCompRow > (grd1.Rows - 1) Then Exit Sub
   
   '處理傳進來比對人員的資料
   If grd1.TextMatrix(intCompRow, 1) = strComp1 Then
      m_strTemp1 = strComp1
      m_strTemp3 = grd1.TextMatrix(intCompRow, 12)
      Call ShowLine1(0)
      If iPrint >= 10000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          PrintTitle2
      End If
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print "請款單分配點數："
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print "請款日"
      Printer.CurrentX = 1500
      Printer.CurrentY = iPrint
      Printer.Print "本所案號"
      Printer.CurrentX = 3500
      Printer.CurrentY = iPrint
      Printer.Print "案件性質"
      Printer.CurrentX = 5000 - Printer.TextWidth("點數")
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      iPrint = iPrint + 300
      Call ShowLine1(1)
      dblSum = 0
      For ii = intCompRow To grd1.Rows - 1
         If grd1.TextMatrix(intCompRow, 1) = strComp1 Then
            m_strTemp1 = strComp1
            m_strTemp3 = grd1.TextMatrix(intCompRow, 12)
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle2
            End If
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print ChangeTStringToTDateString(grd1.TextMatrix(intCompRow, 9))
            Printer.CurrentX = 1500
            Printer.CurrentY = iPrint
            Printer.Print grd1.TextMatrix(intCompRow, 4) & "-" & grd1.TextMatrix(intCompRow, 5) & "-" & grd1.TextMatrix(intCompRow, 6) & "-" & grd1.TextMatrix(intCompRow, 7)
            Printer.CurrentX = 3500
            Printer.CurrentY = iPrint
            Printer.Print grd1.TextMatrix(intCompRow, 8)
            Printer.CurrentX = 5000 - Printer.TextWidth(CheckStr(grd1.TextMatrix(intCompRow, 3)))
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(grd1.TextMatrix(intCompRow, 3))
            dblSum = dblSum + Val(grd1.TextMatrix(intCompRow, 3))
            iPrint = iPrint + 300
            intCompRow = intCompRow + 1
         Else
            Exit For
         End If
      Next ii
      PrintEnd2
   End If
   
   '處理傳進來下一人員(之前人員)的資料
   m_strTemp1 = "": m_strTemp3 = "": bNextTrue = False: dblSum = 0
   For ii = intCompRow To grd1.Rows - 1
      
      If (Val(grd1.TextMatrix(ii, 1)) >= Val(strNext1) And Val(strNext1) <> 0) Then GoTo GoToExit
      
      If (grd1.TextMatrix(ii, 1) <> m_strTemp1) Then
         bNextTrue = True
         PrintEnd2
         dblSum = 0
         'one new data start
         m_strTemp1 = Trim(grd1.TextMatrix(ii, 1))
         m_strTemp3 = Trim(grd1.TextMatrix(ii, 12))
         If Page <> 1 Then
            Call ShowLine1(0)
            iPrint = iPrint + 600
         End If
         If iPrint >= 10000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle
            PrintTitle2
         Else
            PrintTitle2
         End If
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         Printer.Print "請款單分配點數："
         iPrint = iPrint + 300
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         Printer.Print "請款日"
         Printer.CurrentX = 1500
         Printer.CurrentY = iPrint
         Printer.Print "本所案號"
         Printer.CurrentX = 3500
         Printer.CurrentY = iPrint
         Printer.Print "案件性質"
         Printer.CurrentX = 5000 - Printer.TextWidth("點數")
         Printer.CurrentY = iPrint
         Printer.Print "點數"
         iPrint = iPrint + 300
         Call ShowLine1(1)
      End If
      If iPrint >= 10000 Then
         Page = Page + 1
         Printer.NewPage
         PrintTitle
         PrintTitle2
      End If
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print ChangeTStringToTDateString(grd1.TextMatrix(ii, 9))
      Printer.CurrentX = 1500
      Printer.CurrentY = iPrint
      Printer.Print grd1.TextMatrix(ii, 4) & "-" & grd1.TextMatrix(ii, 5) & "-" & grd1.TextMatrix(ii, 6) & "-" & grd1.TextMatrix(ii, 7)
      Printer.CurrentX = 3500
      Printer.CurrentY = iPrint
      Printer.Print grd1.TextMatrix(ii, 8)
      Printer.CurrentX = 5000 - Printer.TextWidth(CheckStr(grd1.TextMatrix(ii, 3)))
      Printer.CurrentY = iPrint
      Printer.Print CheckStr(grd1.TextMatrix(ii, 3))
      dblSum = dblSum + Val(grd1.TextMatrix(ii, 3))
      iPrint = iPrint + 300
      intCompRow = intCompRow + 1
   Next ii
GoToExit:
   If bNextTrue = True Then
      PrintEnd2
   End If
End Sub

'Add By Sindy 2010/5/5
Sub ShowLine1(intType As Integer)
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   If intType = 0 Then '長線
      Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
   ElseIf intType = 1 Then '短線
      Printer.Line (0, iPrint + 150)-(5500, iPrint + 150)
   End If
   iPrint = iPrint + 300
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
      PrintTitle2
   End If
End Sub

'Add By Sindy 2010/5/5
Sub PrintEnd2()
   If dblSum > 0 Then
      Call ShowLine1(1)
      If iPrint >= 10000 Then
         Page = Page + 1
         Printer.NewPage
         PrintTitle
         PrintTitle2
      End If
      Printer.CurrentX = 3500
      Printer.CurrentY = iPrint
      Printer.Print "合計"
      Printer.CurrentX = 5000 - Printer.TextWidth(CheckStr(dblSum))
      Printer.CurrentY = iPrint
      Printer.Print CheckStr(dblSum)
      iPrint = iPrint + 300
   End If
End Sub

'Add By Sindy 2010/5/5
Sub PrintTitle2()
GetPleft
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "業務區：" & m_strTemp1
Printer.CurrentX = 5000
Printer.CurrentY = iPrint
Printer.Print "智權人員：" & m_strTemp3
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

'Add By Sindy 2012/2/3 取得此人員的分配點數小計
Sub GetPointTotSub(strEmp As String)
   dblPointTotSub = 0
   If dblPointRow > 0 Then
      For i = 1 To grd1.Rows - 1
         If strEmp = grd1.TextMatrix(i, 1) Then
            dblPointTotSub = grd1.TextMatrix(i, 13)
            Exit Sub
         End If
      Next i
   End If
End Sub
