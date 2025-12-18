VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030408 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人案件明細表"
   ClientHeight    =   2625
   ClientLeft      =   5055
   ClientTop       =   2985
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3285
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2240
      Width           =   390
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1920
      Width           =   390
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1152
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1584
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1152
      TabIndex        =   0
      Top             =   504
      Width           =   1740
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1152
      MaxLength       =   1
      TabIndex        =   1
      Top             =   900
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1152
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1248
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2136
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1248
      Width           =   930
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1572
      TabIndex        =   7
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2364
      TabIndex        =   8
      Top             =   36
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2505
      Left            =   180
      TabIndex        =   17
      Top             =   2640
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4419
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "(Y/N)"
      Height          =   180
      Left            =   2352
      TabIndex        =   19
      Top             =   2282
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "是否含取消收文資料："
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   18
      Top             =   2282
      Width           =   1830
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   1944
      TabIndex        =   16
      Top             =   1620
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "(Y/N)"
      Height          =   180
      Left            =   2352
      TabIndex        =   15
      Top             =   1944
      Width           =   516
   End
   Begin VB.Label Label1 
      Caption         =   "是否含已發文資料："
      Height          =   180
      Index           =   5
      Left            =   156
      TabIndex        =   14
      Top             =   1968
      Width           =   1716
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   3
      Left            =   156
      TabIndex        =   13
      Top             =   1632
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   156
      TabIndex        =   12
      Top             =   528
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   1
      Left            =   156
      TabIndex        =   11
      Top             =   948
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   156
      TabIndex        =   10
      Top             =   1284
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收文  2.發文)"
      Height          =   180
      Index           =   4
      Left            =   1524
      TabIndex        =   9
      Top             =   924
      Width           =   1416
   End
   Begin VB.Line Line1 
      X1              =   1572
      X2              =   2832
      Y1              =   1392
      Y2              =   1392
   End
End
Attribute VB_Name = "frm030408"
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
     If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
     Else
         If Len(Txt1(1)) = 0 Then
             s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
             Txt1(1).SetFocus
             Exit Sub
         Else
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
                 s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
                 Txt1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             Else
                 If Len(Txt1(5)) = 0 Then
                     s = MsgBox("是否已發文資料不可空白!!", , "USER 輸入錯誤")
                     Txt1(5).SetFocus
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
cnnConnection.Execute "DELETE FROM R030408 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
strPointSql = "" 'Add By Sindy 2010/5/4
If Len(Txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(Txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(Txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/10/22
End If
Select Case Val(Txt1(1))
Case 1
     If Len(Txt1(2)) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(Txt1(2))) & " "
      strPointSql = strPointSql & " AND CP05>=" & Val(ChangeTStringToWString(Txt1(2))) & " " 'Add By Sindy 2010/5/4
     End If
     If Len(Trim(Txt1(3))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(Txt1(3))) & " "
      strPointSql = strPointSql & " AND CP05<=" & Val(ChangeTStringToWString(Txt1(3))) & " " 'Add By Sindy 2010/5/4
     End If
     If Len(Txt1(2)) <> 0 Or Len(Trim(Txt1(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";收文" & Label1(2) & Txt1(2) & "-" & Txt1(3) 'Add By Sindy 2010/10/22
     End If
Case 2
     If Len(Txt1(2)) <> 0 Then
     StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(Txt1(2))) & " "
     strPointSql = strPointSql & " AND CP27>=" & Val(ChangeTStringToWString(Txt1(2))) & " " 'Add By Sindy 2010/5/4
     End If
     If Len(Trim(Txt1(3))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(Txt1(3))) & " "
      strPointSql = strPointSql & " AND CP27<=" & Val(ChangeTStringToWString(Txt1(3))) & " " 'Add By Sindy 2010/5/4
     End If
     If Len(Txt1(2)) <> 0 Or Len(Trim(Txt1(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";發文" & Label1(2) & Txt1(2) & "-" & Txt1(3) 'Add By Sindy 2010/10/22
     End If
Case Else
End Select
Select Case Trim(Txt1(5))
Case "N"
     StrSQL6 = StrSQL6 + " AND (CP27 IS NULL OR CP27='') "
     strPointSql = strPointSql & " AND (CP27 IS NULL OR CP27='') " 'Add By Sindy 2010/5/4
     pub_QL05 = pub_QL05 & ";" & Label1(5) & Txt1(5)  'Add By Sindy 2010/10/22
Case Else
     pub_QL05 = pub_QL05 & ";" & Label1(5) & "Y"  'Add By Sindy 2010/10/22
End Select
If Len(Txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP14='" & Txt1(4) & "' "
    strPointSql = strPointSql & " AND a1n04='" & Txt1(4) & "' " 'Add By Sindy 2010/5/4
    pub_QL05 = pub_QL05 & ";" & Label1(3) & Txt1(4) & LBL1 'Add By Sindy 2010/10/22
End If
'Added by Lydia 2016/05/13 是否含取消收文
If Trim(Txt1(6)) = "" Or Trim(Txt1(6)) <> "Y" Then
     StrSQL6 = StrSQL6 + " AND (CP57 IS NULL OR CP57='') "
     strPointSql = strPointSql & " AND (CP57 IS NULL OR CP57='') "
End If
pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(6)
'end 2016/05/13

'Add By Sindy 2010/5/4
Call QueryPointData(strPointSql)
'2010/5/4 End

CheckOC
Select Case Val(Txt1(1))
Case 1
      'Modify By Sindy 2010/5/4 點數改抓點數分配檔
'     strSql = "SELECT S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(tm10,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP27") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND TM10=N1.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CU10=N2.NA01(+)  " & strSQL1 & StrSQL6
'     strSql = strSql + " union all select S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(sp09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP27") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND SP09=N1.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CU10=N2.NA01(+) " & strSQL2 & StrSQL6
      strSql = "SELECT S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(tm10,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP27") & ",decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,acc1n0,acc1k0 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND TM10=N1.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CU10=N2.NA01(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 AND a1k01(+)=cp60 " & strSQL1 & StrSQL6
      strSql = strSql + " union all select S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(sp09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP27") & ",decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,acc1n0,acc1k0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND SP09=N1.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CU10=N2.NA01(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 AND a1k01(+)=cp60 " & strSQL2 & StrSQL6
Case 2
      'Modify By Sindy 2010/5/4 點數改抓點數分配檔
'     strSql = "SELECT S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(tm10,'000',CPM03,CPM04),nvl(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND TM10=N1.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CU10=N2.NA01(+) " & strSQL1 & StrSQL6
'     strSql = strSql + " union all select S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(sp09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND SP09=N1.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CU10=N2.NA01(+) " & strSQL2 & StrSQL6
      strSql = "SELECT S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(tm10,'000',CPM03,CPM04),nvl(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP05") & ",decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,acc1n0,acc1k0 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND TM10=N1.NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CU10=N2.NA01(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 AND a1k01(+)=cp60 " & strSQL1 & StrSQL6
      strSql = strSql + " union all select S1.ST01," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),decode(sp09,'000',CPM03,CPM04),NVL(N1.NA03,N1.NA04),NVL(N2.NA03,N2.NA04),S2.ST02," & SQLDate("CP05") & ",decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,acc1n0,acc1k0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=S2.ST01(+) AND CP14=S1.ST01(+) AND SP09=N1.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CU10=N2.NA01(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 AND a1k01(+)=cp60 " & strSQL2 & StrSQL6
Case Else
End Select
cnnConnection.Execute "INSERT INTO R030408 " & strSql
CheckOC
strSql = "SELECT * FROM R030408 WHERE ID='" & strUserNum & "' "
With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      'InsertQueryLog (.RecordCount + dblPointTot) 'Add By Sindy 2010/10/22
      InsertQueryLog (.RecordCount + dblPointRow) 'Add By Sindy 2012/2/3
   Else
      'Modify By Sindy 2010/5/6
      'If dblPointTot > 0 Then
      'Modify By Sindy 2012/2/3
      If dblPointRow > 0 Then
         'InsertQueryLog (dblPointTot) 'Add By Sindy 2010/10/22
         InsertQueryLog (dblPointRow) 'Add By Sindy 2012/2/3
         Page = 1
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
End With
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "外商承辦人案件明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Select Case Val(Txt1(1))
Case 1
     Printer.Print "收文日：" & Format(ChangeTStringToTDateString(Txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(3))
Case 2
     Printer.Print "發文日：" & Format(ChangeTStringToTDateString(Txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(3))
Case Else
End Select
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & SavDay1
'Added by Lydia 2016/05/13 是否含取消收文
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print Label1(6).Caption & Txt1(6)
'end 2016/05/13
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Select Case Val(Txt1(1))
Case 1
     Printer.Print "收文日"
Case 2
     Printer.Print "發文日"
Case Else
End Select
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "申請人國籍"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Select Case Val(Txt1(1))
Case 1
     Printer.Print "發文日"
Case 2
     Printer.Print "收文日"
Case Else
End Select
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "點數"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 1 To 10
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
PLeft(2) = 1000
PLeft(3) = 2000
PLeft(4) = 3600
PLeft(5) = 7000
PLeft(6) = 9000
PLeft(7) = 11000 - 500
PLeft(8) = 12000 - 500
PLeft(9) = 13000 + 200 - 500
PLeft(10) = 14000 + 200 - 500
PLeft(11) = 15500 - 500
End Sub

Sub PrintEnd()
Call GetPointTotSub(SavDay3) 'Add By Sindy 2012/2/3 取得此人員的分配點數小計
If Len(SavDay3) = 0 Then
    strSql = "SELECT COUNT(*),SUM(R097012) FROM R030408 WHERE ID='" & strUserNum & "' AND (R097001='' or r097001 is null) "
Else
    strSql = "SELECT COUNT(*),SUM(R097012) FROM R030408 WHERE ID='" & strUserNum & "' AND R097001='" & SavDay3 & "' "
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
End Sub

Sub PrintEnd1()
strSql = "SELECT COUNT(*),SUM(R097012) FROM R030408 WHERE ID='" & strUserNum & "' "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   'Add By Sindy 2010/5/6
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
   End If
   '2010/5/6 End
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "總計"
   'Modify By Sindy 2012/2/3
'   Printer.CurrentX = PLeft(8)
'   Printer.CurrentY = iPrint
'   Printer.Print "件數：" & Format(CheckStr(adoRecordset1.Fields(0)), "###,###,##0")
'   Printer.CurrentX = PLeft(10)
'   Printer.CurrentY = iPrint
'   Printer.Print "點數：" & Format(CheckStr(adoRecordset1.Fields(1)), "###,###,##0.00")
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
'Add By Sindy 2010/5/5
'If dblPointTot > 0 Then
'   ShowLine
'   If iPrint >= 10000 Then
'      Page = Page + 1
'      Printer.NewPage
'      PrintTitle
'   End If
'   Printer.CurrentX = PLeft(8)
'   Printer.CurrentY = iPrint
'   Printer.Print "分配點數"
'   Printer.CurrentX = PLeft(10)
'   Printer.CurrentY = iPrint
'   Printer.Print dblPointTot
'   iPrint = iPrint + 300
'End If
iPrint = iPrint + 600
If iPrint >= 10000 Then
   Page = Page + 1
   Printer.NewPage
   PrintTitle
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "PS.不含非個人點數"
iPrint = iPrint + 300
'2010/5/5 End
End Sub

Sub ShowLine()
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
   iPrint = iPrint + 300
   'Modify By Sindy 2010/5/6
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
   End If
   '2010/5/6 End
End Sub

Sub PrintData()
strSql = "SELECT ST02,R097002,R097003,R097004,R097005,R097006,R097007,R097008,R097009,R097010,R097011,R097012,R097001 FROM R030408,STAFF WHERE R097001=ST01(+) AND ID='" & strUserNum & "' order by r097001,r097003,r097004 "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = "             " 'CheckStr(.Fields(0))
        SavDay2 = "             " ' CheckStr(.Fields(2))
        SavDay3 = "             " ' CheckStr(.Fields(12))
        'PrintTitle
        Call PrintEndPoint("", CheckStr(.Fields(12))) 'Add By Sindy 2010/5/4
        Do While .EOF = False
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(0) <> SavDay1 Then
                If Len(Trim(SavDay1)) <> 0 Then
                  ShowLine
                  PrintEnd
                  Call PrintEndPoint(SavDay3, CheckStr(.Fields(12))) 'Add By Sindy 2010/5/4
                  Page = Page + 1
                  Printer.NewPage
                End If
                SavDay1 = strTemp(0)
                SavDay2 = strTemp(2)
                SavDay3 = CheckStr(.Fields(12))
                PrintTitle
            Else
                If strTemp(2) = SavDay2 Then
                  strTemp(2) = ""
                Else
                  SavDay2 = strTemp(2)
                End If
            End If
            strTemp(4) = StrToStr(strTemp(4), 15)
            strTemp(5) = StrToStr(strTemp(5), 9)
            strTemp(6) = StrToStr(strTemp(6), 5)
            strTemp(7) = StrToStr(strTemp(7), 4)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(9) = StrToStr(strTemp(9), 4)
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
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

Private Sub Form_Load()
MoveFormToCenter Me
Txt1(0) = GetSystemKindByNick
Txt1(6) = "N" 'Added by Lydia 2016/05/13
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm030408 = Nothing
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
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(Txt1(0)), ",,", ""), ",")
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
Case 1
     Select Case Trim(Txt1(1))
     Case "1"
          Txt1(5) = "N"
     Case "2"
          Txt1(5) = "Y"
     Case ""
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(1).SetFocus
          Txt1(1).SelStart = 0
          Txt1(1).SelLength = Len(Txt1(1))
          Exit Sub
     End Select
Case 3, 2
   If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
      Me.Txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 3 Then
     If RunNick(Txt1(Index - 1), Txt1(Index)) Then
         Txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case 4
     LBL1 = GetPrjSalesNM(Txt1(4))
     If Trim(Txt1(4)) <> "" Then
        If Trim(LBL1.Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            Txt1(4).SetFocus
            txt1_GotFocus (4)
            Exit Sub
        End If
     End If
Case 5
     Select Case Trim(Txt1(5))
     Case "Y", "N"
     Case Else
          s = MsgBox("是否含已發文資料只能輸入 Y 或 N !!", , "USER 輸入錯誤")
          Txt1(5).SetFocus
          Txt1(5).SelStart = 0
          Txt1(5).SelLength = Len(Txt1(5))
          Exit Sub
     End Select
'Added by Lydia 2016/05/13
Case 6
     If Txt1(Index) <> "" And Txt1(Index) <> "Y" And Txt1(Index) <> "N" Then
        s = MsgBox("是否含取消收文資料只能輸入 Y 或 N !!", , "USER 輸入錯誤")
        Txt1(Index).SetFocus
        Txt1(Index).SelStart = 0
        Txt1(Index).SelLength = Len(Txt1(Index))
        Exit Sub
     End If
Case Else
End Select
End Sub

'Add By Sindy 2010/4/30
'非自己辦理的案件但點數歸屬自己
Private Sub QueryPointData(strConSql As String)
Dim strEmp As String 'Add By Sindy 2012/2/3
   
   With GRD1
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      .FormatString = "cp12|a1n04|a1n03|a1n05|cp01|cp02|cp03|cp04|cp10|a1k02"
   End With
   intCompRow = 0
   dblPointRow = 0 'Add By Sindy 2012/2/3
   dblPointTot = 0
   strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,DECODE(TM10,'000',CPM03,CPM04),a1k02,cp14," & IIf(Txt1(1) = "1", "cp05", "cp27") & ",a.st02,' ' 點數小計 " & _
                  "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                  "WHERE cp60>'X' AND cp14 is not null " & _
                  "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 and a1n05>0 " & _
                  "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                  "AND a1k25 is null " & _
                  "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp14=b.st01(+) " & _
                  "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
                  "AND CP01 IN (" & SQLGrpStr(Txt1(0), 2) & ") " & _
                  "AND cp01=tm01 AND cp02=tm02 AND cp03=tm03 AND cp04=tm04 " & strConSql
   strSql = strSql & " UNION ALL " & _
                  "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,DECODE(SP09,'000',CPM03,CPM04),a1k02,cp14," & IIf(Txt1(1) = "1", "cp05", "cp27") & ",a.st02,' ' 點數小計 " & _
                  "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                  "WHERE cp60>'X' AND cp14 is not null " & _
                  "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 and a1n05>0 " & _
                  "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                  "AND a1k25 is null " & _
                  "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp14=b.st01(+) " & _
                  "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
                  "AND CP01 IN (" & SQLGrpStr(Txt1(0), 5) & ") " & _
                  "AND cp01=sp01 AND cp02=sp02 AND cp03=sp03 AND cp04=sp04 " & strConSql & _
                  "Order By a1n04,a1k02,cp01,cp02,cp03,cp04 "
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      intCompRow = 1 '代表有分配點數資料,從第1筆開始讀取
      dblPointRow = adoRecordset.RecordCount 'Add By Sindy 2012/2/3
      Set GRD1.Recordset = adoRecordset.Clone
      
      'Modify By Sindy 2012/2/3
'      strSql = "select sum(tt) from (" & _
'                     "SELECT sum(a1n05) tt " & _
'                     "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                     "WHERE cp60>'X' AND cp14 is not null " & _
'                     "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 and a1n05>0 " & _
'                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                     "AND a1k25 is null " & _
'                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp14=b.st01(+) " & _
'                     "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
'                     "AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") " & _
'                     "AND cp01=tm01 AND cp02=tm02 AND cp03=tm03 AND cp04=tm04 " & strConSql
'      strSql = strSql & " UNION ALL " & _
'                     "SELECT sum(a1n05) tt " & _
'                     "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                     "WHERE cp60>'X' AND cp14 is not null " & _
'                     "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 and a1n05>0 " & _
'                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                     "AND a1k25 is null " & _
'                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp14=b.st01(+) " & _
'                     "AND substr(a.st15,1,2)=substr(b.st15,1,2) " & _
'                     "AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") " & _
'                     "AND cp01=sp01 AND cp02=sp02 AND cp03=sp03 AND cp04=sp04 " & strConSql & _
'                   ")"
'      intI = 1
'      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         dblPointTot = adoRecordset.Fields(0)
'      End If
      '計算分配點數的小計及總計
      For i = 1 To GRD1.Rows - 1
         If strEmp <> GRD1.TextMatrix(i, 1) Then
            If strEmp <> "" Then
               For j = 1 To GRD1.Rows - 1
                  If GRD1.TextMatrix(j, 1) = strEmp Then
                     GRD1.TextMatrix(j, 13) = dblPointTotSub
                  End If
               Next j
            End If
            strEmp = GRD1.TextMatrix(i, 1)
            dblPointTotSub = 0
         End If
         dblPointTotSub = dblPointTotSub + Val(GRD1.TextMatrix(i, 3))
         dblPointTot = dblPointTot + Val(GRD1.TextMatrix(i, 3))
      Next i
      If strEmp <> "" Then
         For j = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(j, 1) = strEmp Then
               GRD1.TextMatrix(j, 13) = dblPointTotSub
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
   
   If intCompRow = 0 Or intCompRow > (GRD1.Rows - 1) Then Exit Sub
   
   '處理傳進來比對人員的資料
   If GRD1.TextMatrix(intCompRow, 1) = strComp1 Then
      m_strTemp1 = strComp1
      m_strTemp3 = Trim(GRD1.TextMatrix(intCompRow, 12))
      iPrint = iPrint + 300
      If iPrint >= 10000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle1
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
      ShowLine1
      dblSum = 0
      For ii = intCompRow To GRD1.Rows - 1
         If GRD1.TextMatrix(intCompRow, 1) = strComp1 Then
            m_strTemp1 = strComp1
            m_strTemp3 = Trim(GRD1.TextMatrix(intCompRow, 12))
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print ChangeTStringToTDateString(GRD1.TextMatrix(intCompRow, 9))
            Printer.CurrentX = 1500
            Printer.CurrentY = iPrint
            Printer.Print GRD1.TextMatrix(intCompRow, 4) & "-" & GRD1.TextMatrix(intCompRow, 5) & "-" & GRD1.TextMatrix(intCompRow, 6) & "-" & GRD1.TextMatrix(intCompRow, 7)
            Printer.CurrentX = 3500
            Printer.CurrentY = iPrint
            Printer.Print GRD1.TextMatrix(intCompRow, 8)
            Printer.CurrentX = 5000 - Printer.TextWidth(CheckStr(GRD1.TextMatrix(intCompRow, 3)))
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(GRD1.TextMatrix(intCompRow, 3))
            dblSum = dblSum + Val(GRD1.TextMatrix(intCompRow, 3))
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
   For ii = intCompRow To GRD1.Rows - 1
      
      If (Val(GRD1.TextMatrix(ii, 1)) >= Val(strNext1) And Val(strNext1) <> 0) Then GoTo GoToExit
      
      If (GRD1.TextMatrix(ii, 1) <> m_strTemp1) Then
         bNextTrue = True
         PrintEnd2
         dblSum = 0
         'one new data start
         m_strTemp1 = Trim(GRD1.TextMatrix(ii, 1))
         m_strTemp3 = Trim(GRD1.TextMatrix(ii, 12))
         If Page <> 1 Then
            Page = Page + 1
            Printer.NewPage
         End If
         PrintTitle1
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
         ShowLine1
      End If
      If iPrint >= 10000 Then
         Page = Page + 1
         Printer.NewPage
         PrintTitle1
      End If
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print ChangeTStringToTDateString(GRD1.TextMatrix(ii, 9))
      Printer.CurrentX = 1500
      Printer.CurrentY = iPrint
      Printer.Print GRD1.TextMatrix(ii, 4) & "-" & GRD1.TextMatrix(ii, 5) & "-" & GRD1.TextMatrix(ii, 6) & "-" & GRD1.TextMatrix(ii, 7)
      Printer.CurrentX = 3500
      Printer.CurrentY = iPrint
      Printer.Print GRD1.TextMatrix(ii, 8)
      Printer.CurrentX = 5000 - Printer.TextWidth(CheckStr(GRD1.TextMatrix(ii, 3)))
      Printer.CurrentY = iPrint
      Printer.Print CheckStr(GRD1.TextMatrix(ii, 3))
      dblSum = dblSum + Val(GRD1.TextMatrix(ii, 3))
      iPrint = iPrint + 300
      intCompRow = intCompRow + 1
   Next ii
GoToExit:
   If bNextTrue = True Then
      PrintEnd2
   End If
End Sub

'Add By Sindy 2010/5/5
Sub ShowLine1()
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Line (0, iPrint + 150)-(5500, iPrint + 150)
   iPrint = iPrint + 300
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle1
   End If
End Sub

'Add By Sindy 2010/5/5
Sub PrintEnd2()
   If dblSum > 0 Then
      ShowLine1
      If iPrint >= 10000 Then
         Page = Page + 1
         Printer.NewPage
         PrintTitle1
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
Sub PrintTitle1()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "外商承辦人案件明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Select Case Val(Txt1(1))
Case 1
     Printer.Print "收文日：" & Format(ChangeTStringToTDateString(Txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(3))
Case 2
     Printer.Print "發文日：" & Format(ChangeTStringToTDateString(Txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(3))
Case Else
End Select
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & m_strTemp3
'Added by Lydia 2016/05/13 是否含取消收文
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print Label1(6).Caption & Txt1(6)
'end 2016/05/13
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

'Add By Sindy 2012/2/3 取得此人員的分配點數小計
Sub GetPointTotSub(strEmp As String)
   dblPointTotSub = 0
   If dblPointRow > 0 Then
      For i = 1 To GRD1.Rows - 1
         If strEmp = GRD1.TextMatrix(i, 1) Then
            dblPointTotSub = GRD1.TextMatrix(i, 13)
            Exit Sub
         End If
      Next i
   End If
End Sub
