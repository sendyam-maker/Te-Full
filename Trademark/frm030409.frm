VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030409 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文業績明細表"
   ClientHeight    =   4065
   ClientLeft      =   4080
   ClientTop       =   3900
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3945
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   1
      Top             =   870
      Width           =   240
   End
   Begin VB.OptionButton opt1 
      Caption         =   "依承辦人統計"
      Height          =   225
      Index           =   1
      Left            =   1950
      TabIndex        =   11
      Top             =   2910
      Width           =   1425
   End
   Begin VB.OptionButton opt1 
      Caption         =   "依智權人員統計"
      Height          =   225
      Index           =   0
      Left            =   156
      TabIndex        =   10
      Top             =   2910
      Value           =   -1  'True
      Width           =   1665
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1860
      Width           =   760
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1860
      Width           =   760
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "2"
      Top             =   3210
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1050
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2565
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1050
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2205
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   576
      Width           =   2325
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3540
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1200
      Width           =   760
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1935
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1200
      Width           =   760
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1530
      Width           =   760
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1935
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1530
      Width           =   760
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2316
      TabIndex        =   14
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3108
      TabIndex        =   15
      Top             =   48
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2505
      Left            =   150
      TabIndex        =   30
      Top             =   4170
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4419
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "日期別：            (1.請款 2.發文)"
      Height          =   180
      Index           =   11
      Left            =   156
      TabIndex        =   29
      Top             =   930
      Width           =   2550
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   2310
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   10
      Left            =   156
      TabIndex        =   28
      Top             =   1890
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "(1.有點數  2.全部)"
      Height          =   180
      Index           =   9
      Left            =   1425
      TabIndex        =   27
      Top             =   3240
      Width           =   2430
   End
   Begin VB.Label Label1 
      Caption         =   "列印內容："
      Height          =   180
      Index           =   8
      Left            =   156
      TabIndex        =   26
      Top             =   3255
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "(3.查詢明細  4.查詢總計)"
      Height          =   180
      Index           =   7
      Left            =   1425
      TabIndex        =   25
      Top             =   3780
      Width           =   2430
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   6
      Left            =   156
      TabIndex        =   24
      Top             =   2625
      Width           =   990
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   2205
      TabIndex        =   23
      Top             =   2610
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   5
      Left            =   156
      TabIndex        =   22
      Top             =   2265
      Width           =   990
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   2220
      TabIndex        =   21
      Top             =   2250
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   156
      TabIndex        =   20
      Top             =   600
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   1
      Left            =   156
      TabIndex        =   19
      Top             =   3585
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   156
      TabIndex        =   18
      Top             =   1245
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   3
      Left            =   156
      TabIndex        =   17
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "(1.列印明細  2.列印總計)"
      Height          =   180
      Index           =   4
      Left            =   1425
      TabIndex        =   16
      Top             =   3570
      Width           =   2430
   End
   Begin VB.Line Line1 
      X1              =   1230
      X2              =   2490
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line2 
      X1              =   1575
      X2              =   2325
      Y1              =   1620
      Y2              =   1620
   End
End
Attribute VB_Name = "frm030409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay(0 To 2) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String
Dim StrTemp4(0 To 4) As String, StrTemp5(0 To 4) As String
'Add By Sindy 2010/5/4
Dim intCompRow As Integer
Dim dblPointTot As Double
Dim m_strTemp1 As String, m_strTemp3 As String
Dim dblSum As Double
'2010/5/4 End
Dim dblPointTotSub As Double, dblPointRow As Integer 'Add By Sindy 2012/2/3
'Added by Lydia 2017/01/04 增加法務案件
Dim StrSQL3 As String
Dim strSysKind As String 'Add by Amy 2021/07/23 可查詢系統別

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
Printer.Orientation = 2
DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Sindy 98/03/12
         If Len(txt1(11)) = 0 Then
            s = MsgBox("日期別不可空白!!", , "USER 輸入錯誤")
            txt1(11).SetFocus
            Exit Sub
         End If
         '98/03/12 End
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
             s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
             If Len(txt1(7)) = 0 Then
                 s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                 txt1(7).SetFocus
                 Exit Sub
             End If
             If Len(txt1(8)) = 0 Then
                 s = MsgBox("列印內容不可空白!!", , "USER 輸入錯誤")
                 txt1(8).SetFocus
                 Exit Sub
             End If
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
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R030409 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = "" 'Added by Lydia 2017/01/04
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   'Added by Lydia 2017/01/04 增加法務案
   If SQLGrpStr(txt1(0), 3) <> "' '" Then
      StrSQL3 = StrSQL3 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") "
   End If
   'END 2017/01/04
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/22
End If
If Len(Trim(txt1(1))) <> 0 Then
   'Modify By Sindy 98/03/12 增加依請款日查詢
   If Trim(txt1(11)) = "1" Then
      '請款日(起)
      StrSQL6 = StrSQL6 + " AND CP60=A1K01 AND A1K02 >=" & Val(ChangeTStringToWString(txt1(1))) - 19110000 & " "
   Else
      '發文日(起)
      StrSQL6 = StrSQL6 + " AND CP60=A1K01(+) AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   End If
End If
If Len(Trim(txt1(2))) <> 0 Then
   'Modify By Sindy 98/03/12 增加依請款日查詢
   If Trim(txt1(11)) = "1" Then
      '請款日(迄)
      StrSQL6 = StrSQL6 + " AND A1K02 <=" & Val(ChangeTStringToWString(txt1(2))) - 19110000 & " "
   Else
      '發文日(迄)
      StrSQL6 = StrSQL6 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
End If
If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   If Trim(txt1(11)) = "1" Then
      pub_QL05 = pub_QL05 & ";請款" & Label1(2) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/22
   Else
      pub_QL05 = pub_QL05 & ";發文" & Label1(2) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/22
   End If
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4)  'Add By Sindy 2010/10/22
End If
'智權人員
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP13='" & txt1(5) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(5) & lbl1(0) 'Add By Sindy 2010/10/22
End If
'承辦人
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP14='" & txt1(6) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(6) & lbl1(1) 'Add By Sindy 2010/10/22
End If
'If Len(txt1(8)) <> 0 Then
'    If Me.txt1(8).Text = "1" Then
'        StrSQL6 = StrSQL6 + " AND CP18 > 0 "
'    End If
'End If
'申請國家(起)
If Len(txt1(9)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10 >= '" & Me.txt1(9).Text & "' "
    strSQL2 = strSQL2 + " AND SP09 >= '" & Me.txt1(9).Text & "' "
    If StrSQL3 <> "" Then StrSQL3 = StrSQL3 + " AND LC15 >= '" & Me.txt1(9).Text & "' " 'Added by Lydia 2017/01/04
End If
'申請國家(迄)
If Len(txt1(10)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10 <= '" & Me.txt1(10).Text & "z' "
    strSQL2 = strSQL2 + " AND SP09 <= '" & Me.txt1(10).Text & "z' "
    If StrSQL3 <> "" Then StrSQL3 = StrSQL3 + " AND LC15 >= '" & Me.txt1(10).Text & "z' " 'Added by Lydia 2017/01/04
End If
If Len(txt1(9)) <> 0 Or Len(txt1(10)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(10) & txt1(9) & "-" & txt1(10)  'Add By Sindy 2010/10/22
End If

'Add By Sindy 2010/5/4
Call QueryPointData
'2010/5/4 End

CheckOC
'edit by nickc 2007/12/17 改以選擇的功能
'Modify By Sindy 2010/4/27 點數改抓點數分配檔
If opt1(1).Value = True Then
    'Modify By Sindy 98/03/12 加工作時數, 及增加依請款日查詢
'    strSQL = "SELECT S2.ST01, A0901,S1.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(TM10,'000',CPM03,CPM04),CP18,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " union all select S2.ST01, A0901,S1.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(SP09,'000',CPM03,CPM04),CP18,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
    'R098001 : 承辦人ID
    'R098002 : 業務區
    'R098003 : 智權人員名稱
    If Trim(txt1(11)) = "1" Then '1.請款
      strSql = "SELECT S2.ST01, A0901,S1.ST02," & SqlDateT("A1K02") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(TM10,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S2.ST15 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 " & strSQL1 & StrSQL6
      strSql = strSql + " union all select S2.ST01, A0901,S1.ST02," & SqlDateT("A1K02") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(SP09,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S2.ST15 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 " & strSQL2 & StrSQL6
      'Added by Lydia 2017/01/04 增加法務案
      If StrSQL3 <> "" Then
         strSql = strSql + " union all select S2.ST01, A0901,S1.ST02," & SqlDateT("A1K02") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(LC15,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S2.ST15 FROM CASEPROGRESS,LAWCASE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 " & StrSQL3 & StrSQL6
      End If
      'end 2017/01/04
    Else
      strSql = "SELECT S2.ST01, A0901,S1.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(TM10,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S2.ST15 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 " & strSQL1 & StrSQL6
      strSql = strSql + " union all select S2.ST01, A0901,S1.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(SP09,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S2.ST15 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 " & strSQL2 & StrSQL6
      'Added by Lydia 2017/01/04 增加法務案
      If StrSQL3 <> "" Then
         strSql = strSql + " union all select S2.ST01, A0901,S1.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(LC15,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S2.ST15 FROM CASEPROGRESS,LAWCASE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)=cp14 " & StrSQL3 & StrSQL6
      End If
      'end 2017/01/04
    End If
Else
    'Modify By Sindy 98/03/12 加工作時數, 及增加依請款日查詢
'    strSQL = "SELECT A0901,S1.ST01,S2.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(TM10,'000',CPM03,CPM04),CP18,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
'    strSQL = strSQL + " union all select A0901,S1.ST01,S2.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(SP09,'000',CPM03,CPM04),CP18,'" & strUserNum & "'  FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
    'R098001 : 業務區
    'R098002 : 智權人員ID
    'R098003 : 承辦人名稱
    If Trim(txt1(11)) = "1" Then '1.請款
      strSql = "SELECT A0901,S1.ST01,S2.ST02," & SqlDateT("A1K02") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(TM10,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S1.ST15 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 " & strSQL1 & StrSQL6
      strSql = strSql + " union all select A0901,S1.ST01,S2.ST02," & SqlDateT("A1K02") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(SP09,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S1.ST15 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 " & strSQL2 & StrSQL6
      'Added by Lydia 2017/01/04 增加法務案
      If StrSQL3 <> "" Then
         strSql = strSql + " union all select A0901,S1.ST01,S2.ST02," & SqlDateT("A1K02") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(LC15,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S1.ST15 FROM CASEPROGRESS,LAWCASE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 " & StrSQL3 & StrSQL6
      End If
      'end 2017/01/04
    Else
      strSql = "SELECT A0901,S1.ST01,S2.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(TM10,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S1.ST15 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 " & strSQL1 & StrSQL6
      strSql = strSql + " union all select A0901,S1.ST01,S2.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(SP09,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S1.ST15 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 " & strSQL2 & StrSQL6
      'Added by Lydia 2017/01/04 增加法務案
      If StrSQL3 <> "" Then
         strSql = strSql + " union all select A0901,S1.ST01,S2.ST02," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(LC15,'000',CPM03,CPM04),decode(substr(cp60,1,1),'X',decode(a1k25,null,a1n05,''),cp18),'" & strUserNum & "',CP113,S1.ST15 FROM CASEPROGRESS,LAWCASE,STAFF S1,STAFF S2,CASEPROPERTYMAP,ACC090,ACC1K0,acc1n0 WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP12=A0901(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)=cp13 " & strSQL2 & StrSQL6
      End If
      'end 2017/01/04
    End If
End If
cnnConnection.Execute "INSERT INTO R030409 " & strSql

'Add By Sindy 2010/10/14
'查詢時,將分配點數資料也一併insert到R030409
Dim strDate As String
If txt1(7) = "3" Or txt1(7) = "4" Then
   If grd1.Rows > 1 And grd1.TextMatrix(1, 1) <> "" Then '有分配點數資料則需要insert至R030409
      For i = 1 To grd1.Rows - 1
         If txt1(11) = "1" Then '請款日期
            strDate = ChangeTStringToTDateString(grd1.TextMatrix(i, 9))
         Else '發文日期
            strDate = ChangeTStringToTDateString(ChangeWStringToTString(grd1.TextMatrix(i, 11)))
         End If
         '承辦人統計
         If opt1(1).Value = True Then
            cnnConnection.Execute "INSERT INTO R030409 values(" & _
            "'" & grd1.TextMatrix(i, 1) & "'" & _
            ",'','','" & strDate & "'" & _
            ",'" & grd1.TextMatrix(i, 4) & "-" & grd1.TextMatrix(i, 5) & "-" & grd1.TextMatrix(i, 6) & "-" & grd1.TextMatrix(i, 7) & "'" & _
            ",'" & grd1.TextMatrix(i, 8) & "'" & _
            "," & grd1.TextMatrix(i, 3) & _
            ",'" & strUserNum & "','','" & grd1.TextMatrix(i, 13) & "')"
         '智權人員統計
         Else
            cnnConnection.Execute "INSERT INTO R030409 values(''," & _
            "'" & grd1.TextMatrix(i, 1) & "','" & GetPrjSalesNM(grd1.TextMatrix(i, 10)) & "'" & _
            ",'" & strDate & "'" & _
            ",'" & grd1.TextMatrix(i, 4) & "-" & grd1.TextMatrix(i, 5) & "-" & grd1.TextMatrix(i, 6) & "-" & grd1.TextMatrix(i, 7) & "'" & _
            ",'" & grd1.TextMatrix(i, 8) & "'" & _
            "," & grd1.TextMatrix(i, 3) & _
            ",'" & strUserNum & "','','" & grd1.TextMatrix(i, 13) & "')"
         End If
      Next i
   End If
End If

'Modify By Sindy 2010/10/7
If Len(txt1(8)) <> 0 Then
    If Me.txt1(8).Text = "1" Then '有點數
        cnnConnection.Execute "delete from R030409 where ID='" & strUserNum & "' and (R098007 <=0 or R098007 is null)"
    End If
End If
'2010/10/7 End

strSql = "SELECT * FROM R030409 WHERE ID='" & strUserNum & "' "
With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      If txt1(7) = "1" Or txt1(7) = "2" Then
         'InsertQueryLog (.RecordCount + dblPointTot) 'Add By Sindy 2010/10/22
         InsertQueryLog (.RecordCount + dblPointRow) 'Modify By Sindy 2012/2/3
      Else
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/22
      End If
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
         Select Case Me.txt1(7).Text
         Case "1", "2"
            ShowPrintOk
         Case Else
         End Select
      '2010/5/6 End
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/10/22
         ShowNoData
      End If
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
End With
CheckOC
PrintData
Select Case Me.txt1(7).Text
Case "1", "2"
    ShowPrintOk
Case Else
End Select
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
Page = 1
Select Case Val(txt1(7))
'edit by nickc 2008/05/21 改與 frm030407 同
Case 1, 2 '列印明細
     'edit by nickc 2007/12/17 改以選擇的功能
    If opt1(1).Value = True Then
         'Modify By Sindy 98/03/12 加工作時數
'         strSQL = "SELECT ST02, A0902, R098003, R098004, R098005, R098006, R098007, R098001, R098002 FROM R030409,STAFF, ACC090 WHERE R098001=ST01(+) And R098002=A0901(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002,R098003, R098005 "
         'Modify By Sindy 2010/9/9 增加 R098009 調整order by
         'strSql = "SELECT ST02, A0902, R098003, R098004, R098005, R098006, R098007, R098001, R098002, R098008 FROM R030409,STAFF, ACC090 WHERE R098001=ST01(+) And R098002=A0901(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002,R098003, R098005 "
         strSql = "SELECT ST02, A0902, R098003, R098004, R098005, R098006, R098007, R098001, R098002, R098008,R098009 FROM R030409,STAFF, ACC090 WHERE R098001=ST01(+) And R098002=A0901(+) AND ID='" & strUserNum & "' ORDER BY R098009, R098001, R098005 "
     Else
         'Modify By Sindy 98/03/12 加工作時數
'         strSQL = "SELECT A0902, ST02, R098003, R098004, R098005, R098006, R098007, R098002, R098001 FROM R030409,STAFF, ACC090 WHERE R098001=A0901(+) And R098002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002,R098003 , R098005 "
         'Modify By Sindy 2010/9/9 增加 R098009 調整order by
         'strSql = "SELECT A0902, ST02, R098003, R098004, R098005, R098006, R098007, R098002, R098001, R098008 FROM R030409,STAFF, ACC090 WHERE R098001=A0901(+) And R098002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R098001 ,R098002,R098003 , R098005 "
         strSql = "SELECT A0902, ST02, R098003, R098004, R098005, R098006, R098007, R098002, R098001, R098008,R098009 FROM R030409,STAFF, ACC090 WHERE R098001=A0901(+) And R098002=ST01(+) AND ID='" & strUserNum & "' ORDER BY R098009, R098002, R098005 "
     End If
     CheckOC
'edit by nickc 2008/05/21 改與 frm030407 同
'Case 2 '列印總計
'     PrintTitle
'     PrintEnd1
'     ShowLine
'     Printer.EndDoc
'     Exit Sub
Case 3, 4
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Me.Hide
    frm030409_1.Show
    frm030409_1.StrMenu
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
Case Else
End Select

With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        SavDay1 = "             " 'CheckStr(.Fields(10))&CheckStr(.Fields(7))
        SavDay(0) = "             " 'CheckStr(.Fields(8))
        SavDay(1) = "             " 'CheckStr(.Fields(1))
        SavDay(2) = "             " 'CheckStr(.Fields(2))
        'Call PrintEndPoint("", CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
        Call PrintEndPoint("", CheckStr(.Fields(10)) & CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
        Do While .EOF = False
            For i = 0 To 7 '6
                If i = 7 Then
                  strTemp(i) = CheckStr(.Fields(9))
                Else
                  strTemp(i) = CheckStr(.Fields(i))
                End If
            Next i
            'add by nickc 2007/12/17 加入承辦
            '依承辦人統計
            If opt1(1).Value = True Then
                'If "" & .Fields(7).Value <> SavDay1 Then
                If ("" & .Fields(10).Value & "" & .Fields(7).Value) <> SavDay1 Then
                    If Len(Trim(SavDay1)) <> 0 Or Len(Trim(SavDay(0))) <> 0 Or Len(Trim(SavDay(1))) <> 0 Or Len(Trim(SavDay(2))) <> 0 Then
                         If txt1(7) = "1" Then
                            If iPrint <> 2500 Then ShowLine
                            PrintEnd
                            'Call PrintEndPoint(SavDay1, CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            Call PrintEndPoint(SavDay1, CheckStr(.Fields(10)) & CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            If iPrint <> 2500 Then ShowLine
                            If iPrint <> 2500 Then iPrint = iPrint + 600
                         ElseIf txt1(7) = "2" Then
                            PrintEnd
                            'Call PrintEndPoint(SavDay1, CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            Call PrintEndPoint(SavDay1, CheckStr(.Fields(10)) & CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            ShowLine
                         End If
                    End If
                    'SavDay1 = "" & .Fields(7).Value
                    SavDay1 = "" & .Fields(10).Value & "" & .Fields(7).Value
                    SavDay(0) = "" & .Fields(7).Value
                    SavDay(1) = strTemp(1)
                    SavDay(2) = strTemp(2)
                Else
                    If SavDay(0) = "" & .Fields(7).Value Then
                        strTemp(0) = ""
                        If SavDay(1) = strTemp(1) Then
                            strTemp(1) = ""
                            If SavDay(2) = strTemp(2) Then
                                strTemp(2) = ""
                            Else
                                SavDay(2) = strTemp(2)
                            End If
                        Else
                            SavDay(1) = strTemp(1)
                            SavDay(2) = strTemp(2)
                        End If
                    Else
                        SavDay(0) = "" & .Fields(7).Value
                        SavDay(1) = strTemp(1)
                        SavDay(2) = strTemp(2)
                    End If
                End If
            '依智權人員統計
            Else
                'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
                'If "" & .Fields(8).Value <> SavDay1 Then
                'If "" & .Fields(7).Value <> SavDay1 Then
                If ("" & .Fields(10).Value & "" & .Fields(7).Value) <> SavDay1 Then
                    If Len(Trim(SavDay1)) <> 0 Or Len(Trim(SavDay(0))) <> 0 Or Len(Trim(SavDay(1))) <> 0 Or Len(Trim(SavDay(2))) <> 0 Then
                         If txt1(7) = "1" Then
                            ShowLine 'If iPrint <> 2500 Then
                            PrintEnd
                            'Call PrintEndPoint(SavDay1, CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            Call PrintEndPoint(SavDay1, CheckStr(.Fields(10)) & CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            ShowLine
                            iPrint = iPrint + 600
                         ElseIf txt1(7) = "2" Then
                            PrintEnd
                            'Call PrintEndPoint(SavDay1, CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            Call PrintEndPoint(SavDay1, CheckStr(.Fields(10)) & CheckStr(.Fields(7))) 'Add By Sindy 2010/5/4
                            ShowLine
                         End If
                    End If
                    'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
                    'SavDay1 = "" & .Fields(8).Value
                    'SavDay(0) = "" & .Fields(8).Value
                    'SavDay1 = "" & .Fields(7).Value
                    SavDay1 = "" & .Fields(10).Value & "" & .Fields(7).Value
                    SavDay(0) = "" & .Fields(7).Value
                    SavDay(1) = strTemp(1)
                    SavDay(2) = strTemp(2)
                Else
                    'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
                    'If SavDay(0) = "" & .Fields(8).Value Then
                    If SavDay(0) = "" & .Fields(7).Value Then
                        strTemp(0) = ""
                        If SavDay(1) = strTemp(1) Then
                            strTemp(1) = ""
                            If SavDay(2) = strTemp(2) Then
                                strTemp(2) = ""
                            Else
                                SavDay(2) = strTemp(2)
                            End If
                        Else
                            SavDay(1) = strTemp(1)
                            SavDay(2) = strTemp(2)
                        End If
                    Else
                        'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
                        'SavDay(0) = "" & .Fields(8).Value
                        SavDay(0) = "" & .Fields(7).Value
                        SavDay(1) = strTemp(1)
                        SavDay(2) = strTemp(2)
                    End If
                End If
            End If
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            'edit by nickc 2008/05/21 改與 frm030407 同
            If Val(txt1(7)) = 1 Then
                strTemp(0) = StrToStr(strTemp(0), 9)
                strTemp(1) = StrToStr(strTemp(1), 9)
                strTemp(2) = StrToStr(strTemp(2), 9)
                strTemp(5) = StrToStr(strTemp(5), 12)
                PrintDatil
            Else
                strTemp(0) = StrToStr(strTemp(0), 9)
                strTemp(1) = StrToStr(strTemp(1), 9)
                For i = 2 To 7 '6
                    strTemp(i) = ""
                Next i
                '2008/6/6 cancel by sonia 總計不印
                'If strTemp(0) <> "" Or strTemp(1) <> "" Then
                '    PrintDatil
                'End If
                '2008/6/6 end
                strTemp(0) = ""
                strTemp(1) = ""
            End If
            .MoveNext
        Loop
    End If
End With
If txt1(7) = "1" Then
   ShowLine
End If
PrintEnd
Call PrintEndPoint(SavDay1, "") 'Add By Sindy 2010/5/4
ShowLine
PrintEnd1
Printer.EndDoc
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "發文業績明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
'Modify By Sindy 98/03/12 增加依請款日查詢
If Trim(txt1(11)) = "1" Then '1.請款
   Printer.Print "請款日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Else
   Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'edit by nickc 2007/12/17 改以選擇的功能
If opt1(1).Value = True Then
    Printer.Print "承辦人"
Else
    Printer.Print "業務區"
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'edit by nickc 2007/12/17 改以選擇的功能
If opt1(1).Value = True Then
    Printer.Print "業務區"
Else
    Printer.Print "智權人員"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
'edit by nickc 2007/12/17 改以選擇的功能
If opt1(1).Value = True Then
    Printer.Print "智權人員"
Else
    Printer.Print "承辦人"
End If
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
'Modify By Sindy 98/03/12 增加依請款日查詢
If Trim(txt1(11)) = "1" Then '1.請款
   Printer.Print "請款日"
Else
   Printer.Print "發文日"
End If
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "點數"
'Add By Sindy 98/03/12
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "工作時數"
'98/03/12 End
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 0 To 5
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
Printer.CurrentX = PLeft(6) + 300 - Printer.TextWidth(strTemp(6))
Printer.CurrentY = iPrint
Printer.Print strTemp(6)
'Add By Sindy 98/03/12
Printer.CurrentX = PLeft(7) + 300 - Printer.TextWidth(strTemp(7))
Printer.CurrentY = iPrint
Printer.Print strTemp(7)
'98/03/12 End
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 2000
PLeft(2) = 4000
PLeft(3) = 6000
PLeft(4) = 8000
PLeft(5) = 10000
PLeft(6) = 12500
PLeft(7) = 14500 'Add By Sindy 98/03/12
End Sub

Sub PrintEnd()
Call GetPointTotSub(Right(Trim(SavDay1), 5)) 'Add By Sindy 2012/2/3 取得此人員的分配點數小計
If Len(SavDay1) = 0 Then
    'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
    '2009/2/3 add by sonia 小計加印承辦人姓名
    If opt1(1).Value = True Then
        'Modify By Sindy 98/03/12 加工作時數
        'strSQL = "SELECT COUNT(*),SUM(R098007),st02 FROM R030409,staff WHERE ID='" & strUserNum & "' AND (R098001='' or r098001 is null) and R098001=st01(+) group by st02 "
        strSql = "SELECT COUNT(*),SUM(R098007),st02,SUM(R098008) FROM R030409,staff WHERE ID='" & strUserNum & "' AND (R098001='' or r098001 is null) and R098001=st01(+) group by st02 "
    Else
        'Modify By Sindy 98/03/12 加工作時數
        'strSQL = "SELECT COUNT(*),SUM(R098007),st02 FROM R030409,staff WHERE ID='" & strUserNum & "' AND (R098002='' or r098002 is null) and R098001=st01(+) group by st02 "
        strSql = "SELECT COUNT(*),SUM(R098007),st02,SUM(R098008) FROM R030409,staff WHERE ID='" & strUserNum & "' AND (R098002='' or r098002 is null) and R098001=st01(+) group by st02 "
    End If
Else
    'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
    '2008/6/9 add by sonia 小計加印承辦人姓名
    If opt1(1).Value = True Then
        'Modify By Sindy 98/03/12 加工作時數
        'strSQL = "SELECT COUNT(*),SUM(R098007),st02 FROM R030409,staff WHERE ID='" & strUserNum & "' AND R098001='" & SavDay1 & "' and R098001=st01(+) group by st02 "
        'Modify By Sindy 2010/9/9 調整where
        'strSql = "SELECT COUNT(*),SUM(R098007),st02,SUM(R098008) FROM R030409,staff WHERE ID='" & strUserNum & "' AND R098001='" & SavDay1 & "' and R098001=st01(+) group by st02 "
        strSql = "SELECT COUNT(*),SUM(R098007),st02,SUM(R098008) FROM R030409,staff WHERE ID='" & strUserNum & "' AND R098009||R098001='" & SavDay1 & "' and R098001=st01(+) group by st02 "
    Else
        'Modify By Sindy 98/03/12 加工作時數
        'strSQL = "SELECT COUNT(*),SUM(R098007),st02 FROM R030409,staff WHERE ID='" & strUserNum & "' AND R098002='" & SavDay1 & "' and R098002=st01(+) group by st02 "
        'Modify By Sindy 2010/9/9 調整where
        'strSql = "SELECT COUNT(*),SUM(R098007),st02,SUM(R098008) FROM R030409,staff WHERE ID='" & strUserNum & "' AND R098002='" & SavDay1 & "' and R098002=st01(+) group by st02 "
        strSql = "SELECT COUNT(*),SUM(R098007),st02,SUM(R098008) FROM R030409,staff WHERE ID='" & strUserNum & "' AND R098009||R098002='" & SavDay1 & "' and R098002=st01(+) group by st02 "
    End If
End If
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
   End If
   '2008/6/9 add by sonia 小計加印承辦人姓名
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print CheckStr(adoRecordset1.Fields(2))
   Printer.CurrentX = 3500
   '2008/6/8 end
   Printer.CurrentY = iPrint
   Printer.Print "小計"
   'Modify By Sindy 2012/2/3
'   Printer.CurrentX = PLeft(5)
'   Printer.CurrentY = iPrint
'   Printer.Print "件數：" & CheckStr(adoRecordset1.Fields(0))
'   Printer.CurrentX = PLeft(6)
'   Printer.CurrentY = iPrint
'   Printer.Print "點數：" & CheckStr(adoRecordset1.Fields(1))
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "件數：" & CheckStr(adoRecordset1.Fields(0))
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "點數：" & CheckStr(adoRecordset1.Fields(1))
   If dblPointTotSub > 0 Then
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "分配點數：" & CheckStr(dblPointTotSub)
   End If
   '2012/2/3 End
   'Add By Sindy 98/03/12
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "工作時數：" & CheckStr(adoRecordset1.Fields(3))
   '98/03/12 End
   iPrint = iPrint + 300
End If
CheckOC2
ShowLine
If Len(SavDay1) = 0 Then
    'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
    If opt1(1).Value = True Then
        strSql = "SELECT R098006,COUNT(*) FROM R030409 WHERE ID='" & strUserNum & "' AND (R098001='' or r098001 is null) GROUP BY R098006 "
    Else
        strSql = "SELECT R098006,COUNT(*) FROM R030409 WHERE ID='" & strUserNum & "' AND (R098002='' or r098002 is null) GROUP BY R098006 "
    End If
Else
    'edit by nickc 2008/05/15 陳經理親自來說改成以人統計
    If opt1(1).Value = True Then
        'Modify By Sindy 2010/9/9 調整where
        'strSql = "SELECT R098006,COUNT(*) FROM R030409 WHERE ID='" & strUserNum & "' AND R098001='" & SavDay1 & "' GROUP BY R098006 "
        strSql = "SELECT R098006,COUNT(*) FROM R030409 WHERE ID='" & strUserNum & "' AND R098009||R098001='" & SavDay1 & "' GROUP BY R098006 "
    Else
        'Modify By Sindy 2010/9/9 調整where
        'strSql = "SELECT R098006,COUNT(*) FROM R030409 WHERE ID='" & strUserNum & "' AND R098002='" & SavDay1 & "' GROUP BY R098006 "
        strSql = "SELECT R098006,COUNT(*) FROM R030409 WHERE ID='" & strUserNum & "' AND R098009||R098002='" & SavDay1 & "' GROUP BY R098006 "
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
'Modify By Sindy 98/03/12
'strSQL = "SELECT COUNT(*),SUM(R098007) FROM R030409 WHERE ID='" & strUserNum & "' "
strSql = "SELECT COUNT(*),SUM(R098007),SUM(R098008) FROM R030409 WHERE ID='" & strUserNum & "' "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
   End If
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "總計"
   'Modify By Sindy 2012/2/3
'   Printer.CurrentX = PLeft(5)
'   Printer.CurrentY = iPrint
'   Printer.Print "件數：" & CheckStr(adoRecordset1.Fields(0))
'   Printer.CurrentX = PLeft(6)
'   Printer.CurrentY = iPrint
'   Printer.Print "點數：" & CheckStr(adoRecordset1.Fields(1))
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "件數：" & CheckStr(adoRecordset1.Fields(0))
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "點數：" & CheckStr(adoRecordset1.Fields(1))
   If dblPointTot > 0 Then
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "分配點數：" & CheckStr(dblPointTot)
   End If
   '2012/2/3 End
   'Add By Sindy 98/03/12
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "工作時數：" & CheckStr(adoRecordset1.Fields(2))
   '98/03/12 End
   iPrint = iPrint + 300
End If
CheckOC2
   strSql = "SELECT R098006,COUNT(*) FROM R030409 WHERE ID='" & strUserNum & "' GROUP BY R098006 "
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
   
   'Add By Sindy 2010/5/5
'   If dblPointTot > 0 Then
'      ShowLine
'      If iPrint >= 10000 Then
'         Page = Page + 1
'         Printer.NewPage
'         PrintTitle
'         PrintTitle1
'      End If
'      Printer.CurrentX = PLeft(5)
'      Printer.CurrentY = iPrint
'      Printer.Print "分配點數"
'      Printer.CurrentX = PLeft(6)
'      Printer.CurrentY = iPrint
'      Printer.Print dblPointTot
'      iPrint = iPrint + 300
'   End If
   If opt1(1).Value = True Then
      iPrint = iPrint + 600
'      If iPrint >= 10000 Then
'         Page = Page + 1
'         Printer.NewPage
'         PrintTitle
'         PrintTitle1
'      End If
      Printer.CurrentX = 0
      Printer.CurrentY = iPrint
      Printer.Print "PS.不含非個人點數"
      iPrint = iPrint + 300
   End If
   '2010/5/5 End
End Sub

Sub ShowLine()
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
   iPrint = iPrint + 300
   'Add By Sindy 2010/5/7
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
   End If
   '2010/5/7 End
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'Add by Amy 2021/07/23 開放外商CF承辦人可查詢TF案件
strSysKind = GetSystemKindByNick
If Pub_strUserST11 = "B3" Or Pub_strUserST11 = "B5" Or strUserNum = "78011" Or strUserNum = "80030" Then
    strSysKind = strSysKind & "TF,"
End If
txt1(0) = strSysKind 'Modify by Amy 2021/07/23 原:GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm030409 = Nothing
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
Select Case Index
Case 7 '列印別
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 Then
        KeyAscii = 0
    End If
Case 8 '列印內容
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(strSysKind), ",,", ""), ",") 'Modify by Amy 2021/07/23 原:GetSystemKindByNick
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
Case 4, 10
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
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
     lbl1(0) = GetPrjSalesNM(txt1(5))
     If Trim(txt1(Index)) <> "" Then
        If Trim(lbl1(0).Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 6
     lbl1(1) = GetPrjSalesNM(txt1(6))
     If Trim(txt1(Index)) <> "" Then
        If Trim(lbl1(1).Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 7
     Select Case Trim(txt1(7))
     Case "1", "2", "3", "4", ""
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 或 3 或 4 !!", , "USER 輸入錯誤")
          txt1(7).SetFocus
          txt1(7).SelStart = 0
          txt1(7).SelLength = Len(txt1(7))
          Exit Sub
     End Select
Case Else
End Select
End Sub

Sub PrintSubTot1()
Dim lngX0 As Long
   
   lngX0 = 0
   'Add By Sindy 2010/5/7
   If iPrint >= 10000 Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
   End If
   '2010/5/7 End
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
Private Sub QueryPointData()
Dim strPointSql As String, strPointSql_T As String, strPointSql_S As String
Dim strEmpCol As String, strEmpType As String
Dim strEmp As String 'Add By Sindy 2012/2/3
Dim strPointSql_L As String 'Added by Lydia 2017/01/04

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
   
   strPointSql = ""
   strPointSql_T = ""
   strPointSql_S = ""
   strPointSql_L = "" 'Added by Lydia 2017/01/04
   
   If Len(txt1(0)) <> 0 Then
      If Trim(txt1(11)) = "1" Then
         strPointSql_T = strPointSql_T + " AND a1k13 IN (" & SQLGrpStr(txt1(0), 2) & ") "
         strPointSql_S = strPointSql_S + " AND a1k13 IN (" & SQLGrpStr(txt1(0), 5) & ") "
         'Added by Lydia 2017/01/04 增加法務案件
         If StrSQL3 <> "" Then strPointSql_L = strPointSql_L + " AND a1k13 IN (" & SQLGrpStr(txt1(0), 3) & ") "
      Else
         strPointSql_T = strPointSql_T + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
         strPointSql_S = strPointSql_S + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
         'Added by Lydia 2017/01/04 增加法務案件
         If StrSQL3 <> "" Then strPointSql_L = strPointSql_L + " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") "
      End If
   End If
   If Len(Trim(txt1(1))) <> 0 Then
      If Trim(txt1(11)) = "1" Then
         '請款日(起)
         strPointSql = strPointSql & " AND A1K02 >=" & Val(ChangeTStringToWString(txt1(1))) - 19110000 & " "
      Else
         '發文日(起)
         strPointSql = strPointSql & " AND CP60=A1K01(+) AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
      End If
   End If
   If Len(Trim(txt1(2))) <> 0 Then
      If Trim(txt1(11)) = "1" Then
         '請款日(迄)
         strPointSql = strPointSql & " AND A1K02 <=" & Val(ChangeTStringToWString(txt1(2))) - 19110000 & " "
      Else
         '發文日(迄)
         strPointSql = strPointSql & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
      End If
   End If
   '業務區
   If Len(txt1(3)) <> 0 Then
       strPointSql = strPointSql & " AND a.st15>='" & txt1(3) & "' "
   End If
   If Len(txt1(4)) <> 0 Then
       strPointSql = strPointSql & " AND a.st15<='" & txt1(4) & "' "
   End If
   '智權人員
   If Len(txt1(5)) <> 0 Then
       strPointSql = strPointSql & " AND a1n04='" & txt1(5) & "' "
   End If
   '承辦人
   If Len(txt1(6)) <> 0 Then
       strPointSql = strPointSql & " AND a1n04='" & txt1(6) & "' "
   End If
   '申請國家(起)
   If Len(txt1(9)) <> 0 Then
       strPointSql_T = strPointSql_T + " AND TM10 >= '" & Me.txt1(9).Text & "' "
       strPointSql_S = strPointSql_S + " AND SP09 >= '" & Me.txt1(9).Text & "' "
       If strPointSql_L <> "" Then strPointSql_L = strPointSql_L + " AND LC15 >= '" & Me.txt1(9).Text & "' " 'Added by Lydia 2017/01/04
   End If
   '申請國家(迄)
   If Len(txt1(10)) <> 0 Then
       strPointSql_T = strPointSql_T + " AND TM10 <= '" & Me.txt1(10).Text & "z' "
       strPointSql_S = strPointSql_S + " AND SP09 <= '" & Me.txt1(10).Text & "z' "
       If strPointSql_L <> "" Then strPointSql_L = strPointSql_L + " AND LC15 <= '" & Me.txt1(10).Text & "z' " 'Added by Lydia 2017/01/04
   End If
   
   '依承辦人統計
   If opt1(1).Value = True Then
      strEmpCol = "cp14"
      strEmpType = "2"
   '依智權人員統計
   Else
      strEmpCol = "cp13"
      strEmpType = "1"
   End If
   
   'Modify by Morgan 2010/9/9 調整語法(以CP為主)
   'AND substr(a.st15,1,2)=substr(b.st15,1,2) ex:FCT-029223
   'Modify By Sindy 2010/10/6 請款和發文SQL切開,以免漏掉acc1n0資料,ex:FCT-029223/FCT-029841
   '請款
   If Trim(txt1(11)) = "1" Then
      strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,a1k13,a1k14,a1k15,a1k16,DECODE(TM10,'000',CPM03,CPM04),a1k02," & strEmpCol & ",cp27,a.st02,a.st15,' ' 點數小計 " & _
                     "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                     "WHERE a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='" & strEmpType & "' and cp09(+)=a1n03 and (a1n04<>" & strEmpCol & " or " & strEmpCol & " is null) and a1n05>0 " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
                     "AND a1k13=tm01(+) AND a1k14=tm02(+) AND a1k15=tm03(+) AND a1k16=tm04(+) " & strPointSql & strPointSql_T
                     
      'Added by Lydia 2017/01/04 增加法務案
      If strPointSql_L <> "" Then
      strSql = strSql & " UNION ALL " & _
                     "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,a1k13,a1k14,a1k15,a1k16,DECODE(LC15,'000',CPM03,CPM04),a1k02," & strEmpCol & ",cp27,a.st02,a.st15,' ' 點數小計 " & _
                     "From CASEPROGRESS,LAWCASE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                     "WHERE a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='" & strEmpType & "' and cp09(+)=a1n03 and (a1n04<>" & strEmpCol & " or " & strEmpCol & " is null) and a1n05>0 " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
                     "AND a1k13=LC01(+) AND a1k14=LC02(+) AND a1k15=LC03(+) AND a1k16=LC04(+) " & strPointSql & strPointSql_L
      End If
      'end 2017/01/04
      
      strSql = strSql & " UNION ALL " & _
                     "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,a1k13,a1k14,a1k15,a1k16,DECODE(SP09,'000',CPM03,CPM04),a1k02," & strEmpCol & ",cp27,a.st02,a.st15,' ' 點數小計 " & _
                     "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                     "WHERE a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='" & strEmpType & "' and cp09(+)=a1n03 and (a1n04<>" & strEmpCol & " or " & strEmpCol & " is null) and a1n05>0 " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
                     "AND a1k13=sp01(+) AND a1k14=sp02(+) AND a1k15=sp03(+) AND a1k16=sp04(+) " & strPointSql & strPointSql_S & _
                     "Order By st15,a1n04,a1k02,a1k13,a1k14,a1k15,a1k16 "
   '發文
   Else
      strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,DECODE(TM10,'000',CPM03,CPM04),a1k02," & strEmpCol & ",cp27,a.st02,a.st15,' ' 點數小計 " & _
                     "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                     "WHERE cp60>'X' AND " & strEmpCol & " is not null " & _
                     "AND a1n01(+)=cp60 AND a1n02='" & strEmpType & "' AND a1n03(+)=cp09 AND a1n04<>" & strEmpCol & " and a1n05>0 " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND a1k25 is null " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
                     "AND cp01=tm01(+) AND cp02=tm02(+) AND cp03=tm03(+) AND cp04=tm04(+) " & strPointSql & strPointSql_T
                     
      'Added by Lydia 2017/01/04 增加法務案
      If strPointSql_L <> "" Then
      strSql = strSql & " UNION ALL " & _
                     "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,DECODE(LC15,'000',CPM03,CPM04),a1k02," & strEmpCol & ",cp27,a.st02,a.st15,' ' 點數小計 " & _
                     "From CASEPROGRESS,LAWCASE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                     "WHERE cp60>'X' AND " & strEmpCol & " is not null " & _
                     "AND a1n01(+)=cp60 AND a1n02='" & strEmpType & "' AND a1n03(+)=cp09 AND a1n04<>" & strEmpCol & " and a1n05>0 " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND a1k25 is null " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
                     "AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) " & strPointSql & strPointSql_L
      End If
      'end 2017/01/04
      
      strSql = strSql & " UNION ALL " & _
                     "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,DECODE(SP09,'000',CPM03,CPM04),a1k02," & strEmpCol & ",cp27,a.st02,a.st15,' ' 點數小計 " & _
                     "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
                     "WHERE cp60>'X' AND " & strEmpCol & " is not null " & _
                     "AND a1n01(+)=cp60 AND a1n02='" & strEmpType & "' AND a1n03(+)=cp09 AND a1n04<>" & strEmpCol & " and a1n05>0 " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND a1k25 is null " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
                     "AND cp01=sp01(+) AND cp02=sp02(+) AND cp03=sp03(+) AND cp04=sp04(+) " & strPointSql & strPointSql_S & _
                     "Order By st15,a1n04,a1k02,cp01,cp02,cp03,cp04 "
   End If
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      intCompRow = 1 '代表有分配點數資料,從第1筆開始讀取
      dblPointRow = adoRecordset.RecordCount 'Add By Sindy 2012/2/3
      Set grd1.Recordset = adoRecordset.Clone
      
      'Modify By Sindy 2012/2/3
'      '請款
'      If Trim(txt1(11)) = "1" Then
'         strSql = "select sum(tt) from (" & _
'                        "SELECT sum(a1n05) tt " & _
'                        "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                        "WHERE a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='" & strEmpType & "' and cp09(+)=a1n03 and (a1n04<>" & strEmpCol & " or " & strEmpCol & " is null) and a1n05>0 " & _
'                        "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                        "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
'                        "AND a1k13=tm01(+) AND a1k14=tm02(+) AND a1k15=tm03(+) AND a1k16=tm04(+) " & strPointSql & strPointSql_T
'         strSql = strSql & " UNION ALL " & _
'                        "SELECT sum(a1n05) tt " & _
'                        "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                        "WHERE a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='" & strEmpType & "' and cp09(+)=a1n03 and (a1n04<>" & strEmpCol & " or " & strEmpCol & " is null) and a1n05>0 " & _
'                        "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                        "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
'                        "AND a1k13=sp01(+) AND a1k14=sp02(+) AND a1k15=sp03(+) AND a1k16=sp04(+) " & strPointSql & strPointSql_S & _
'                      ")"
'      '發文
'      Else
'         strSql = "select sum(tt) from (" & _
'                        "SELECT sum(a1n05) tt " & _
'                        "From CASEPROGRESS,TRADEMARK,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                        "WHERE cp60>'X' AND " & strEmpCol & " is not null " & _
'                        "AND a1n01(+)=cp60 AND a1n02='" & strEmpType & "' AND a1n03(+)=cp09 AND a1n04<>" & strEmpCol & " and a1n05>0 " & _
'                        "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                        "AND a1k25 is null " & _
'                        "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
'                        "AND cp01=tm01(+) AND cp02=tm02(+) AND cp03=tm03(+) AND cp04=tm04(+) " & strPointSql & strPointSql_T
'         strSql = strSql & " UNION ALL " & _
'                        "SELECT sum(a1n05) tt " & _
'                        "From CASEPROGRESS,SERVICEPRACTICE,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                        "WHERE cp60>'X' AND " & strEmpCol & " is not null " & _
'                        "AND a1n01(+)=cp60 AND a1n02='" & strEmpType & "' AND a1n03(+)=cp09 AND a1n04<>" & strEmpCol & " and a1n05>0 " & _
'                        "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                        "AND a1k25 is null " & _
'                        "AND CP12=A0901(+) AND a1n04=a.st01(+) AND " & strEmpCol & "=b.st01(+) " & _
'                        "AND cp01=sp01(+) AND cp02=sp02(+) AND cp03=sp03(+) AND cp04=sp04(+) " & strPointSql & strPointSql_S & _
'                      ")"
'      End If
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
                     grd1.TextMatrix(j, 14) = dblPointTotSub
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
               grd1.TextMatrix(j, 14) = dblPointTotSub
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
   'If grd1.TextMatrix(intCompRow, 1) = strComp1 Then
   If grd1.TextMatrix(intCompRow, 13) & grd1.TextMatrix(intCompRow, 1) = strComp1 Then
      m_strTemp1 = strComp1
      m_strTemp3 = grd1.TextMatrix(intCompRow, 12)
      Call ShowLine1(0)
      If iPrint >= 10000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle1
      End If
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print m_strTemp3 & "　請款單分配點數："
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
         'If grd1.TextMatrix(intCompRow, 1) = strComp1 Then
         If grd1.TextMatrix(intCompRow, 13) & grd1.TextMatrix(intCompRow, 1) = strComp1 Then
            m_strTemp1 = strComp1
            m_strTemp3 = grd1.TextMatrix(intCompRow, 12)
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
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
      
      'If (Val(grd1.TextMatrix(ii, 1)) >= Val(strNext1) And Val(strNext1) <> 0) Then GoTo GoToExit
      If (grd1.TextMatrix(ii, 13) & grd1.TextMatrix(ii, 1) >= strNext1 And Len(strNext1) <> 0) Then GoTo GoToExit
      
      'If (grd1.TextMatrix(ii, 1) <> m_strTemp1) Then
      If (grd1.TextMatrix(ii, 13) & grd1.TextMatrix(ii, 1) <> m_strTemp1) Then
         bNextTrue = True
         PrintEnd2
         dblSum = 0
         'one new data start
         'm_strTemp1 = Trim(grd1.TextMatrix(ii, 1))
         m_strTemp1 = Trim(grd1.TextMatrix(ii, 13)) & Trim(grd1.TextMatrix(ii, 1))
         m_strTemp3 = Trim(grd1.TextMatrix(ii, 12))
         If Page <> 1 Then
            Call ShowLine1(0)
            iPrint = iPrint + 600
         End If
         If iPrint >= 10000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle1
         End If
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         Printer.Print m_strTemp3 & "　請款單分配點數："
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
         PrintTitle1
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
      PrintTitle1
   End If
End Sub

'Add By Sindy 2010/5/5
Sub PrintEnd2()
   If dblSum > 0 Then
      Call ShowLine1(1)
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
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "發文業績明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
'Modify By Sindy 98/03/12 增加依請款日查詢
If Trim(txt1(11)) = "1" Then '1.請款
   Printer.Print "請款日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Else
   Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
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
      For i = 1 To grd1.Rows - 1
         If strEmp = grd1.TextMatrix(i, 1) Then
            dblPointTotSub = grd1.TextMatrix(i, 14)
            Exit Sub
         End If
      Next i
   End If
End Sub
