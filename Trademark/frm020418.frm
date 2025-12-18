VERSION 5.00
Begin VB.Form frm020418 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標處國內客戶收/發文件數月報"
   ClientHeight    =   2120
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2120
   ScaleWidth      =   4180
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   180
      TabIndex        =   7
      Top             =   1500
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   9
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4620
      TabIndex        =   6
      Text            =   "ALL"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3060
      TabIndex        =   3
      Top             =   150
      Width           =   945
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2070
      TabIndex        =   2
      Top             =   150
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   1
      Left            =   2370
      MaxLength       =   7
      TabIndex        =   1
      Top             =   870
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   0
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Top             =   870
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "民國年"
      Height          =   180
      Left            =   3330
      TabIndex        =   5
      Top             =   960
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2100
      X2              =   2790
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收/發文區間："
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1125
   End
End
Attribute VB_Name = "frm020418"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
'Create by nickc 2006/07/04
Option Explicit

Dim strTemp(0 To 36) As String
Dim PLeft(0 To 36) As Integer
Dim iPrint As Integer, intPage As Integer
Dim iTop As Integer, iEnd As Integer, iTop2 As Integer

'add by nickc 2008/01/04
Dim LineH As Integer

Private Type Cal950705
    UserName As String
    MonthCountIn As Integer
    YearCountIn As Integer
    MonthCountOut As Integer
    CPM As String
    CPMCode As String
    TotMonthCountIn As Integer
    TotYearCountIn As Integer
    TotMonthCountOut As Integer
End Type
Dim oCal() As Cal950705
Dim strPrinter As String 'Add By Sindy 2015/7/3

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Trim(txt1(0).Text) = "" Then
            MsgBox "日期起不可空白！", vbExclamation
            txt1(0).SetFocus
            Exit Sub
         End If
         If Trim(txt1(0).Text) = "" Then
            MsgBox "日期迄不可空白！", vbExclamation
            txt1(0).SetFocus
            Exit Sub
         End If
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/21 清除查詢印表記錄檔欄位
         PUB_RestorePrinter Combo1.Text 'Add By Sindy 2015/7/3
         StrMenu
         PUB_RestorePrinter strPrinter 'Add By Sindy 2015/7/3
         Screen.MousePointer = vbDefault
         Me.Enabled = True
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

Sub StrMenu()
Dim strSql As String
Dim orsTmp As New ADODB.Recordset
Dim oI As Integer
Dim oJ As Integer
Dim oK As Integer
Dim ChkCPM As Boolean
Dim intRunCnt As Integer
Dim intRow As Integer

If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/10/21
   InsertQueryLog ("") 'Add By Sindy 2010/10/21
End If
LineH = 270
'先將所有商標處的人抓出來 按照部門別，編號排序
'edit by nickc 2007/11/02  踢掉 P29 不然會有北京巨京
'strSQL = "select * from staff where st03>='P20' and st03<='P29' and st04='1' and st01<'A' order by st15,st01 "
'Add By Sindy 2010/11/26
'strSql = "select * from staff where st03>='P20' and st03<='P28' and st04='1' and st01<'A' order by st15,st01 "
strSql = "select * from staff where st03>='P20' and st03<='P28' and st04='1' and st01<'F' order by st15,st01 "
Set orsTmp = New ADODB.Recordset
With orsTmp
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        ReDim oCal(.RecordCount + 1, 0) As Cal950705
        .MoveFirst
        Do While Not .EOF
            oCal(.AbsolutePosition, 0).UserName = CheckStr(.Fields("st02"))
            .MoveNext
        Loop
        oCal(.RecordCount + 1, 0).UserName = "合計"
        '抓取智權人員為商標處所承接的案件量
        'modify by sonia 2025/6/6 下面各句都要剔除有FC代理人(cp161)的案件T-252925(MCTF01但申請人國籍台灣)
        '本月收文
        strSql = " select oSys,cp12,cp13,st02,cp10,CPM,count(oNow05) TNow05,count(oAll05) TAll05,count(oNow27) TNow27 from ("
        strSql = strSql & "      SELECT '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,cp09 oNow05,null oAll05,null oNow27 FROM CASEPROGRESS,patent,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP05>=" & ChangeTStringToWString(txt1(0)) & "  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 1) & ") "
        strSql = strSql & " union select decode(cp01||tm10,'T000','1'||cp01,'T020','3'||cp01,decode(cp01,'CFT','4'||cp01,'5'||cp01)) oSys,cp10,decode(cp01,'CFT','國外商標',decode(cp01||tm10,'T020','大陸','')||nvl(cpm03,cpm04)) CPM,cp12,cp13,st02,cp09 oNow05,null oAll05,null oNow27 FROM CASEPROGRESS,trademark,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP05>=" & ChangeTStringToWString(txt1(0)) & "  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 2) & ") "
        strSql = strSql & " union select '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,cp09 oNow05,null oAll05,null oNow27 FROM CASEPROGRESS,lawcase,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  and CP05>=" & ChangeTStringToWString(txt1(0)) & "  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 3) & ") "
        strSql = strSql & " union select '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,cp09 oNow05,null oAll05,null oNow27 FROM CASEPROGRESS,hirecase,customer,staff,casepropertymap WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP05>=" & ChangeTStringToWString(txt1(0)) & "  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 4) & ") "
        strSql = strSql & " union select decode(cp01,'TS','2'||cp01,'S','2'||cp01,'5'||cp01) oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,cp09 oNow05,null oAll05,null oNow27 FROM CASEPROGRESS,servicepractice,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  and CP05>=" & ChangeTStringToWString(txt1(0)) & "  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 5) & ") "
        '累計收文
        strSql = strSql & " union select '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,cp09 oAll05,null oNow27 FROM CASEPROGRESS,patent,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP05>=" & Mid(ChangeTStringToWString(txt1(0)), 1, 4) & "0101  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 1) & ") "
        strSql = strSql & " union select decode(cp01||tm10,'T000','1'||cp01,'T020','3'||cp01,decode(cp01,'CFT','4'||cp01,'5'||cp01)) oSys,cp10,decode(cp01,'CFT','國外商標',decode(cp01||tm10,'T020','大陸','')||nvl(cpm03,cpm04)) CPM,cp12,cp13,st02,null oNow05,cp09 oAll05,null oNow27 FROM CASEPROGRESS,trademark,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP05>=" & Mid(ChangeTStringToWString(txt1(0)), 1, 4) & "0101  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 2) & ") "
        strSql = strSql & " union select '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,cp09 oAll05,null oNow27 FROM CASEPROGRESS,lawcase,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  and CP05>=" & Mid(ChangeTStringToWString(txt1(0)), 1, 4) & "0101  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 3) & ") "
        strSql = strSql & " union select '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,cp09 oAll05,null oNow27 FROM CASEPROGRESS,hirecase,customer,staff,casepropertymap WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+)  and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP05>=" & Mid(ChangeTStringToWString(txt1(0)), 1, 4) & "0101  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 4) & ") "
        strSql = strSql & " union select decode(cp01,'TS','2'||cp01,'S','2'||cp01,'5'||cp01) oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,cp09 oAll05,null oNow27 FROM CASEPROGRESS,servicepractice,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and CP05>=" & Mid(ChangeTStringToWString(txt1(0)), 1, 4) & "0101  AND CP05<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 5) & ") "
        '本月發文
        strSql = strSql & " union SELECT '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,null oAll05,cp09 oNow27 FROM CASEPROGRESS,patent,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP27>=" & ChangeTStringToWString(txt1(0)) & "  AND CP27<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 1) & ") "
        strSql = strSql & " union select decode(cp01||tm10,'T000','1'||cp01,'T020','3'||cp01,decode(cp01,'CFT','4'||cp01,'5'||cp01)) oSys,cp10,decode(cp01,'CFT','國外商標',decode(cp01||tm10,'T020','大陸','')||nvl(cpm03,cpm04)) CPM,cp12,cp13,st02,null oNow05,null oAll05,cp09 oNow27 FROM CASEPROGRESS,trademark,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP27>=" & ChangeTStringToWString(txt1(0)) & "  AND CP27<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 2) & ") "
        strSql = strSql & " union select '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,null oAll05,cp09 oNow27 FROM CASEPROGRESS,lawcase,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  and CP27>=" & ChangeTStringToWString(txt1(0)) & "  AND CP27<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 3) & ") "
        strSql = strSql & " union select '5'||cp01 oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,null oAll05,cp09 oNow27 FROM CASEPROGRESS,hirecase,customer,staff,casepropertymap WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and CP27>=" & ChangeTStringToWString(txt1(0)) & "  AND CP27<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 4) & ") "
        strSql = strSql & " union select decode(cp01,'TS','2'||cp01,'S','2'||cp01,'5'||cp01) oSys,cp10,nvl(cpm03,cpm04) CPM,cp12,cp13,st02,null oNow05,null oAll05,cp09 oNow27 FROM CASEPROGRESS,servicepractice,customer,staff,casepropertymap WHERE cp161||cp139 is null and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  and CP27>=" & ChangeTStringToWString(txt1(0)) & "  AND CP27<=" & ChangeTStringToWString(txt1(1)) & "  AND CP12>='P20'  AND CP12<='P29'  AND CP09 < 'B'  And CU10 >= '001'  And CU10 <= '009z'  and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 5) & ") "
        
        strSql = strSql & ") AAA group by oSys,cp12,cp13,st02,cp10,CPM order by oSys,cp10 " 'cp12,cp13 "
        Set orsTmp = New ADODB.Recordset
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                For oI = 1 To UBound(oCal, 1)
                    If CheckStr(.Fields("st02")) = oCal(oI, 0).UserName Then
                        ChkCPM = False
                        If UBound(oCal, 2) = 0 Then
                            ChkCPM = False
                        Else
                            '檢查有無案件性質
                            For oJ = 1 To UBound(oCal, 2)
                                If oCal(oI, oJ).CPMCode = CheckStr(.Fields("oSys")) Then
                                    If oCal(oI, oJ).CPM = CheckStr(.Fields("CPM")) Then
                                        oCal(oI, oJ).MonthCountIn = oCal(oI, oJ).MonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                        oCal(oI, oJ).MonthCountOut = oCal(oI, oJ).MonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                        oCal(oI, oJ).YearCountIn = oCal(oI, oJ).YearCountIn + Val(CheckStr(.Fields("TAll05")))
                                        oCal(UBound(oCal, 1), oJ).MonthCountIn = oCal(UBound(oCal, 1), oJ).MonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                        oCal(UBound(oCal, 1), oJ).MonthCountOut = oCal(UBound(oCal, 1), oJ).MonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                        oCal(UBound(oCal, 1), oJ).YearCountIn = oCal(UBound(oCal, 1), oJ).YearCountIn + Val(CheckStr(.Fields("TAll05")))
                                        oCal(oI, 0).TotMonthCountIn = oCal(oI, 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                        oCal(oI, 0).TotMonthCountOut = oCal(oI, 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                        oCal(oI, 0).TotYearCountIn = oCal(oI, 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                                        oCal(UBound(oCal, 1), 0).TotMonthCountIn = oCal(UBound(oCal, 1), 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                                        oCal(UBound(oCal, 1), 0).TotMonthCountOut = oCal(UBound(oCal, 1), 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                                        oCal(UBound(oCal, 1), 0).TotYearCountIn = oCal(UBound(oCal, 1), 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                                        ChkCPM = True
                                        Exit For
                                    End If
                                End If
                            Next oJ
                        End If
                        If ChkCPM = False Then
                            '還沒新增任何案件性質時
                            ReDim Preserve oCal(UBound(oCal, 1), UBound(oCal, 2) + 1) As Cal950705
                            For oK = 1 To UBound(oCal, 1)
                                oCal(oK, UBound(oCal, 2)).CPM = CheckStr(.Fields("CPM"))
                                oCal(oK, UBound(oCal, 2)).CPMCode = CheckStr(.Fields("oSys"))
                            Next oK
                            oCal(oI, UBound(oCal, 2)).MonthCountIn = CheckStr(.Fields("TNow05"))
                            oCal(oI, UBound(oCal, 2)).MonthCountOut = CheckStr(.Fields("TNow27"))
                            oCal(oI, UBound(oCal, 2)).YearCountIn = CheckStr(.Fields("TAll05"))
                            oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountIn = oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountIn + Val(CheckStr(.Fields("TNow05")))
                            oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountOut = oCal(UBound(oCal, 1), UBound(oCal, 2)).MonthCountOut + Val(CheckStr(.Fields("TNow27")))
                            oCal(UBound(oCal, 1), UBound(oCal, 2)).YearCountIn = oCal(UBound(oCal, 1), UBound(oCal, 2)).YearCountIn + Val(CheckStr(.Fields("TAll05")))
                            oCal(oI, 0).TotMonthCountIn = oCal(oI, 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                            oCal(oI, 0).TotMonthCountOut = oCal(oI, 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                            oCal(oI, 0).TotYearCountIn = oCal(oI, 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                            oCal(UBound(oCal, 1), 0).TotMonthCountIn = oCal(UBound(oCal, 1), 0).TotMonthCountIn + Val(CheckStr(.Fields("TNow05")))
                            oCal(UBound(oCal, 1), 0).TotMonthCountOut = oCal(UBound(oCal, 1), 0).TotMonthCountOut + Val(CheckStr(.Fields("TNow27")))
                            oCal(UBound(oCal, 1), 0).TotYearCountIn = oCal(UBound(oCal, 1), 0).TotYearCountIn + Val(CheckStr(.Fields("TAll05")))
                        End If
                        Exit For
                    End If
                Next oI
                .MoveNext
            Loop
        Else
            MsgBox "沒有資料可以列印！", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "商標處無人在職！沒有資料可印", vbExclamation
        Exit Sub
    End If
End With
'印表
'Add By Sindy 2011/11/1
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
Printer.PaperSize = 9  'PDF
'2011/11/1 End
'intPage = 1
'Add By Sindy 2010/6/1
intRunCnt = 1
If UBound(oCal, 1) >= 13 Then
   intRunCnt = 2
End If
For intPage = 1 To intRunCnt
   If intPage = 2 Then Printer.NewPage
'2010/6/1 End
   PrintTitle
   For oI = 0 To 36
       strTemp(oI) = ""
   Next oI
   'Add By Sindy 2010/6/1
   If intPage = 1 Then '第1頁
      If UBound(oCal, 1) < 13 Then
         intRow = UBound(oCal, 1)
      Else
         intRow = 12
      End If
   End If
   If intPage = 2 Then '第2頁
      intRow = UBound(oCal, 1) - 12
   End If
   '2010/6/1 End
   For oJ = 1 To UBound(oCal, 2)
       'Add By Sindy 2011/11/1
       'Modify By Sindy 2012/1/2
       'If oJ >= 33 Then
       If (oJ Mod 33) = 0 Then
       '2012/1/2 End
          ShowLine True
          iEnd = iPrint - 20
          PrintEndLine
          Printer.NewPage
          PrintTitle
       End If
       '2011/11/1 End
       
       strTemp(0) = oCal(1, oJ).CPM
       Printer.CurrentX = PLeft(0) + 20
       Printer.CurrentY = iPrint
       Printer.Print StrToStr(strTemp(0), 6)
       'For oI = 1 To UBound(oCal, 1)
       For oI = 1 To intRow
           'Debug.Print oCal(oI, 0).UserName; oCal(oI, oJ).CPM; oCal(oI, oJ).MonthCountIn; oCal(oI, oJ).YearCountIn; oCal(oI, oJ).MonthCountOut
           If intPage = 1 Then
               strTemp(1) = IIf(oCal(oI, oJ).MonthCountIn = 0, "", oCal(oI, oJ).MonthCountIn)
               strTemp(2) = IIf(oCal(oI, oJ).YearCountIn = 0, "", oCal(oI, oJ).YearCountIn)
               strTemp(3) = IIf(oCal(oI, oJ).MonthCountOut = 0, "", oCal(oI, oJ).MonthCountOut)
           ElseIf intPage = 2 Then
               strTemp(1) = IIf(oCal(oI + 12, oJ).MonthCountIn = 0, "", oCal(oI + 12, oJ).MonthCountIn)
               strTemp(2) = IIf(oCal(oI + 12, oJ).YearCountIn = 0, "", oCal(oI + 12, oJ).YearCountIn)
               strTemp(3) = IIf(oCal(oI + 12, oJ).MonthCountOut = 0, "", oCal(oI + 12, oJ).MonthCountOut)
           End If
           Printer.CurrentX = PLeft(1 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(1))) - 20
           Printer.CurrentY = iPrint
           Printer.Print strTemp(1)
           Printer.CurrentX = PLeft(2 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(2))) - 20
           Printer.CurrentY = iPrint
           Printer.Print strTemp(2)
           Printer.CurrentX = PLeft(3 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(3))) - 20
           Printer.CurrentY = iPrint
           Printer.Print strTemp(3)
       Next oI
       ShowLine True
       iPrint = iPrint + LineH
       If iPrint + LineH >= Printer.ScaleHeight Then
           iEnd = iPrint + LineH
           PrintEndLine
           PrintTitle
       End If
   Next oJ
   ShowLine True
   Printer.CurrentX = PLeft(0) + 20
   Printer.CurrentY = iPrint
   Printer.Print "個人總計"
   'For oI = 1 To UBound(oCal, 1)
   For oI = 1 To intRow
       'Debug.Print oCal(oI, 0).UserName; oCal(oI, oJ).CPM; oCal(oI, oJ).MonthCountIn; oCal(oI, oJ).YearCountIn; oCal(oI, oJ).MonthCountOut
       If intPage = 1 Then
            strTemp(1) = IIf(oCal(oI, 0).TotMonthCountIn = 0, "", oCal(oI, 0).TotMonthCountIn)
            strTemp(2) = IIf(oCal(oI, 0).TotYearCountIn = 0, "", oCal(oI, 0).TotYearCountIn)
            strTemp(3) = IIf(oCal(oI, 0).TotMonthCountOut = 0, "", oCal(oI, 0).TotMonthCountOut)
       ElseIf intPage = 2 Then
            strTemp(1) = IIf(oCal(oI + 12, 0).TotMonthCountIn = 0, "", oCal(oI + 12, 0).TotMonthCountIn)
            strTemp(2) = IIf(oCal(oI + 12, 0).TotYearCountIn = 0, "", oCal(oI + 12, 0).TotYearCountIn)
            strTemp(3) = IIf(oCal(oI + 12, 0).TotMonthCountOut = 0, "", oCal(oI + 12, 0).TotMonthCountOut)
       End If
       Printer.CurrentX = PLeft(1 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(1))) - 20
       Printer.CurrentY = iPrint
       Printer.Print strTemp(1)
       Printer.CurrentX = PLeft(2 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(2))) - 20
       Printer.CurrentY = iPrint
       Printer.Print strTemp(2)
       Printer.CurrentX = PLeft(3 + ((oI - 1) * 3)) + ((PLeft(3 + ((oI - 1) * 3)) - PLeft(2 + ((oI - 1) * 3))) - Printer.TextWidth(strTemp(3))) - 20
       Printer.CurrentY = iPrint
       Printer.Print strTemp(3)
   Next oI
   iPrint = iPrint + LineH
   iPrint = iPrint + 20
   ShowLine
   iEnd = iPrint - 20
   PrintEndLine
Next intPage
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2015/7/3
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2015/7/3
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2015/7/3 END
   
   Set frm020418 = Nothing
   Printer.DrawWidth = 1  '2010/12/1 add by sonia
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If CheckIsTaiwanDate(txt1(Index), True) = False Then
      txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 4 Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
End Sub

Sub PrintTitle()
Dim i As Integer
Dim intRow As Integer
GetPleft
iPrint = 0
Printer.DrawWidth = 10
'Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(GetTitleNick & "每月收/發文案件統計") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "每月收/發文案件統計"
iPrint = iPrint + Printer.TextHeight(GetTitleNick & "每月收/發文案件統計")
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 9000 + 5000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 230
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "日期：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
Printer.CurrentX = 9000 + 5000
Printer.CurrentY = iPrint
Printer.Print "頁　　數：" & Printer.Page 'intPage
iPrint = iPrint + 230
iTop = iPrint - 20
ShowLine
'Add By Sindy 2010/6/1
If intPage = 1 Then '第1頁
   If UBound(oCal, 1) < 13 Then
      intRow = UBound(oCal, 1)
   Else
      intRow = 12
   End If
End If
If intPage = 2 Then '第2頁
   intRow = UBound(oCal, 1) - 12
End If
'2010/6/1 End
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
   Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + ((((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) * 3) - Printer.TextWidth(oCal(i, 0).UserName)) / 2)
   Printer.CurrentY = iPrint
   If intPage = 1 Then
      Printer.Print oCal(i, 0).UserName
   ElseIf intPage = 2 Then
      Printer.Print oCal(i + 12, 0).UserName
   End If
Next i
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (PLeft(1), iPrint - 20)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iPrint - 20)
iTop2 = iPrint - 20
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("本")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "本"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("累")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "累"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("本")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "本"
Next i
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0) + (((PLeft(1) - PLeft(0)) - (Printer.TextWidth("案件性質"))) / 2)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("月")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "月"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("計")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "計"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("月")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "月"
Next i
iPrint = iPrint + 300
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("收")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "收"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("收")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "收"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("送")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "送"
Next i
iPrint = iPrint + 300
'For i = 1 To UBound(oCal, 1)
For i = 1 To intRow
    Printer.CurrentX = PLeft(1 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("文")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "文"
    Printer.CurrentX = PLeft(2 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("文")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "文"
    Printer.CurrentX = PLeft(3 + ((i - 1) * 3)) + (((PLeft(3 + ((i - 1) * 3)) - PLeft(2 + ((i - 1) * 3))) - Printer.TextWidth("件")) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "件"
Next i
iPrint = iPrint + 300
ShowLine True
End Sub

Sub GetPleft()
Dim i As Integer
Erase PLeft
PLeft(0) = 0
Printer.FontSize = 10 'Add By Sindy 2015/7/2
For i = 1 To 36
    PLeft(i) = 900 + ((Printer.TextWidth("　") * 2) * i - 1)
Next i
End Sub

Sub ShowLine(Optional oContent As Boolean = False)
If oContent = False Then
    Printer.DrawWidth = 10
    Printer.CurrentX = 0
    Printer.CurrentY = iPrint
    Printer.Line (PLeft(0), iPrint - 20)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iPrint - 20)
Else
    Printer.DrawWidth = 10
    Printer.Line (PLeft(0), iPrint - 20)-(PLeft(1), iPrint - 20)
    Printer.DrawWidth = 2
    Printer.Line (PLeft(1), iPrint - 20)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iPrint - 20)
End If
End Sub

Sub PrintEndLine()
Dim i As Integer
For i = 0 To UBound(PLeft)
    If (i - 1) Mod 3 = 0 Or (i - 1) Mod 3 = -1 Then
        Printer.Line (PLeft(i), iTop)-(PLeft(i), iEnd)
    Else
        Printer.DrawWidth = 2
        Printer.Line (PLeft(i), iTop2)-(PLeft(i), iEnd)
        Printer.DrawWidth = 10
    End If
Next i
Printer.Line (PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iTop)-(PLeft(UBound(PLeft)) + PLeft(UBound(PLeft)) - PLeft(UBound(PLeft) - 1), iEnd)
End Sub
