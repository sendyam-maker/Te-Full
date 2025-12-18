VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040329 
   BorderStyle     =   1  '單線固定
   Caption         =   "帳單輸入三個月未結匯明細"
   ClientHeight    =   5592
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8448
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5592
   ScaleWidth      =   8448
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   168
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   12
         Top             =   255
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4515
      Left            =   60
      TabIndex        =   8
      Top             =   1020
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   7959
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   2
      Left            =   7425
      TabIndex        =   7
      Top             =   90
      Width           =   885
   End
   Begin VB.CheckBox Check1 
      Caption         =   "內商"
      Height          =   225
      Index           =   4
      Left            =   5805
      TabIndex        =   6
      Top             =   750
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CFL"
      Height          =   225
      Index           =   3
      Left            =   4380
      TabIndex        =   5
      Top             =   750
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CFT、CFC、S"
      Height          =   225
      Index           =   2
      Left            =   2955
      TabIndex        =   4
      Top             =   750
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "P、PS"
      Height          =   225
      Index           =   1
      Left            =   1530
      TabIndex        =   3
      Top             =   750
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CFP、CPS"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   750
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   405
      Index           =   1
      Left            =   6502
      TabIndex        =   1
      Top             =   90
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5580
      TabIndex        =   0
      Top             =   90
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   810
      Left            =   9435
      TabIndex        =   9
      Top             =   2325
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1439
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
End
Attribute VB_Name = "frm040329"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
'Create by nickc 2005/04/27
Option Explicit
'列印控制
Dim PLeft(0 To 11) As Integer
Dim strTemp(0 To 11) As String
Dim iPrint As Integer
Dim Page As Integer
Dim SeekTemp1 As String, SeekTemp2 As String, SeekPrint As Integer, SeekPrintL As Integer, SeekTempPrint As String
Dim i As Integer, j As Integer, strTemp1 As Variant, strTemp2 As Variant, s As Integer
Dim strPrinter As String 'Added by Morgan 2025/3/21

Sub SetGrid()
With grd1
      .Cols = 12
      .row = 0
      .col = 0: .Text = "輸入日期"
      .ColWidth(0) = 855
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .Text = "本所案號"
      .ColWidth(1) = 1425
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .Text = "申請國家"
      .ColWidth(2) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .Text = "案件性質"
      .ColWidth(3) = 1170
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .Text = "帳單編號"
      .ColWidth(4) = 1035
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .Text = "代理人 D/N No."
      .ColWidth(5) = 1400
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .Text = "帳單日期"
      .ColWidth(6) = 855
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .Text = "幣別"
      .ColWidth(7) = 465
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .Text = "金額"
      .ColWidth(8) = 705
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .Text = "代理人編號"
      .ColWidth(9) = 1050
      .CellAlignment = flexAlignCenterCenter
      .col = 10: .Text = "代理人"
      .ColWidth(10) = 2085
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .Text = "輸入人員"
      .ColWidth(11) = 855
      .CellAlignment = flexAlignCenterCenter

End With
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim bolDataIsOk As Boolean
Dim StrSqlBy040329 As String
Select Case Index
Case 0    '查詢
         grd1.Clear
         grd1.Rows = 2
         SetGrid
         StrSqlBy040329 = ""
         strSQL1 = "''"   '專利
         strSQL2 = "''"   '商標
         StrSQL3 = "''"   '法務
         StrSQL4 = "''"   '服務
         If Check1(0).Value = 1 Then
            strSQL1 = strSQL1 & ",'CFP'"
            StrSQL4 = StrSQL4 & ",'CPS'"
         End If
         If Check1(1).Value = 1 Then
            strSQL1 = strSQL1 & ",'P'"
            StrSQL4 = StrSQL4 & ",'PS'"
         End If
         If Check1(2).Value = 1 Then
            strSQL2 = strSQL2 & ",'CFT'"
            StrSQL4 = StrSQL4 & ",'CFC','S'"
         End If
         If Check1(3).Value = 1 Then
            StrSQL3 = StrSQL3 & ",'CFL'"
         End If
         If Check1(4).Value = 1 Then
            strSQL2 = strSQL2 & ",'T','TF'"
            StrSQL4 = StrSQL4 & "," & GetAddStr(GetSystemKindByNickTformSP)
         End If
         bolDataIsOk = False
         'Modify by Morgan 2010/8/18 百年蟲
         '" & SqlDateT("A1514") & "-->substrb(' '||sqldatet(A1514),-9)
         '" & SqlDateT("A1502") & "-->substrb(' '||sqldatet(A1502),-9)
         If strSQL1 <> "''" Then
            StrSqlBy040329 = "select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(pa09,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(pa09,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02 "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, patent, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & strSQL1 & ") "
            bolDataIsOk = True
         End If
         If strSQL2 <> "''" Then
            If bolDataIsOk = True Then
               StrSqlBy040329 = StrSqlBy040329 & " union "
            End If
            StrSqlBy040329 = StrSqlBy040329 & " select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(tm10,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(tm10,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02 "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, trademark, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm10=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & strSQL2 & ") "
            bolDataIsOk = True
         End If
         If StrSQL3 <> "''" Then
            If bolDataIsOk = True Then
               StrSqlBy040329 = StrSqlBy040329 & " union "
            End If
            StrSqlBy040329 = StrSqlBy040329 & " select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(lc15,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(lc15,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02 "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, lawcase, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and lc15=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & StrSQL3 & ") "
            bolDataIsOk = True
         End If
         If StrSQL4 <> "''" Then
            If bolDataIsOk = True Then
               StrSqlBy040329 = StrSqlBy040329 & " union "
            End If
            StrSqlBy040329 = StrSqlBy040329 & " select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(sp09,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(sp09,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02 "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, servicepractice, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp09=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & StrSQL4 & ") "
            bolDataIsOk = True
         End If
         If StrSqlBy040329 = "" Then
            cmdok(1).Enabled = False
            MsgBox "請最少勾一種系統別！", vbCritical, "錯誤！"
            Exit Sub
         End If
         StrSqlBy040329 = StrSqlBy040329 & " order by A1503,1,A1501,2"
         Screen.MousePointer = vbHourglass
         grd1.MousePointer = flexArrowHourGlass
         CheckOC3
         With AdoRecordSet3
               .CursorLocation = adUseClient
               .Open StrSqlBy040329, cnnConnection, adOpenStatic, adLockReadOnly
               If .RecordCount <> 0 Then
                  cmdok(1).Enabled = True
               Else
                  ShowNoData
                  cmdok(1).Enabled = False
               End If
               Set grd1.Recordset = AdoRecordSet3
               SetGrid
               Screen.MousePointer = vbDefault
               grd1.MousePointer = flexDefault
         End With
         CheckOC3
Case 1    '列印
         
         'Removed by Moran 2025/3/21
         'If Combo1.ListIndex >= SeekPrint Then
         '   j = Combo1.ListIndex + 1
         'Else
         '   j = Combo1.ListIndex
         'End If
         'Set Printer = Printers(j)
         'end 2025/3/21
         
         StrSqlBy040329 = ""
         strSQL1 = "''"   '專利
         strSQL2 = "''"   '商標
         StrSQL3 = "''"   '法務
         StrSQL4 = "''"   '服務
         If Check1(0).Value = 1 Then
            strSQL1 = strSQL1 & ",'CFP'"
            StrSQL4 = StrSQL4 & ",'CPS'"
         End If
         If Check1(1).Value = 1 Then
            strSQL1 = strSQL1 & ",'P'"
            StrSQL4 = StrSQL4 & ",'PS'"
         End If
         If Check1(2).Value = 1 Then
            strSQL2 = strSQL2 & ",'CFT'"
            StrSQL4 = StrSQL4 & ",'CFC','S'"
         End If
         If Check1(3).Value = 1 Then
            StrSQL3 = StrSQL3 & ",'CFL'"
         End If
         If Check1(4).Value = 1 Then
            strSQL2 = strSQL2 & ",'T','TF'"
            StrSQL4 = StrSQL4 & "," & GetAddStr(GetSystemKindByNickTformSP)
         End If
         bolDataIsOk = False
         If strSQL1 <> "''" Then
            StrSqlBy040329 = "select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(pa09,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(pa09,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,decode(cp01,'CFP',1,2) as oSort "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, patent, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & strSQL1 & ") "
            bolDataIsOk = True
         End If
         If strSQL2 <> "''" Then
            If bolDataIsOk = True Then
               StrSqlBy040329 = StrSqlBy040329 & " union "
            End If
            StrSqlBy040329 = StrSqlBy040329 & " select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(tm10,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(tm10,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,decode(cp01,'CFT',3,5) as oSort  "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, trademark, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm10=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & strSQL2 & ") "
            bolDataIsOk = True
         End If
         If StrSQL3 <> "''" Then
            If bolDataIsOk = True Then
               StrSqlBy040329 = StrSqlBy040329 & " union "
            End If
            StrSqlBy040329 = StrSqlBy040329 & " select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(lc15,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(lc15,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,4 as oSort  "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, lawcase, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and lc15=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & StrSQL3 & ") "
            bolDataIsOk = True
         End If
         If StrSQL4 <> "''" Then
            If bolDataIsOk = True Then
               StrSqlBy040329 = StrSqlBy040329 & " union "
            End If
            StrSqlBy040329 = StrSqlBy040329 & " select substrb(' '||sqldatet(A1514),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,na03,NVL(DECODE(sp09,'000',CPM03,CPM04),CP10),"
            StrSqlBy040329 = StrSqlBy040329 & " decode(substr(A1501,1,1),'U',A1501,'*'||A1501) as A1501,A1504,substrb(' '||sqldatet(A1502),-9),A1505,AXF04,A1503,decode(sp09,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,decode(cp01,'CPS',1,'PS',2,'CFC',3,'S',3,5) as oSort  "
            StrSqlBy040329 = StrSqlBy040329 & " From r040329, Caseprogress, servicepractice, nation, casepropertymap, staff, fagent "
            StrSqlBy040329 = StrSqlBy040329 & " where AXF02=cp09(+) and id='" & strUserNum & "' "
            StrSqlBy040329 = StrSqlBy040329 & "  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp09=na01(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(A1503,1,8)=fa01(+) and substr(A1503,9,1)=fa02(+) "
            StrSqlBy040329 = StrSqlBy040329 & " and A1516=st01(+) and cp01 in (" & StrSQL4 & ") "
            bolDataIsOk = True
         End If
         If StrSqlBy040329 = "" Then
            cmdok(1).Enabled = False
            MsgBox "請最少勾一種系統別！", vbCritical, "錯誤！"
            Exit Sub
         End If
         StrSqlBy040329 = StrSqlBy040329 & " order by oSort,A1503,1,A1501,2 "
         Screen.MousePointer = vbHourglass
         grd1.MousePointer = flexArrowHourGlass
         CheckOC3
         With AdoRecordSet3
               .CursorLocation = adUseClient
               .Open StrSqlBy040329, cnnConnection, adOpenStatic, adLockReadOnly
               If .RecordCount <> 0 Then
                     Set grd2.Recordset = AdoRecordSet3
                     PUB_RestorePrinter Combo1 'Added by Morgan 2025/3/21
                     PrintData
                     PUB_RestorePrinter strPrinter, SeekPrintL 'Added by Morgan 2025/3/21
               End If
               Screen.MousePointer = vbDefault
               grd1.MousePointer = flexDefault
         End With
         CheckOC3
Case 2
         Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   cmdok(1).Enabled = False
   '先將資料放在暫存
   Screen.MousePointer = vbHourglass
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   'Modified by Morgan 2025/3/21
   'For i = 0 To Printers.Count - 1
   '    Set Printer = Printers(i)
   '    If Printer.DeviceName <> strSql Then
   '        Combo1.AddItem Printer.DeviceName, j
   '        j = j + 1
   '    End If
   '    If Printer.DeviceName = strSql Then
   '        SeekPrint = i
   '    End If
   'Next i
   'Combo1.Text = Combo1.List(0)
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   'end 2025/3/21
   
   DoEvents
   cnnConnection.Execute "delete from r040329 where id='" & strUserNum & "' "
   '未結匯
   cnnConnection.Execute "insert into r040329 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID) (select A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,'" & strUserNum & "' from Acc150, Acc151 where A1514<=(to_number(to_char(add_months(sysdate,-3),'yyyymmdd'))-19110000) and A1501=AXF01(+) and not exists(select * from acc190 where A1902=A1501) and A1501>='U090' and a1512 is null and A1507 is null) "
   '2009/12/4 add by sonia 先抓上述未結匯帳單之抵帳單但抵帳單輸入日期為三個月內者,否則會與下一句重覆
   cnnConnection.Execute "insert into r040329 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID) (select A1612,A1601,A1604,A1602,A1605,AXG04,AXG02,A1603,A1614,'" & strUserNum & "' from Acc160, Acc161 where A1612>(to_number(to_char(add_months(sysdate,-3),'yyyymmdd'))-19110000) and A1601=AXG01(+) and exists(select * from r040329 where AXG02=AXF02) and A1607 is null )"
   '2009/12/4 end
   '抵帳單
   'edit by nickc 2005/05/05
   'cnnConnection.Execute "insert into r040329 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID) (select A1612,A1601,A1604,A1602,A1605,AXG04,AXG02,A1603,A1614,'" & strUserNum & "' from Acc160, Acc161 where A1612<=(to_number(to_char(add_months(sysdate,-3),'yyyymmdd'))-19110000) and A1601=AXG01(+) and not exists(select * from acc190 where A1902=A1601) )"
   cnnConnection.Execute "insert into r040329 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID) (select A1612,A1601,A1604,A1602,A1605,AXG04,AXG02,A1603,A1614,'" & strUserNum & "' from Acc160, Acc161 where A1612<=(to_number(to_char(add_months(sysdate,-3),'yyyymmdd'))-19110000) and A1601=AXG01(+) and not exists(select * from acc190 where A1902=A1601) and A1607 is null )"
   '刪除有問題的單號
   'edit by nickc 2005/05/05
   'cnnConnection.Execute "delete from r040329 where id='" & strUserNum & "' and A1501 in ('V09200129','V09200130','V09200131','V09200132','V09200133','U09207356') "
   cnnConnection.Execute "delete from r040329 where id='" & strUserNum & "' and A1501='U09207356' "
   Screen.MousePointer = vbDefault
   DoEvents
   SetGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Modified by Morgan 2025/3/21
   'Set Printer = Printers(SeekPrint)
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   'end 2025/3/21
   Set frm040329 = Nothing
End Sub


'報表列印
Private Sub PrintData()
   
   Dim ii As Integer
   Dim SeekPrintKind As String
   Printer.Orientation = 2
   strTemp(0) = "'"
   strTemp(4) = ""
   strTemp(5) = ""
   strTemp(6) = ""
   strTemp(10) = ""
   strTemp(11) = ""
   GetPleft
   With grd2
      Page = 1
      SeekPrintKind = .TextMatrix(1, 12)
      PrintTitle SeekPrintKind
      For ii = 1 To .Rows - 1
         If SeekPrintKind <> .TextMatrix(ii, 12) Then
            Printer.Font.Size = 12
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.Font.Size = 10
            Printer.NewPage
            Page = Page + 1
            SeekPrintKind = .TextMatrix(ii, 12)
            PrintTitle SeekPrintKind
         End If
         Printer.Font.Size = 10
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(0) = .TextMatrix(ii, 0), "", .TextMatrix(ii, 0))
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 1)
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print StrToStr(.TextMatrix(ii, 2), 4)
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iPrint
         Printer.Print StrToStr(.TextMatrix(ii, 3), 6)
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4), "", .TextMatrix(ii, 4))
         Printer.CurrentX = PLeft(5)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(5) = .TextMatrix(ii, 5), "", .TextMatrix(ii, 5))
         Printer.CurrentX = PLeft(6)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(6) = .TextMatrix(ii, 6), "", .TextMatrix(ii, 6))
         Printer.CurrentX = PLeft(7)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 7)
         Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(.TextMatrix(ii, 8), "0.0"))
         Printer.CurrentY = iPrint
         Printer.Print Format(.TextMatrix(ii, 8), "0.0")
         Printer.CurrentX = PLeft(9)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 9)
         Printer.CurrentX = PLeft(10)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(10) = StrToStr(.TextMatrix(ii, 10), 15), "", StrToStr(.TextMatrix(ii, 10), 15))
         Printer.CurrentX = PLeft(11)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(11) = .TextMatrix(ii, 11), "", .TextMatrix(ii, 11))
         If strTemp(4) <> .TextMatrix(ii, 4) Then
            strTemp(0) = .TextMatrix(ii, 0)
            strTemp(4) = .TextMatrix(ii, 4)
            strTemp(5) = .TextMatrix(ii, 5)
            strTemp(6) = .TextMatrix(ii, 6)
            strTemp(10) = StrToStr(.TextMatrix(ii, 10), 15)
            strTemp(11) = .TextMatrix(ii, 11)
         End If
         iPrint = iPrint + 300
         If iPrint > 10000 And ii <> .Rows - 1 Then
            If SeekPrintKind = .TextMatrix(ii + 1, 12) Then
               Printer.Font.Size = 12
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               Printer.NewPage
               Page = Page + 1
               PrintTitle SeekPrintKind
            End If
         End If
      Next ii
   End With
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = PLeft(0) + 1000
   PLeft(2) = PLeft(1) + 1600
   PLeft(3) = PLeft(2) + 950
   PLeft(4) = PLeft(3) + 1300
   PLeft(5) = PLeft(4) + 1150
   PLeft(6) = PLeft(5) + 2000
   PLeft(7) = PLeft(6) + 1000
   PLeft(8) = PLeft(7) + 800
   PLeft(9) = PLeft(8) + 650
   PLeft(10) = PLeft(9) + 1200
   PLeft(11) = PLeft(10) + 2800
End Sub

Sub PrintTitle(oClass As String)
   GetPleft
   
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "帳單輸入三個月未結匯明細"

   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & GetPrjSalesNM(strUserNum)
   Printer.CurrentX = 7500
   Printer.CurrentY = iPrint
   Printer.Print "系統別：" & IIf(oClass = "1", "CFP、CPS", IIf(oClass = "2", "P、PS", IIf(oClass = "3", "CFT、CFC、S", IIf(oClass = "4", "CFL", "內商"))))
   Printer.CurrentX = 13500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")

   iPrint = iPrint + 300
   Printer.CurrentX = 13500
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
   Printer.Print "輸入日期"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "帳單編號"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "代理人 D/N No."
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "帳單日期"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "幣別"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "金額"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "代理人編號"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "代理人"
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print "輸入人員"
   iPrint = iPrint + 300
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.Font.Size = 10
   iPrint = iPrint + 300
   
End Sub
