VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100101_i 
   BorderStyle     =   1  '單線固定
   Caption         =   "工時統計"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9435
   Begin VB.CommandButton CmdACSPFrate 
      Caption         =   "智財顧問專業分配比例"
      Height          =   400
      Left            =   3810
      TabIndex        =   15
      Top             =   90
      Width           =   1965
   End
   Begin VB.CommandButton CmdCR1 
      Caption         =   "相關卷號工時統計"
      Height          =   400
      Left            =   5850
      TabIndex        =   14
      Top             =   90
      Width           =   1875
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2625
      MaxLength       =   7
      TabIndex        =   1
      Top             =   600
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1530
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   1005
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   225
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   4
      Left            =   3225
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   7
      Top             =   255
      Width           =   345
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   3
      Left            =   2910
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      Top             =   255
      Width           =   225
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   2
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   5
      Top             =   255
      Width           =   765
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
      Height          =   270
      Index           =   1
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   4
      Top             =   255
      Width           =   465
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7755
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8580
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Bindings        =   "frm100101_i.frx":0000
      Height          =   4260
      Left            =   75
      TabIndex        =   13
      Top             =   1425
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   7514
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "總工作時數："
      Height          =   180
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Line Line4 
      X1              =   1725
      X2              =   3435
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2775
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   5430
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   300
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文期間："
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/04/21 Form2.0已修改; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/6/23
Option Explicit
Dim cmdState As Integer 'Added by Lydia 2021/05/12 查詢按鈕
Dim m_blnColOrderAsc As Boolean 'Added by Lydia 2021/05/12 欄位資料由小到大排序
Dim intLastRow As Integer 'Added by Lydia 2021/05/12
'Added by Lydia 2021/05/12 在Grid的欄位位置
Dim colCP09 As Integer, colCP10 As Integer, colCP113 As Integer '收文號CP09，案件性質CP10，工作時數CP113

'Remove by Lydia 2021/05/12 改寫法
'Private Sub SetDataListWidth()
'Dim ii As Integer
'
'   With grdDataList
'      .Visible = False
''2011/6/27 MODIFY BY SONIA
''      If p_bolHeaderOnly = False Then
''         .Clear
''         .Rows = 2: .Cols = 13: .FixedRows = 1: .FixedCols = 0
''         .MergeCol(0) = True
''         .MergeCells = flexMergeRestrictColumns
''      End If
'      .Cols = 13
''2011/6/27 END
'
'      If txtCode(1) <> "L" And txtCode(1) <> "FCL" And txtCode(1) <> "CFL" And _
'         txtCode(1) <> "LA" And txtCode(1) <> "LIN" Then
'         .row = 0
'         .col = 0: .ColWidth(.col) = 800: .Text = "收文日"
'         .CellAlignment = flexAlignRightCenter
'         .col = 1: .ColWidth(.col) = 960: .Text = "總收文號"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 2: .ColWidth(.col) = 950: .Text = "案件性質"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 3: .ColWidth(.col) = 960: .Text = "相關收文號 "
'         .CellAlignment = flexAlignLeftCenter
'         .col = 4: .ColWidth(.col) = 600: .Text = "承辦人"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 5: .ColWidth(.col) = 600: .Text = "智權人員"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 6: .ColWidth(.col) = 800: .Text = "本所期限"
'         .CellAlignment = flexAlignRightCenter
'         .col = 7: .ColWidth(.col) = 800: .Text = "法定期限"
'         .CellAlignment = flexAlignRightCenter
'         .col = 8: .ColWidth(.col) = 800: .Text = "發文日"
'         .CellAlignment = flexAlignRightCenter
'         .col = 9: .ColWidth(.col) = 800: .Text = "工作時數"
'         .CellAlignment = flexAlignRightCenter
'         .col = 10: .ColWidth(.col) = 980: .Text = "取消收文日"
'         .CellAlignment = flexAlignRightCenter
'         .col = 11: .ColWidth(.col) = 1000: .Text = "進度備註"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 12: .ColWidth(.col) = 0: .Text = ""
'         .CellAlignment = flexAlignLeftCenter
'      Else
'         .row = 0
'         .col = 0: .ColWidth(.col) = 800: .Text = "收文日"
'         .CellAlignment = flexAlignRightCenter
'         .col = 1: .ColWidth(.col) = 960: .Text = "總收文號"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 2: .ColWidth(.col) = 1845: .Text = "備註主題(案件性質)"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 3: .ColWidth(.col) = 960: .Text = "相對人"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 4: .ColWidth(.col) = 960: .Text = "相關收文號"
'         .CellAlignment = flexAlignLeftCenter
'         'Modified by Lydia 2015/10/05
''         .col = 5: .ColWidth(.col) = 800: .Text = "承辦律師"
'         .col = 5: .ColWidth(.col) = 800: .Text = "承辦人"
'         .CellAlignment = flexAlignLeftCenter
'         'Modified by Lydia 2015/10/05
''         .col = 6: .ColWidth(.col) = 800: .Text = "承辦法務"
'         .col = 6: .ColWidth(.col) = 800: .Text = "協辦人員"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 7: .ColWidth(.col) = 600: .Text = "智權人員"
'         .CellAlignment = flexAlignLeftCenter
'         .col = 8: .ColWidth(.col) = 800: .Text = "發文日"
'         .CellAlignment = flexAlignRightCenter
'         .col = 9: .ColWidth(.col) = 800: .Text = "工作時數"
'         .CellAlignment = flexAlignRightCenter
'         .col = 10: .ColWidth(.col) = 980: .Text = "取消收文日"
'         .CellAlignment = flexAlignRightCenter
'         .col = 11: .ColWidth(.col) = 800: .Text = "本所期限"
'         .CellAlignment = flexAlignRightCenter
'         .col = 12: .ColWidth(.col) = 800: .Text = "法定期限"
'         .CellAlignment = flexAlignRightCenter
'      End If
'
'      For ii = 13 To .Cols - 1
'         .ColWidth(ii) = 0
'      Next
'      .Refresh
'      .Visible = True
'   End With
'End Sub

'Added by Lydia 2021/05/12
Private Sub SetGrd(Optional ByVal bolReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
Dim ii As Integer

    If cmdState = 1 Or (txtCode(1) <> "L" And txtCode(1) <> "FCL" And txtCode(1) <> "CFL" And _
         txtCode(1) <> "LA" And txtCode(1) <> "LIN") Then
         arrGridHeadText = Array("V", "收文日", "本所案號", "總收文號", "案件性質", "相關收文號", "承辦人", "智權人員", "本所期限", "法定期限", "發文日", "工作時數", "取消收文日", "進度備註", "CP10")
         If cmdState = 1 Then '相關卷號工時統計：顯示本所案號
              arrGridHeadWidth = Array(260, 800, 1000, 960, 950, 1000, 800, 800, 800, 800, 800, 800, 980, 1200, 0)
         Else
              arrGridHeadWidth = Array(260, 800, 0, 960, 950, 1000, 800, 800, 800, 800, 800, 800, 980, 1200)
         End If
    Else
         arrGridHeadText = Array("V", "收文日", "本所案號", "總收文號", "備註主題(案件性質)", "相對人", "相關收文號", "承辦人", "協辦人員", "智權人員", "發文日", "工作時數", "取消收文日", "本所期限", "法定期限", "CP10")
         arrGridHeadWidth = Array(260, 800, 0, 960, 1845, 960, 960, 800, 800, 800, 800, 980, 800, 800, 800, 0)
    End If
    
   '因為查詢條件不同，欄位不同
   colCP09 = PUB_MGridGetId("總收文號", GrdDataList)
   colCP10 = PUB_MGridGetId("CP10", GrdDataList)
   colCP113 = PUB_MGridGetId("工作時數", GrdDataList)
   
   GrdDataList.Visible = False
   GrdDataList.Cols = UBound(arrGridHeadText) + 1
   
   If bolReset = True Then
         GrdDataList.Clear
         GrdDataList.Rows = 2
   End If
   
    For iRow = 0 To GrdDataList.Cols - 1
       GrdDataList.row = 0
       GrdDataList.col = iRow
       GrdDataList.Text = arrGridHeadText(iRow)
       If iRow <= UBound(arrGridHeadWidth) Then
            GrdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
       Else
            GrdDataList.ColWidth(iRow) = 0
       End If
       GrdDataList.CellAlignment = flexAlignCenterCenter
    Next
    
   For intI = 1 To GrdDataList.Rows - 1
        GrdDataList.row = intI
        For iRow = 0 To GrdDataList.Cols - 1
           GrdDataList.col = iRow
           GrdDataList.CellBackColor = QBColor(15)
        Next iRow
   Next intI
    GrdDataList.Visible = True
   
End Sub

Private Function doQuery() As Boolean
Dim i As Integer, douTot As Double
Dim intCaseKind As Integer 'Added by Lydia 2021/05/12

On Error GoTo ErrHnd
   
   'Added by Lydia 2021/05/12 ACS智財顧問專業分配比例管制：本所案號若為ACS且有收文過智財顧問112，再增加智財顧問專業分配比例，點選ACS之智財顧問112進度才可按此按鈕
   If txtCode(1).Tag = "" Then
        CmdACSPFrate.Visible = False
        If strSrvDate(1) >= ACS_PFrateStart And txtCode(1) = "ACS" Then
            strExc(1) = txtCode(1)
            strExc(2) = txtCode(2)
            strExc(3) = txtCode(3)
            strExc(4) = txtCode(4)
            If PUB_ChkCPExist(strExc, "112") = True Then
                 CmdACSPFrate.Visible = True
            End If
        End If
        txtCode(1).Tag = txtCode(1).Text
   End If
   'end 2021/05/12
   
   Call SetGrd(True) 'Added by Lydia 2021/05/12 清空資料
   
   Screen.MousePointer = vbHourglass
   doQuery = False
   
   '2011/6/27 add by sonia
   strExc(0) = ""
   If Text2 <> "" Then strExc(0) = " and cp05>=" & DBDATE(Text1) & " and cp05<=" & DBDATE(Text2) & " "
   '2011/6/27 end
   
   'Added by Lydia 2021/05/12 ACS智財顧問專業分配比例管制：相關卷號工時統計->以畫面上之本所案號抓出所有相關卷號(CaseRelation1)，再抓出收文日期符合畫面收文期間(可空白)條件--(取消)且工作時數CP113有值的所有案號進度
   If cmdState = 1 Then
       Call Proc_R100101_i(txtCode(1), txtCode(2), txtCode(3), txtCode(4))
       'Memo by Lydia 2021/05/25 不用限制有工作時數，相關卷號的收文全部顯示 -- and nvl(cp113,0) > 0
       'Added by Lydia 2021/07/01 改為只有ACS案(全部收文)不限制CP113有值，其他案都是CP113有值或是未發文未取消收文的才出現。
       strExc(0) = strExc(0) & " and ((cp01='ACS' and cp159=0) or (nvl(cp113,0) > 0) or (cp158=0 and cp159=0) ) "
       strSql = "select '' as V, substr(sqldatet(cp05),1,10) 收文日,cp01||'-'||cp02||decode(cp03,'0',null,'-'||cp03)||decode(cp04,'00',null,'-'||cp04) 本所案號, " & _
                   "cp09 總收文號, decode(nvl(pa09,nvl(tm10,nvl(lc15,nvl(sp09,'000')))),'000',nvl(cpm03,cp10),nvl(cpm04,cp10)) 案件性質, " & _
                   "cp43 相關收文號,nvl(s1.st02,cp14) 承辦人,nvl(s2.st02,cp13) 智權人員,substr(sqldatet(cp06),1,10)  本所期限, " & _
                   "substr(sqldatet(cp07),1,10)  法定期限,substr(sqldatet(cp27),1,10)  發文日,cp113 工作時數,substr(sqldatet(cp57),1,10)  取消收文日,substr(cp64,1,500)  進度備註,CP10 " & _
                   "from (select cp05,cp01,cp02,cp03,cp04,cp09,cp10,cp43,cp14,cp13,cp06,cp07,cp27,cp113,cp57,cp64 " & _
                   "from caseprogress where (cp01,cp02,cp03,cp04) in (select r001001, r001002, r001003, r001004 from r100101_i where id='" & strUserNum & "' ) " & strExc(0) & _
                   ") vt1, staff s1, staff s2,casepropertymap ,patent,trademark,lawcase,servicepractice,hirecase " & _
                   "where cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & _
                   "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                   "and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
                   "and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
                   "and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) " & _
                   "and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) "
        strSql = strSql & "order by 本所案號 asc, 總收文號 asc "
   Else
   'end 2021/05/12
        '2011/6/27 modify by sonia 依系統類別區分,原用union
        Select Case txtCode(1)
           Case "CFP", "FCP", "P"   '專利
              'Modified by Lydia 2021/05/12 +V,本所案號,CP10; 限制欄位長度substr
              strSql = "SELECT '' as V,substr(sqldatet(cp05),1,10) 收文日,CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) 本所案號," & _
                          "cp09 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) 案件性質,cp43 相關收文號,NVL(S1.ST02,CP14) 承辦人,NVL(S2.ST02,CP13) 智權人員," & _
                          "substr(sqldatet(cp06),1,10) 本所期限,substr(sqldatet(cp07),1,10) 法定期限,substr(sqldatet(cp27),1,10) 發文日,cp113 工作時數,substr(sqldatet(cp57),1,10) 取消收文日,substr(cp64,1,500) 進度備註,CP10 " & _
                          "FROM caseprogress,staff s1,staff s2,casepropertymap,patent " & _
                          "WHERE cp01='" & txtCode(1) & "' and cp02='" & txtCode(2) & "' and cp03='" & txtCode(3) & "' and cp04='" & txtCode(4) & "' " & _
                          "and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 " & _
                          strExc(0) & _
                          "and cp14=s1.st01(+) " & "and cp13=s2.st01(+) " & _
                          "and cpm01(+)=cp01 and cpm02(+)=cp10 "
           'Added by Lydia 2021/05/25 ACS案的顯示方式和專利案相同
           Case "ACS"
              strSql = "SELECT '' as V,substr(sqldatet(cp05),1,10) 收文日,CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) 本所案號," & _
                          "cp09 總收文號,NVL(DECODE(lc15,'000',CPM03,CPM04),CP10) 案件性質,cp43 相關收文號,NVL(S1.ST02,CP14) 承辦人,NVL(S2.ST02,CP13) 智權人員," & _
                          "substr(sqldatet(cp06),1,10) 本所期限,substr(sqldatet(cp07),1,10) 法定期限,substr(sqldatet(cp27),1,10) 發文日,cp113 工作時數,substr(sqldatet(cp57),1,10) 取消收文日,substr(cp64,1,500) 進度備註,CP10 " & _
                          "FROM caseprogress,staff s1,staff s2,casepropertymap,lawcase " & _
                          "WHERE cp01='" & txtCode(1) & "' and cp02='" & txtCode(2) & "' and cp03='" & txtCode(3) & "' and cp04='" & txtCode(4) & "' " & _
                          "and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 " & _
                          strExc(0) & _
                          "and cp14=s1.st01(+) " & "and cp13=s2.st01(+) " & _
                          "and cpm01(+)=cp01 and cpm02(+)=cp10 "
           'end 2021/05/25
           Case "CFT", "FCT", "T", "TF"   '商標
              'Modified by Lydia 2021/05/12 +V,本所案號,CP10; 限制欄位長度substr
              strSql = "SELECT '' as V,substr(sqldatet(cp05),1,10) 收文日,CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) 本所案號" & _
                          ",cp09 總收文號,NVL(DECODE(tm10,'000',CPM03,CPM04),CP10) 案件性質,cp43 相關收文號,NVL(S1.ST02,CP14) 承辦人,NVL(S2.ST02,CP13) 智權人員," & _
                          "substr(sqldatet(cp06),1,10) 本所期限,substr(sqldatet(cp07),1,10) 法定期限,substr(sqldatet(cp27),1,10) 發文日,cp113 工作時數,substr(sqldatet(cp57),1,10) 取消收文日,substr(cp64,1,500) 進度備註,CP10 " & _
                          "FROM caseprogress,staff s1,staff s2,casepropertymap,trademark " & _
                          "WHERE cp01='" & txtCode(1) & "' and cp02='" & txtCode(2) & "' and cp03='" & txtCode(3) & "' and cp04='" & txtCode(4) & "' " & _
                          "and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 " & _
                          strExc(0) & _
                          "and cp14=s1.st01(+) " & "and cp13=s2.st01(+) " & _
                          "and cpm01(+)=cp01 and cpm02(+)=cp10 "
           'modify by sonia 2021/4/7 +ACS
           'Modified by Lydia 2021/05/25 拿掉ACS案
           Case "CFL", "FCL", "L", "LIN"         '法務
              'Modified by Lydia 2015/10/05
              'Modified by Lydia 2021/05/12 +V,本所案號,CP10; 限制欄位長度substr
              strSql = "SELECT '' as V,substr(sqldatet(cp05),1,10) 收文日,CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) 本所案號" & _
                          ",cp09 總收文號,substr(CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',1,500) 案件性質,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相對人," & _
                          "cp43 相關收文號,NVL(S1.ST02,CP14) 承辦人,decode(CP29,S3.ST01,S3.ST02) 協辦人員,NVL(S2.ST02,CP13) 智權人員," & _
                          "substr(sqldatet(cp27),1,10) 發文日,cp113 工作時數,substr(sqldatet(cp57),1,10) 取消收文日,substr(sqldatet(cp06),1,10) 本所期限,substr(sqldatet(cp07),1,10) 法定期限,CP10 " & _
                          "FROM caseprogress,staff s1,staff s2,staff s3,casepropertymap,lawcase,CUSTOMER " & _
                          "WHERE cp01='" & txtCode(1) & "' and cp02='" & txtCode(2) & "' and cp03='" & txtCode(3) & "' and cp04='" & txtCode(4) & "' " & _
                          "and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 " & _
                          strExc(0) & _
                          "and cp14=s1.st01(+) " & "and cp13=s2.st01(+) and cp29=s3.st01(+) " & _
                          "and cpm01(+)=cp01 and cpm02(+)=cp10 AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) "
           Case "LA"                      '顧問
              'Modified by Lydia 2015/10/05
              'Modified by Lydia 2021/05/12 +V,本所案號,CP10; 限制欄位長度substr
              strSql = "SELECT '' as V,substr(sqldatet(cp05),1,10) 收文日,CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) 本所案號" & _
                           ",cp09 總收文號,decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') 案件性質,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相對人," & _
                           "cp43 相關收文號,NVL(S1.ST02,CP14) 承辦人,decode(CP29,S3.ST01,S3.ST02) 協辦人員,NVL(S2.ST02,CP13) 智權人員," & _
                           "substr(sqldatet(cp27),1,10) 發文日,cp113 工作時數,substr(sqldatet(cp57),1,10) 取消收文日,substr(sqldatet(cp06),1,10) 本所期限,substr(sqldatet(cp07),1,10) 法定期限,CP10 " & _
                          "FROM caseprogress,staff s1,staff s2,staff s3,casepropertymap,hirecase,CUSTOMER " & _
                          "WHERE cp01='" & txtCode(1) & "' and cp02='" & txtCode(2) & "' and cp03='" & txtCode(3) & "' and cp04='" & txtCode(4) & "' " & _
                          "and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 " & _
                          strExc(0) & _
                          "and cp14=s1.st01(+) " & "and cp13=s2.st01(+) and cp29=s3.st01(+) " & _
                          "and cpm01(+)=cp01 and cpm02(+)=cp10 AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) "
           Case Else                  '服務
              'Modified by Lydia 2021/05/12 +V,本所案號,CP10 ; 限制欄位長度substr
              strSql = "SELECT '' as V,substr(sqldatet(cp05),1,10) 收文日,CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) 本所案號" & _
                          ",cp09 總收文號,NVL(DECODE(sp09,'000',CPM03,CPM04),CP10) 案件性質,cp43 相關收文號,NVL(S1.ST02,CP14) 承辦人,NVL(S2.ST02,CP13) 智權人員," & _
                          "substr(sqldatet(cp06),1,10) 本所期限,substr(sqldatet(cp07),1,10) 法定期限,substr(sqldatet(cp27),1,10) 發文日,cp113 工作時數,substr(sqldatet(cp57),1,10) 取消收文日,substr(cp64,1,500) 進度備註,CP10 " & _
                          "FROM caseprogress,staff s1,staff s2,casepropertymap,servicepractice " & _
                          "WHERE cp01='" & txtCode(1) & "' and cp02='" & txtCode(2) & "' and cp03='" & txtCode(3) & "' and cp04='" & txtCode(4) & "' " & _
                          "and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 " & _
                          strExc(0) & _
                          "and cp14=s1.st01(+) " & "and cp13=s2.st01(+) " & _
                          "and cpm01(+)=cp01 and cpm02(+)=cp10 "
        End Select
        strSql = strSql & "order by 總收文號 asc "
   End If 'Added by Lydia 2021/05/12
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   Text3 = "": douTot = 0
   If intI = 1 Then
      GrdDataList.FixedCols = 0 'Added by Lydia 2021/05/12
      Set GrdDataList.Recordset = RsTemp.Clone
      'Modified by Lydia 2021/05/12
      'Call SetDataListWidth
      Call SetGrd
      GrdDataList.FixedCols = 5
      'end 2021/05/12
      
      '計算總工作時數
      Me.Enabled = False
      For i = 1 To GrdDataList.Rows - 1
         'Modified by Lydia 2021/05/12 改變數
         'grdDataList.col = 9
         'grdDataList.row = i
         'douTot = douTot + Val(grdDataList.TextMatrix(i, 9))
         GrdDataList.col = colCP113
         GrdDataList.row = i
         douTot = douTot + Val(GrdDataList.TextMatrix(i, colCP113))
         'end 2021/05/12
      Next i
      Text3 = douTot
      Me.Enabled = True
   Else
      MsgBox "無符合資料！", vbInformation
      Screen.MousePointer = vbDefault
      Exit Function
   End If
   
   doQuery = True
ErrHnd:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub cmdExit_Click()
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
End Sub

Public Sub cmdSearch_Click()
   '2011/6/27 modify by sonia 原控制不可空白,改為必須同時空白或同時輸入
   If (Text1 = "" And Text2 <> "") Or (Text1 <> "" And Text2 = "") Then
      MsgBox "收文期間條件必須同時空白或同時輸入！", vbInformation
      Text1.SetFocus
      Exit Sub
   End If
   'If Text2 = "" Then
   '   MsgBox "收文迄止日期不可空白！", vbInformation
   '   Text2.SetFocus
   '   Exit Sub
   'End If
   '2011/6/27 end
   
   'Screen.MousePointer = vbHourglass
   
   cmdState = 0 'Added by Lydia 2021/05/12
   doQuery
   'Screen.MousePointer = vbDefault
End Sub

Public Sub QueryData()
   Text1 = ""
   Text2 = ""
   If txtCode(1) = "LA" Then
      strSql = "SELECT cp53,cp54 FROM CaseProgress WHERE cp01='" & txtCode(1) & "' and cp02='" & txtCode(2) & "' and cp03='" & txtCode(3) & "' and cp04='" & txtCode(4) & "' and cp10='0' and cp27 is null and cp57 is null order by cp05 desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val("" & RsTemp("cp53")) > 0 Then
            Text1 = ChangeWStringToTString("" & RsTemp("cp53"))
         End If
         If Val("" & RsTemp("cp54")) > 0 Then
            Text2 = ChangeWStringToTString("" & RsTemp("cp54"))
         End If
         If Text1 <> "" And Text2 <> "" Then
            doQuery
         'Added by Lydia 2021/05/25 彈提醒
         Else
             MsgBox "本案尚未輸入聘任期間！", vbInformation + vbOKOnly
         'end 2021/05/25
         End If
      End If
   '2011/6/27 add by sonia
   Else
      doQuery
   '2011/6/27 end
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'Modified by Lydia 2021/05/12
   'SetDataListWidth
   Call SetGrd(True) '清空資料
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MenuEnabled
   Set frm100101_i = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text1_LostFocus()
   If PUB_CheckKeyInDate(Me.Text1) = -1 Then
      Me.Text1.SetFocus
      Text1_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Text2_LostFocus()
   If PUB_CheckKeyInDate(Me.Text2) = -1 Then
      Me.Text2.SetFocus
      Text2_GotFocus
      Exit Sub
   End If
   If Not nickChgRan(Text1, Text2, "收文期間") Then
      Text1.SetFocus
      Text1_GotFocus
      Exit Sub
   End If
End Sub

'Added by Lydia 2021/05/12 相關卷號工時統計
Private Sub CmdCR1_Click()

   If (Text1 = "" And Text2 <> "") Or (Text1 <> "" And Text2 = "") Then
      MsgBox "收文期間條件必須同時空白或同時輸入！", vbInformation
      Text1.SetFocus
      Exit Sub
   End If
   
   cmdState = 1
   doQuery
   
End Sub

'Added by Lydia 2021/05/12 智財顧問專業分配比例
Private Sub CmdACSPFrate_Click()
Dim intP As Integer, intJ As Integer

    If PUB_CheckFormExist("frm081031_3") Then
        MsgBox "請先關閉〔智財顧問專業分配比例〕畫面！"
        Exit Sub
    End If
    
    Me.Enabled = False

    For intP = 1 To GrdDataList.Rows - 1
        GrdDataList.col = 0
        GrdDataList.row = intP
        If Trim(GrdDataList.Text) = "V" Then
           GrdDataList.col = 0
           GrdDataList.Text = ""
           For intJ = 0 To GrdDataList.Cols - 1
                GrdDataList.col = intJ
                GrdDataList.CellBackColor = QBColor(15)
           Next intJ
           GrdDataList.col = 2
           
           If "" & GrdDataList.TextMatrix(intP, colCP09) <> "" And "" & GrdDataList.TextMatrix(intP, colCP10) = "112" Then
               Screen.MousePointer = vbHourglass
               Call frm081031_3.SetParent(Me, "" & GrdDataList.TextMatrix(intP, colCP09), "Q") '僅供查詢
               frm081031_3.Show
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
           End If
        End If
    Next intP
    Me.Enabled = True
End Sub

'Added by Lydia 2021/05/12
Private Sub GrdDataList_Click()
   GridClick GrdDataList, intLastRow, 0, 0, 0, "V"
End Sub

'Added by Lydia 2021/05/12
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GrdDataList, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GrdDataList.col = nCol
   GrdDataList.row = nRow
   If Me.GrdDataList.row < 1 And Me.GrdDataList.Text <> "V" Then
      If InStr("工作時數", Me.GrdDataList.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.GrdDataList.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GrdDataList.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GrdDataList.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GrdDataList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
