VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040205a 
   BorderStyle     =   1  '單線固定
   Caption         =   "FC收款請款點數查詢"
   ClientHeight    =   5730
   ClientLeft      =   2130
   ClientTop       =   3135
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton Cmdok 
      Caption         =   "請款單(F)"
      Height          =   350
      Index           =   2
      Left            =   5985
      TabIndex        =   3
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Index           =   1
      Left            =   8076
      TabIndex        =   2
      Top             =   12
      Width           =   1200
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "總計(&A)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   7248
      TabIndex        =   1
      Top             =   12
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   5244
      Left            =   0
      TabIndex        =   0
      Top             =   432
      Width           =   9276
      _ExtentX        =   16351
      _ExtentY        =   9260
      _Version        =   393216
      Cols            =   14
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
      _Band(0).Cols   =   14
   End
End
Attribute VB_Name = "frm040205a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; Grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2007/10/9 整理 by sonia
Option Explicit

Dim strSQL1k0 As String, strSQLCP As String
Dim strSalesArea As String   '2007/11/20 add by sonia
Dim strAccSystem As String   '2007/11/20 add by sonia
Dim s As Integer, strSql As String, strTemp As Variant, StrTest As String, i As Integer, j As Integer
Dim tmpKey  As String, TmpRow As Integer


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '總計
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         frm040205b.Show
         frm040205b.StrMenu
         Me.Hide
         Screen.MousePointer = vbDefault
         Do
            DoEvents
            If bolToEndByNick = True Then Unload Me: Exit Sub
         Loop Until Not frm040205b.Visible
         Unload frm040205b
         Me.Show
         Me.Enabled = True
      Case 1    '回前畫面
         Me.Tag = "1" 'Add By Sindy 2018/7/23
         Me.Hide
      Case 2    '查詢請款單
         tmpKey = ""
         Screen.MousePointer = vbHourglass
         If TmpRow = 0 Then
            Screen.MousePointer = vbDefault
            s = MsgBox("請選擇一筆才能顯示請款單！", , "警告！")
            Exit Sub
         End If
         GRD1.row = TmpRow
         GRD1.col = 2
         tmpKey = GRD1.Text
         'Modify By Sindy 2014/2/18
'         frm040205c.Show
'         frm040205c.StrMenu tmpKey
'         Screen.MousePointer = vbDefault
'         Me.Hide
'         Do
'            DoEvents
'            If bolToEndByNick = True Then Unload Me: Exit Sub
'         Loop Until Not frm040205c.Visible
'         Unload frm040205c
'         Me.Show
         strFormLink = Me.Name
         strItemNo = tmpKey
         Frmacc2211.Show
         Me.Enabled = False
         Screen.MousePointer = vbDefault
         '2014/2/18 END
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   TmpRow = 0
End Sub
Sub StrMenu()
   SetGridWidth
   Select Case frm040205.txt1(1)
      Case "1"
         frm040205a.Caption = "FC請款點數查詢"
         Cmdok(2).Enabled = True
         StrMenu1            '請款
      Case "2"
         frm040205a.Caption = "FC收款點數查詢"
         Cmdok(2).Enabled = False
         StrMenu2            '收款
      Case Else
   End Select
End Sub
Sub SetGridWidth()
   With GRD1
      .row = 0
      .col = 0
      .ColWidth(0) = 1550
      .Text = "本所案號"
      .col = 1
      .ColWidth(1) = 800
      .Text = "案件性質"
      .col = 2
      .ColWidth(2) = 1000
      .Text = "單據編號"
      .col = 3
      .ColWidth(3) = 800
      .Text = "單據日期"
      .col = 4
      .ColWidth(4) = 600
      .Text = "幣別"
      .col = 5
      .ColWidth(5) = 1000
      .Text = "外幣金額"
      .col = 6
      .ColWidth(6) = 1000
      .Text = "台幣金額"
      .col = 7
      .ColWidth(7) = 1000
      .Text = "規費"
      .col = 8
      .ColWidth(8) = 500
      .Text = "結清"
      .col = 9
      .ColWidth(9) = 3000
      .Text = "代理人"
      .col = 10
      .ColWidth(10) = 1000
      .Text = "翻譯費"
      .col = 11
      .ColWidth(11) = 0
      .Text = ""
      .col = 12
      .ColWidth(12) = 0
      .Text = ""
      .col = 13
      .ColWidth(13) = 0
      .Text = ""
   End With
End Sub

Sub StrMenu1()         '請款
Dim StrSQLa As String
Dim ii As Integer
Dim strA1K01 As String '請款單號

   strSQL1k0 = "": strSQLCP = "": strSalesArea = ""
   '系統類別
   '2007/11/16 modify by sonia
   'If Len(frm040205.txt1(0)) <> 0 Then
   If frm040205.txt1(0) <> "ALL" Then
   '2007/11/16 end
      strSQL1k0 = strSQL1k0 & " and a1k13 in (" & GetAddStr(frm040205.txt1(0)) & ") "
   End If
   If Trim(frm040205.txt1(0)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(0) & frm040205.txt1(0) 'Add By Sindy 2010/9/28
   End If
   pub_QL05 = pub_QL05 & ";" & frm040205.Label1(1) & "請款" 'Add By Sindy 2010/9/28
   '請款日期
   If Len(Trim(frm040205.txt1(2))) <> 0 Then
      strSQL1k0 = strSQL1k0 & " and A1K02>=" & Val(frm040205.txt1(2)) & " "
   End If
   If Len(Trim(frm040205.txt1(3))) <> 0 Then
      strSQL1k0 = strSQL1k0 & " AND A1K02<=" & Val(frm040205.txt1(3)) & " "
   End If
   If Len(Trim(frm040205.txt1(2))) <> 0 Or Len(Trim(frm040205.txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(2) & Trim(frm040205.txt1(2)) & "-" & Trim(frm040205.txt1(3)) 'Add By Sindy 2010/9/28
   End If
   '國籍
   If Len(frm040205.txt1(4)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10>='" & frm040205.txt1(4) & "' "
   End If
   If Len(frm040205.txt1(5)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10<='" & frm040205.txt1(5) & "z' "
   End If
   If Len(Trim(frm040205.txt1(4))) <> 0 Or Len(Trim(frm040205.txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(4) & Trim(frm040205.txt1(4)) & "-" & Trim(frm040205.txt1(5))  'Add By Sindy 2010/9/28
   End If
   If Len(frm040205.txt1(7)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10>='" & frm040205.txt1(7) & "' "
   End If
   If Len(frm040205.txt1(8)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10<='" & frm040205.txt1(8) & "' "
   End If
   If Len(Trim(frm040205.txt1(7))) <> 0 Or Len(Trim(frm040205.txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(7) & Trim(frm040205.txt1(7)) & "-" & Trim(frm040205.txt1(8))   'Add By Sindy 2010/9/28
   End If
   'Add by Morgan 2003/12/04  'CP13||''=是因為不會用INDEX比較快
   '智權人員
   If Len(frm040205.txt1(9)) <> 0 Then
      strSQLCP = strSQLCP & " and CP13||''='" & frm040205.txt1(9) & "' "
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(8) & Trim(frm040205.txt1(9)) & frm040205.lbl1   'Add By Sindy 2010/9/28
   End If
   '業務區
   If Len(frm040205.txt1(10)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''>='" & frm040205.txt1(10) & "' "
   End If
   If Len(frm040205.txt1(11)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''<='" & frm040205.txt1(11) & "' "
   End If
   'End 2003/12/04
   If Len(frm040205.txt1(10)) <> 0 Or Len(frm040205.txt1(11)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(9) & Trim(frm040205.txt1(10)) & "-" & Trim(frm040205.txt1(11))  'Add By Sindy 2010/9/28
   End If
   
   StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
   'Patent
   '2005/5/17 MODIFY BY SONIA 外幣金額改抓A1K08
   '2007/11/16 modify by sonia 銷帳也不抓,同一請款單只抓收文日收文號最大者
   'strSQL = "SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(PA57,'Y','＊','') AS 本所案號, DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11,2) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
   '         " FROM ACC1K0,fagent,nation,CaseProgress,Patent,CasePropertyMap,SYSTEMKIND " & _
   '         " WHERE (A1K12=0 OR A1K12 IS NULL) and " & SQLNewFag("A1K03", "fa") & " " & _
   '         " and fa10=NA01(+) And A1K01=CP60 AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLcp
   'Modify by Morgan 2010/8/11 百年蟲 " & SqlDateT("A1K02") & "--> substrb(' '||sqldatet(A1K02),-9)
'   strSql = "SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(PA57,'Y','＊','') AS 本所案號, DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11-(nvl(a1k06,0)*a1k10),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
'            " FROM fagent,nation,CaseProgress,Patent,CasePropertyMap,SYSTEMKIND, " & _
'            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP FROM ACC1K0,CASEPROGRESS " & _
'            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10) NEW " & _
'            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
'            " and fa10=NA01(+) AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   '2005/5/17 END
   'Modify By Sindy 2012/10/11 +a1k31,改a1k08 - a1k31及a1k06,a1k10相關sql
   strSql = "SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(PA57,'Y','＊','') AS 本所案號, DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 - nvl(a1k31,0) AS 外幣金額,ROUND(A1K11-nvl(a1k06,0),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
            " FROM fagent,nation,CaseProgress,Patent,CasePropertyMap,SYSTEMKIND, " & _
            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP,a1k31 FROM ACC1K0,CASEPROGRESS " & _
            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,a1k31) NEW " & _
            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
            " and fa10=NA01(+) AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   'Trademark
   '2005/5/17 MODIFY BY SONIA 外幣金額改抓A1K08
   '2007/11/16 modify by sonia 銷帳也不抓,同一請款單只抓收文日收文號最大者
   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(TM29,'Y','＊','') AS 本所案號, DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11,2) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
   '         " FROM ACC1K0,fagent,nation,CaseProgress,Trademark,CasePropertyMap,SYSTEMKIND " & _
   '         " WHERE (A1K12=0 OR A1K12 IS NULL) and " & SQLNewFag("A1K03", "fa") & " " & _
   '         " and fa10=NA01(+) And A1K01=CP60 AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLcp
   'Modify by Morgan 2010/8/11 百年蟲 " & SqlDateT("A1K02") & "--> substrb(' '||sqldatet(A1K02),-9)
'   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(TM29,'Y','＊','') AS 本所案號, DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11-(nvl(a1k06,0)*a1k10),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
'            " FROM fagent,nation,CaseProgress,Trademark,CasePropertyMap,SYSTEMKIND, " & _
'            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP FROM ACC1K0,CASEPROGRESS " & _
'            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10) NEW " & _
'            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
'            " and fa10=NA01(+) And CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   '2005/5/17 END
   'Modify By Sindy 2012/10/11 +a1k31,改a1k08 - a1k31及a1k06,a1k10相關sql
   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(TM29,'Y','＊','') AS 本所案號, DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 - nvl(a1k31,0) AS 外幣金額,ROUND(A1K11-nvl(a1k06,0),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
            " FROM fagent,nation,CaseProgress,Trademark,CasePropertyMap,SYSTEMKIND, " & _
            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP,a1k31 FROM ACC1K0,CASEPROGRESS " & _
            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,a1k31) NEW " & _
            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
            " and fa10=NA01(+) And CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   'LawCase
   '2005/5/17 MODIFY BY SONIA 外幣金額改抓A1K08
   '2007/11/16 modify by sonia 銷帳也不抓,同一請款單只抓收文日收文號最大者
   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(LC08,'Y','＊','') AS 本所案號, DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11,2) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
   '         " FROM ACC1K0,fagent,nation,CaseProgress,lawCase,CasePropertyMap,SYSTEMKIND " & _
   '         " WHERE (A1K12=0 OR A1K12 IS NULL) and " & SQLNewFag("A1K03", "fa") & " " & _
   '         " and fa10=NA01(+) And A1K01=CP60 AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLcp
   'Modify by Morgan 2010/8/11 百年蟲 " & SqlDateT("A1K02") & "--> substrb(' '||sqldatet(A1K02),-9)
'   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(LC08,'Y','＊','') AS 本所案號, DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11-(nvl(a1k06,0)*a1k10),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
'            " FROM fagent,nation,CaseProgress,lawCase,CasePropertyMap,SYSTEMKIND, " & _
'            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP FROM ACC1K0,CASEPROGRESS " & _
'            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10) NEW " & _
'            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
'            " and fa10=NA01(+) And CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   '2005/5/17 END
   'Modify By Sindy 2012/10/11 +a1k31,改a1k08 - a1k31及a1k06,a1k10相關sql
   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(LC08,'Y','＊','') AS 本所案號, DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 - nvl(a1k31,0) AS 外幣金額,ROUND(A1K11-nvl(a1k06,0),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
            " FROM fagent,nation,CaseProgress,lawCase,CasePropertyMap,SYSTEMKIND, " & _
            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP,a1k31 FROM ACC1K0,CASEPROGRESS " & _
            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,a1k31) NEW " & _
            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
            " and fa10=NA01(+) And CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   'HireCase
   '2005/5/17 MODIFY BY SONIA 外幣金額改抓A1K08
   '2007/11/16 modify by sonia 銷帳也不抓,同一請款單只抓收文日收文號最大者
   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(HC09,'Y','＊','') AS 本所案號,CPM03 AS 案件性質,A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11,2) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
   '         " FROM ACC1K0,fagent,nation,CaseProgress,HireCase,CasePropertyMap,SYSTEMKIND " & _
   '         " WHERE (A1K12=0 OR A1K12 IS NULL) and " & SQLNewFag("A1K03", "fa") & " " & _
   '         " and fa10=NA01(+) And A1K01=CP60 AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLcp
   'Modify by Morgan 2010/8/11 百年蟲 " & SqlDateT("A1K02") & "--> substrb(' '||sqldatet(A1K02),-9)
'   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(HC09,'Y','＊','') AS 本所案號,CPM03 AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11-(nvl(a1k06,0)*a1k10),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
'            " FROM fagent,nation,CaseProgress,HireCase,CasePropertyMap,SYSTEMKIND, " & _
'            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP FROM ACC1K0,CASEPROGRESS " & _
'            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10) NEW " & _
'            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
'            " and fa10=NA01(+) AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   '2005/5/17 END
   'Modify By Sindy 2012/10/11 +a1k31,改a1k08 - a1k31及a1k06,a1k10相關sql
   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(HC09,'Y','＊','') AS 本所案號,CPM03 AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 - nvl(a1k31,0) AS 外幣金額,ROUND(A1K11-nvl(a1k06,0),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
            " FROM fagent,nation,CaseProgress,HireCase,CasePropertyMap,SYSTEMKIND, " & _
            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP,a1k31 FROM ACC1K0,CASEPROGRESS " & _
            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,a1k31) NEW " & _
            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
            " and fa10=NA01(+) AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   'ServicePractice
   '2005/5/17 MODIFY BY SONIA 外幣金額改抓A1K08
   '2007/11/16 modify by sonia 銷帳也不抓,同一請款單只抓收文日收文號最大者
   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(SP15,'Y','＊','') AS 本所案號,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11,2) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
   '         " FROM ACC1K0,fagent,nation,CaseProgress,ServicePractice,CasePropertyMap,SYSTEMKIND " & _
   '         " WHERE (A1K12=0 OR A1K12 IS NULL) and " & SQLNewFag("A1K03", "fa") & " " & _
   '         " and fa10=NA01(+) And A1K01=CP60 AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLcp
   'Modify by Morgan 2010/8/11 百年蟲 " & SqlDateT("A1K02") & "--> substrb(' '||sqldatet(A1K02),-9)
'   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(SP15,'Y','＊','') AS 本所案號,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11-(nvl(a1k06,0)*a1k10),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
'            " FROM fagent,nation,CaseProgress,ServicePractice,CasePropertyMap,SYSTEMKIND, " & _
'            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP FROM ACC1K0,CASEPROGRESS " & _
'            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10) NEW " & _
'            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
'            " and fa10=NA01(+) AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   '2005/5/17 END
   'Modify By Sindy 2012/10/11 +a1k31,改a1k08 - a1k31及a1k06,a1k10相關sql
   strSql = strSql + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(SP15,'Y','＊','') AS 本所案號,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,A1K01 AS 單據編號,substrb(' '||sqldatet(A1K02),-9) AS 單據日期,A1K18 AS 幣別,A1K08 - nvl(a1k31,0) AS 外幣金額,ROUND(A1K11-nvl(a1k06,0),0) AS 台幣金額,A1K09 AS 規費,a1k29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,A1K06 " & _
            " FROM fagent,nation,CaseProgress,ServicePractice,CasePropertyMap,SYSTEMKIND, " & _
            "(SELECT A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,MAX(CP05||CP09) CP,a1k31 FROM ACC1K0,CASEPROGRESS " & _
            " WHERE (A1K12=0 OR A1K12 IS NULL) and A1K25 IS NULL " & strSQL1k0 & " and a1k01=cp60 " & strSalesArea & "GROUP BY A1K01,A1K02,A1K03,A1K06,A1K08,A1K09,A1K11,A1K13,A1K14,A1K15,A1K16,A1K18,A1K29,a1k10,a1k31) NEW " & _
            " where CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " " & _
            " and fa10=NA01(+) AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND CP01=CPM01(+) AND CP10=CPM02(+) AND A1K13=SK01(+) " & strSQLCP
   strSql = strSql & " ORDER BY 本所案號, 案件性質, 單據編號, 單據日期 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/28
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/9/28
      ShowNoData
      Me.Hide
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   Set GRD1.Recordset = adoRecordset
'2007/11/19 cancel by sonia
'   'Add By Cheng 2003/05/26 若同個請款單號有多筆收文資料, 則只保留一筆資料
'   If Me.Grd1.Rows > 1 Then
'      strA1k01 = ""
'      For ii = 1 To Me.Grd1.Rows - 1
'         If ii > Me.Grd1.Rows - 1 Then Exit For
'         If strA1k01 = Me.Grd1.TextMatrix(ii, 2) Then
'            Me.Grd1.RemoveItem ii
'            ii = ii - 1
'         Else
'            strA1k01 = Me.Grd1.TextMatrix(ii, 2)
'         End If
'      Next ii
'   End If
'2007/11/19 end
   Me.Enabled = False
   
   With GRD1
      
      GRD1.Visible = False
      For i = 1 To .Rows - 1
         .row = i
         '2007/11/20 add by sonia 抓翻譯費
         .col = 2
         adoRecordset1.CursorLocation = adUseClient
         adoRecordset1.Open "select ax206 from acc021, caseprogress where cp60='" & .Text & "' and cp10 in ('201','927') " & _
                            "and cp01||cp02||cp03||cp04=ax214(+) and '6130'=ax205(+)", adoTaie, adOpenStatic, adLockReadOnly
         .col = 10
         If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            If Not IsNull(adoRecordset1.Fields(0).Value) Then
               .Text = adoRecordset1.Fields(0).Value
            Else
               .Text = ""
            End If
         Else
            .Text = ""
         End If
         adoRecordset1.Close
         '2007/11/20 end
         .col = 5
         .CellAlignment = flexAlignRightCenter
         .col = 6
         .CellAlignment = flexAlignRightCenter
'         Int1 = Val(.Text)
         .col = 7
         .CellAlignment = flexAlignRightCenter
'         .col = 12
'         .CellAlignment = flexAlignRightCenter
'         Int2 = Val(.Text)
         .col = 8
         .CellAlignment = flexAlignRightCenter
         .col = 10
         .CellAlignment = flexAlignRightCenter
      Next i
      GRD1.Visible = True
   End With
   CheckOC
   Me.Enabled = True
End Sub

Sub StrMenu2()           '收款
Dim strTempName As String
Dim ii As Integer
Dim strA0Y01 As String   '收款單號
Dim straccSales As String  '2007/11/20 add by sonia
Dim str1P0Sales As String  '2007/11/20 add by sonia
Dim StrSQLa As String      '2007/11/26 add by sonia

   strSQL1k0 = "": strSQLCP = "": strSalesArea = ""
   strAccSystem = "": straccSales = "": str1P0Sales = ""
   '2007/11/16 modify by sonia
   'If Len(frm040205.txt1(0)) <> 0 Then
   If frm040205.txt1(0) <> "ALL" Then
   '2007/11/16 end
      strSQL1k0 = " and a1k13 in (" & GetAddStr(frm040205.txt1(0)) & ") "
      strAccSystem = " and substr(ax214, 1, Length(ax214) - 9) in (" & GetAddStr(frm040205.txt1(0)) & ") "
   End If
   If Trim(frm040205.txt1(0)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(0) & frm040205.txt1(0) 'Add By Sindy 2010/9/28
   End If
   pub_QL05 = pub_QL05 & ";" & frm040205.Label1(1) & "收款" 'Add By Sindy 2010/9/28
   If Len(Trim(frm040205.txt1(2))) <> 0 Or Len(Trim(frm040205.txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(2) & Trim(frm040205.txt1(2)) & "-" & Trim(frm040205.txt1(3)) 'Add By Sindy 2010/9/28
   End If
   '2007/11/26 add by sonia
   '國籍
   If Len(frm040205.txt1(4)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10>='" & frm040205.txt1(4) & "' "
   End If
   If Len(frm040205.txt1(5)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10<='" & frm040205.txt1(5) & "z' "
   End If
   '2007/11/26 end
   If Len(Trim(frm040205.txt1(4))) <> 0 Or Len(Trim(frm040205.txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(4) & Trim(frm040205.txt1(4)) & "-" & Trim(frm040205.txt1(5))  'Add By Sindy 2010/9/28
   End If
   If Len(frm040205.txt1(7)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10>='" & frm040205.txt1(7) & "' "
   End If
   If Len(frm040205.txt1(8)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10<='" & frm040205.txt1(8) & "' "
   End If
   If Len(Trim(frm040205.txt1(7))) <> 0 Or Len(Trim(frm040205.txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(7) & Trim(frm040205.txt1(7)) & "-" & Trim(frm040205.txt1(8))   'Add By Sindy 2010/9/28
   End If
   'Add by Morgan 2003/12/04
   '智權人員
   If Len(frm040205.txt1(9)) <> 0 Then
      strSQLCP = strSQLCP & " and CP13||''='" & frm040205.txt1(9) & "' "
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(8) & Trim(frm040205.txt1(9)) & frm040205.lbl1   'Add By Sindy 2010/9/28
   End If
   '業務區
   If Len(frm040205.txt1(10)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''>='" & frm040205.txt1(10) & "' "
   End If
   If Len(frm040205.txt1(11)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''<='" & frm040205.txt1(11) & "' "
   End If
   If Len(frm040205.txt1(10)) <> 0 Or Len(frm040205.txt1(11)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(9) & Trim(frm040205.txt1(10)) & "-" & Trim(frm040205.txt1(11))  'Add By Sindy 2010/9/28
   End If
   'End 2003/12/04
   '2007/11/20 add by sonia 非個人時再加印財務調整傳票
   If frm040205.txt1(9) = "" Then
      Select Case Mid(frm040205.txt1(10), 1, 2)
         Case "F3"
            straccSales = " and ax209='F4101' "
            str1P0Sales = " and a1p16='F4101' "
         Case "F2"
            'modify by sonia 2021/1/20 +F4104,F4105
            straccSales = " and ax209 in ('F4102','F4104','F4105') "
            str1P0Sales = " and a1p16 in ('F4102','F4104','F4105') "
         Case "F1"
            'modify by sonia 2021/1/20 +F4106,F4107
            straccSales = " and ax209 in ('F4103','F4106','F4107') "
            str1P0Sales = " and a1p16 in ('F4103','F4106','F4107') "
      End Select
      If Mid(frm040205.txt1(10), 1, 2) = "F1" And Mid(frm040205.txt1(11), 1, 2) = "F4" Then
            'modify by sonia 2021/1/20 +F4104~F4107
            straccSales = " and ax209 in ('F4101','F4102','F4103','F4104','F4105','F4106','F4107') "
            str1P0Sales = " and a1p16 in ('F4101','F4102','F4103','F4104','F4105','F4106','F4107') "
      End If
   Else
      straccSales = " and ax209='" & frm040205.txt1(9) & "' "
   End If
   '2007/11/20 end
   
'   StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
'   'Patent
'   '2007/11/20 modify by sonia同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
'   'strSQL = "select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(pa57,'Y','＊','') AS 本所案號,DECODE(pa09,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Y06 AS 外幣金額,A0Y04 AS 台幣金額,0 AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'', A1K01 " & _
'   '         " FROM ACC0Y0,ACC1K0,acc0z0,CaseProgress,patent,CasePropertyMap " & _
'   '         " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A1K01=CP60 And CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND CP01=CPM01(+) AND CP10=CPM02(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLCP
'   '2011/4/12 modify by sonia 同一案號不同請款單同時收款故要加a1k01
'   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
'   strSql = "SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(PA57,'Y','＊','') AS 本所案號,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Z04 AS 外幣金額,round(A0Z04*A0Y04,2)||'' AS 台幣金額,decode(round(A0Z04*A0Y04,2)-a1k30,0,A1K09,0) AS 規費,A1K29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,'' , cp60 " & _
'            " FROM fagent,nation,CaseProgress,Patent,CasePropertyMap,acc1p0, acc021,SYSTEMKIND, " & _
'            "(SELECT A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,MAX(CP05||CP09) CP,a1k01 FROM ACC0Y0,ACC1K0,acc0z0,CASEPROGRESS " & _
'            " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A0Z02=CP60(+)" & strSalesArea & strSQL1k0 & "GROUP BY A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,a1k01) NEW " & _
'            " WHERE cp09 in substr(new.cp,9,9) and " & SQLNewFag("A0Y07", "fa") & " and fa10=NA01(+) and cp01=sk01(+) And CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQLCP & "and new.a0y01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+)" & str1P0Sales & _
'            "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'            "union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(PA57,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,PA01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,patent where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & "and a0202=a1p22(+) and 'F'=a1p02 and a1p04 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=PA01 AND substr(ax214, Length(ax214) - 8, 6)=PA02 AND substr(ax214, Length(ax214) - 2, 1)=PA03 AND substr(ax214, Length(ax214) - 1, 2)=PA04 union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(PA57,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,PA01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress,patent where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
'            "and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=PA01 AND substr(ax214, Length(ax214) - 8, 6)=PA02 AND substr(ax214, Length(ax214) - 2, 1)=PA03 AND substr(ax214, Length(ax214) - 1, 2)=PA04"
'   'TradeMark
'   '2007/11/20 modify by sonia同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
'   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(TM29,'Y','＊','') AS 本所案號,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Y06 AS 外幣金額,A0Y04 AS 台幣金額,0 AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'', A1K01 " & _
'   '         " FROM ACC0Y0,ACC1K0,acc0z0,CaseProgress,TradeMark,CasePropertyMap " & _
'   '         " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A1K01=CP60 And CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND CP01=CPM01(+) AND CP10=CPM02(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLCP
'   '2011/4/12 modify by sonia 同一案號不同請款單同時收款故要加a1k01
'   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
'   strSql = strSql + " union SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(TM29,'Y','＊','') AS 本所案號,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Z04 AS 外幣金額,round(A0Z04*A0Y04,2)||'' AS 台幣金額,decode(round(A0Z04*A0Y04,2)-a1k30,0,A1K09,0) AS 規費,A1K29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,'' , cp60 " & _
'            " FROM fagent,nation,CaseProgress,TRADEMARK,CasePropertyMap,acc1p0, acc021,SYSTEMKIND, " & _
'            "(SELECT A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,MAX(CP05||CP09) CP,a1k01 FROM ACC0Y0,ACC1K0,acc0z0,CASEPROGRESS " & _
'            " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A0Z02=CP60(+)" & strSalesArea & strSQL1k0 & "GROUP BY A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,a1k01) NEW " & _
'            " WHERE cp09 in substr(new.cp,9,9) and " & SQLNewFag("A0Y07", "fa") & " and fa10=NA01(+) and cp01=sk01(+) And CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQLCP & "and new.a0y01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+)" & str1P0Sales & _
'            "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'            "union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(TM29,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,TM01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,trademark where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & "and a0202=a1p22(+) and 'F'=a1p02 and a1p04 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=TM01 AND substr(ax214, Length(ax214) - 8, 6)=TM02 AND substr(ax214, Length(ax214) - 2, 1)=TM03 AND substr(ax214, Length(ax214) - 1, 2)=TM04 union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(TM29,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,TM01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress,TRADEMARK where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
'            "and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=TM01 AND substr(ax214, Length(ax214) - 8, 6)=TM02 AND substr(ax214, Length(ax214) - 2, 1)=TM03 AND substr(ax214, Length(ax214) - 1, 2)=TM04"
'   'LawCase
'   '2007/11/20 modify by sonia同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
'   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(LC08,'Y','＊','') AS 本所案號,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Y06 AS 外幣金額,A0Y04 AS 台幣金額,0 AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'', A1K01 " & _
'   '         " FROM ACC0Y0,ACC1K0,acc0z0,CaseProgress,LawCase,CasePropertyMap " & _
'   '         " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A1K01=CP60 And CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP01=CPM01(+) AND CP10=CPM02(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLCP
'   '2011/4/12 modify by sonia 同一案號不同請款單同時收款故要加a1k01
'   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
'   strSql = strSql + " union SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(LC08,'Y','＊','') AS 本所案號,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Z04 AS 外幣金額,round(A0Z04*A0Y04,2)||'' AS 台幣金額,decode(round(A0Z04*A0Y04,2)-a1k30,0,A1K09,0) AS 規費,A1K29 AS 結清," & StrSQLa & "'' AS 翻譯費費,A1K13,'' , cp60 " & _
'            " FROM fagent,nation,CaseProgress,LawCase,CasePropertyMap,acc1p0, acc021,SYSTEMKIND, " & _
'            "(SELECT A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,MAX(CP05||CP09) CP,a1k01 FROM ACC0Y0,ACC1K0,acc0z0,CASEPROGRESS " & _
'            " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A0Z02=CP60(+)" & strSalesArea & strSQL1k0 & "GROUP BY A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,a1k01) NEW " & _
'            " WHERE cp09 in substr(new.cp,9,9) and " & SQLNewFag("A0Y07", "fa") & " and fa10=NA01(+) and cp01=sk01(+) And CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQLCP & "and new.a0y01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+)" & str1P0Sales & _
'            "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'            "union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(LC08,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,LC01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,LawCase where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & "and a0202=a1p22(+) and 'F'=a1p02 and a1p04 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=LC01 AND substr(ax214, Length(ax214) - 8, 6)=LC02 AND substr(ax214, Length(ax214) - 2, 1)=LC03 AND substr(ax214, Length(ax214) - 1, 2)=LC04 union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(LC08,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,LC01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress,LawCase where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
'            "and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=LC01 AND substr(ax214, Length(ax214) - 8, 6)=LC02 AND substr(ax214, Length(ax214) - 2, 1)=LC03 AND substr(ax214, Length(ax214) - 1, 2)=LC04"
'   'HireCase
'   '2007/11/20 modify by sonia同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
'   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(HC09,'Y','＊','') AS 本所案號,CPM03 AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Y06 AS 外幣金額,A0Y04 AS 台幣金額,0 AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'', A1K01 " & _
'   '         " FROM ACC0Y0,ACC1K0,acc0z0,CaseProgress,HireCase,CasePropertyMap " & _
'   '         " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A1K01=CP60 And CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP01=CPM01(+) AND CP10=CPM02(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLCP
'   '2011/4/12 modify by sonia 同一案號不同請款單同時收款故要加a1k01
'   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
'   strSql = strSql + " union SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(HC09,'Y','＊','') AS 本所案號,CPM03 AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Z04 AS 外幣金額,round(A0Z04*A0Y04,2)||'' AS 台幣金額,decode(round(A0Z04*A0Y04,2)-a1k30,0,A1K09,0) AS 規費,A1K29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,'' , cp60 " & _
'            " FROM fagent,nation,CaseProgress,HireCase,CasePropertyMap,acc1p0, acc021,SYSTEMKIND, " & _
'            "(SELECT A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,MAX(CP05||CP09) CP,a1k01 FROM ACC0Y0,ACC1K0,acc0z0,CASEPROGRESS " & _
'            " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A0Z02=CP60(+)" & strSalesArea & strSQL1k0 & "GROUP BY A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,a1k01) NEW " & _
'            " WHERE cp09 in substr(new.cp,9,9) and " & SQLNewFag("A0Y07", "fa") & " and fa10=NA01(+) and cp01=sk01(+) And CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQLCP & "and new.a0y01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+)" & str1P0Sales & _
'            "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'            "union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(HC09,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,HC01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,HireCase where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & "and a0202=a1p22(+) and 'F'=a1p02 and a1p04 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=HC01 AND substr(ax214, Length(ax214) - 8, 6)=HC02 AND substr(ax214, Length(ax214) - 2, 1)=HC03 AND substr(ax214, Length(ax214) - 1, 2)=HC04 union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(HC09,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,HC01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress,HireCase where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
'            "and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=HC01 AND substr(ax214, Length(ax214) - 8, 6)=HC02 AND substr(ax214, Length(ax214) - 2, 1)=HC03 AND substr(ax214, Length(ax214) - 1, 2)=HC04"
'   'ServicePractice
'   '2007/11/20 modify by sonia同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
'   'strSQL = strSQL + " union select A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(SP15,'Y','＊','') AS 本所案號,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Y06 AS 外幣金額,A0Y04 AS 台幣金額,0 AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'', A1K01 " & _
'   '         " FROM ACC0Y0,ACC1K0,acc0z0,CaseProgress,ServicePractice,CasePropertyMap " & _
'   '         " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A1K01=CP60 And CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND CP01=CPM01(+) AND CP10=CPM02(+) And A1K13=CP01 AND A1K14=CP02 AND A1K15=CP03 AND A1K16=CP04 " & strSQLCP
'   '2011/4/12 modify by sonia 同一案號不同請款單同時收款故要加a1k01
'   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
'   strSql = strSql + " union SELECT A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16||DECODE(SP15,'Y','＊','') AS 本所案號,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Z04 AS 外幣金額,round(A0Z04*A0Y04,2)||'' AS 台幣金額,decode(round(A0Z04*A0Y04,2)-a1k30,0,A1K09,0) AS 規費,A1K29 AS 結清," & StrSQLa & "'' AS 翻譯費,A1K13,'' , cp60 " & _
'            " FROM fagent,nation,CaseProgress,ServicePractice,CasePropertyMap,acc1p0, acc021,SYSTEMKIND, " & _
'            "(SELECT A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,MAX(CP05||CP09) CP,a1k01 FROM ACC0Y0,ACC1K0,acc0z0,CASEPROGRESS " & _
'            " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A0Z02=CP60(+)" & strSalesArea & strSQL1k0 & "GROUP BY A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,a1k01) NEW " & _
'            " WHERE cp09 in substr(new.cp,9,9) and " & SQLNewFag("A0Y07", "fa") & " and fa10=NA01(+) and cp01=sk01(+) And CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQLCP & "and new.a0y01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+)" & str1P0Sales & _
'            "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'            "union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(SP15,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,SP01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,ServicePractice where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & "and a0202=a1p22(+) and 'F'=a1p02 and a1p04 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=SP01 AND substr(ax214, Length(ax214) - 8, 6)=SP02 AND substr(ax214, Length(ax214) - 2, 1)=SP03 AND substr(ax214, Length(ax214) - 1, 2)=SP04 union " & _
'            "select substr(ax214, 1, Length(ax214) - 9)||'-'||substr(ax214, Length(ax214) - 8, 6)||'-'||substr(ax214, Length(ax214) - 2, 1)||'-'||substr(ax214, Length(ax214) - 1, 2)||DECODE(SP15,'Y','＊','') AS 本所案號,'' AS 案件性質,'' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,(ax207-ax206)||'' AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,SP01,'' ,'' " & _
'            "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress,ServicePractice where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
'            "and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
'            "AND substr(ax214, 1, Length(ax214) - 9)=SP01 AND substr(ax214, Length(ax214) - 8, 6)=SP02 AND substr(ax214, Length(ax214) - 2, 1)=SP03 AND substr(ax214, Length(ax214) - 1, 2)=SP04 "
'   strSql = strSql + " ORDER BY 本所案號, 案件性質, 單據編號, 單據日期 "
   'Modify By Sindy 2018/7/20 抓資料改到 accrpt0205
   strSql = "select R21214||R21203 AS 本所案號, R21216 AS 案件性質,R21217 AS 單據編號,SqlDateT(R21204) AS 單據日期,R21218 AS 幣別,ROUND(R21207,2) AS 外幣金額,ROUND(R21208,2) AS 台幣金額,R21206 AS 規費,R21219 AS 結清,R21210 AS 翻譯費,'',''" & _
            " from accrpt0205 where R21201='" & strUserNum & "'"
   strSql = strSql + " ORDER BY 本所案號, 案件性質, 單據編號, 單據日期 "
   CheckOC
   '2007/11/27 cancel by sonia
   'If Len(frm040205.txt1(0)) <> 0 Then
   '   strTemp = Split(frm040205.txt1(0), ",")
   'End If
   '2007/11/27 end
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/28
'2007/11/20 CANCEL BY SONIA
'      adoRecordset.MoveFirst
'      Do While adoRecordset.EOF = False
'      s = 0
'      If Len(frm040205.txt1(0)) <> 0 Then
'         If Not IsNull(adoRecordset.Fields(11)) Then
'            For i = 0 To UBound(strTemp)
'               If strTemp(i) = adoRecordset.Fields(11) Then
'                  s = 1
'               End If
'            Next i
'         End If
'      End If
'      If Len(frm040205.txt1(4)) <> 0 And s = 1 Then
'         If "" & GetPrjNationNumber(adoRecordset.Fields(9)) >= frm040205.txt1(4) Then
'            s = 1
'         Else
'            s = 0
'         End If
'      End If
'      If Len(frm040205.txt1(5)) <> 0 And s = 1 Then
'         If "" & GetPrjNationNumber(adoRecordset.Fields(9)) <= frm040205.txt1(5) & "z" Then
'            s = 1
'         Else
'            s = 0
'         End If
'      End If
'         If s = 0 Then
'            adoRecordset.Delete
'         End If
'         adoRecordset.MoveNext
'      Loop
'2007/11/20 END
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/9/28
      ShowNoData
      Me.Hide
      Exit Sub
   End If

   Set GRD1.Recordset = adoRecordset

'2007/11/28 add by sonia 抓翻譯費
   Me.Enabled = False
   
   With GRD1
      
      GRD1.Visible = False
      For i = 1 To .Rows - 1
         .row = i
'         .col = 13
'         adoRecordset1.CursorLocation = adUseClient
'         adoRecordset1.Open "select ax206 from acc021, caseprogress where cp60='" & .Text & "' and cp10 in ('201','927') " & _
'                            "and cp01||cp02||cp03||cp04=ax214(+) and '6130'=ax205(+)", adoTaie, adOpenStatic, adLockReadOnly
'         .col = 10
'         If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'            If Not IsNull(adoRecordset1.Fields(0).Value) Then
'               .Text = adoRecordset1.Fields(0).Value
'            Else
'               .Text = ""
'            End If
'         Else
'            .Text = ""
'         End If
'         adoRecordset1.Close
         .col = 5
         .CellAlignment = flexAlignRightCenter
         .col = 6
         .CellAlignment = flexAlignRightCenter
'         Int1 = Val(.Text)
         .col = 7
         .CellAlignment = flexAlignRightCenter
'         .col = 12
'         .CellAlignment = flexAlignRightCenter
'         Int2 = Val(.Text)
         .col = 8
         .CellAlignment = flexAlignRightCenter
         .col = 10
         .CellAlignment = flexAlignRightCenter
      Next i
      GRD1.Visible = True
   End With
   CheckOC
   Me.Enabled = True
'2007/11/20 end

'2007/11/20 cancel by sonia
'Dim IntTestTemp As Double    '台幣金額
'Dim IntTestTemp2 As Double   '外幣金額
'
'   Grd1.Visible = False
'   For i = 1 To Grd1.Rows - 1
'      DoEvents
'      Grd1.row = i
'      Grd1.col = 2
'      strSQL = Grd1.Text
'      strSQL = "SELECT A0Z04,A0Y04 FROM ACC0Z0,ACC0Y0, ACC1K0, (Select CP01, CP02, CP03, CP04, CP60 From CaseProgress Where " & ChgCaseprogress(Replace(Me.Grd1.TextMatrix(i, 0), "-", "")) & " And CP60='" & Me.Grd1.TextMatrix(i, 13) & "' Group By CP01, CP02, CP03, CP04, CP60 ) C1 WHERE A0Z01='" & Grd1.Text & "' AND A0Y01=A0Z01 And A0Z02=A1K01 And A1K01=C1.CP60 And A1K13=C1.CP01 And A1K14=C1.CP02 And A1K15=C1.CP03 And A1K16=C1.CP04 And A1K01='" & Me.Grd1.TextMatrix(i, 13) & "' And C1.CP01||C1.CP02||C1.CP03||C1.CP04='" & Replace(Me.Grd1.TextMatrix(i, 0), "-", "") & "' "
'      CheckOC3
'      AdoRecordSet3.CursorLocation = adUseClient
'      AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
'         IntTestTemp = 0
'         IntTestTemp2 = 0
'         Do While AdoRecordSet3.EOF = False
'            'Modify By Cheng 2002/02/15 若值為Null則設成0
'   '         IntTestTemp = IntTestTemp + (AdoRecordSet3.Fields(0) * AdoRecordSet3.Fields(1))
'            IntTestTemp = IntTestTemp + (AdoRecordSet3.Fields(0) * IIf(IsNull(AdoRecordSet3.Fields(1)), 0, AdoRecordSet3.Fields(1)))
'            IntTestTemp2 = IntTestTemp2 + Val(CheckStr(AdoRecordSet3.Fields(0)))
'            AdoRecordSet3.MoveNext
'         Loop
'         Grd1.col = 5
'         Grd1.Text = Format(str(IntTestTemp2), "0.00")
'         Grd1.CellAlignment = flexAlignRightCenter
'         Grd1.col = 6
'         Grd1.Text = Format(str(IntTestTemp), "0.00")
'         Grd1.CellAlignment = flexAlignRightCenter
'      Else
'         Grd1.col = 5
'         Grd1.Text = "0.00"
'         Grd1.CellAlignment = flexAlignRightCenter
'         Grd1.col = 6
'         Grd1.Text = "0.00"
'         Grd1.CellAlignment = flexAlignRightCenter
'      End If
'      CheckOC3
'      Grd1.col = 9
'      If PUB_GetAgentName(Me.Grd1.TextMatrix(i, 11), Me.Grd1.Text, strTempName) Then
'         Grd1.Text = strTempName
'      End If
'   Next i
'   Grd1.Visible = True
'   CheckOC
'2007/11/20 end
End Sub
Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2018/7/23
   If Me.Tag = "" Then
      bolToEndByNick = True
   End If
   '2018/7/23 END
   Set frm040205a = Nothing
End Sub

Private Sub Grd1_Click()
Dim j As Integer
Dim i As Integer
Dim tmpcur As Integer

   tmpcur = GRD1.MouseRow
   GRD1.Visible = False
   If TmpRow <> 0 Then
      GRD1.row = TmpRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   GRD1.row = tmpcur
   GRD1.col = 0
   If GRD1.row <> 0 Then
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   TmpRow = tmpcur
   GRD1.Visible = True
End Sub
