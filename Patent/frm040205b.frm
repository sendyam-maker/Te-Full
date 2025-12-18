VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040205b 
   BorderStyle     =   1  '單線固定
   Caption         =   "FC收款請款點數查詢"
   ClientHeight    =   3405
   ClientLeft      =   5235
   ClientTop       =   4320
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5910
   Begin VB.CommandButton Cmdok 
      Caption         =   "結束(&X)"
      Height          =   350
      Index           =   1
      Left            =   4980
      TabIndex        =   11
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   4155
      TabIndex        =   10
      Top             =   45
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   4635
      Left            =   0
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8176
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd2 
      Height          =   1845
      Left            =   90
      TabIndex        =   14
      Top             =   570
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   3254
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
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
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label1 
      Caption         =   "PS : 請款點數含分配給其他部門的點數"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   90
      TabIndex        =   13
      Top             =   3150
      Width           =   3720
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   4
      Left            =   1455
      TabIndex        =   9
      Top             =   2730
      Width           =   2460
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   3
      Left            =   1455
      TabIndex        =   8
      Top             =   2490
      Width           =   2460
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   2
      Left            =   8325
      TabIndex        =   7
      Top             =   4020
      Width           =   2460
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   8325
      TabIndex        =   6
      Top             =   3780
      Width           =   2460
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   8325
      TabIndex        =   5
      Top             =   3540
      Width           =   2460
   End
   Begin VB.Label Label1 
      Caption         =   "收款點數合計："
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   2715
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "台幣收款合計："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   2490
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "請款點數合計："
      Height          =   180
      Index           =   2
      Left            =   6960
      TabIndex        =   2
      Top             =   4020
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "台幣請款合計："
      Height          =   180
      Index           =   1
      Left            =   6960
      TabIndex        =   1
      Top             =   3780
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "請款合計："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   1305
   End
End
Attribute VB_Name = "frm040205b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; Grd2改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2007/11/16 整理 by sonia
Option Explicit

Dim s As Integer, strSql As String, strTemp As Variant, StrTest As String, i As Variant, j As Integer
Dim IntTo1 As Double, IntTo2 As Double, IntTo3 As Double, IntTo5 As Double, strSQL1k0 As String, strSQLCP As String, Int1 As Long, Int2 As Long, IntTo4 As Double


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Me.Tag = "0" 'Add By Sindy 2018/7/23
         Me.Hide
      Case 1
         Me.Tag = "1" 'Add By Sindy 2018/7/23
         bolToEndByNick = True
         Unload Me
         Exit Sub
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2018/7/23
   If Me.Tag = "" Then
      bolToEndByNick = True
   End If
   '2018/7/23 END
   Set frm040205b = Nothing
End Sub

Sub StrMenu()
   Me.Enabled = False
   'Modify By Sindy 2017/4/17
   GRD1.Clear
   Select Case frm040205.txt1(1)
      Case "1"
         Label1(0).Caption = "請款合計："
         Label1(3).Caption = "台幣請款合計："
         Label1(4).Caption = "請款點數合計："
         StrMenu1            '請款
         StrTotle2
      Case "2"
         Label1(0).Caption = "收款合計："
         Label1(3).Caption = "台幣收款合計："
         Label1(4).Caption = "收款點數合計："
         StrMenu2            '收款
      Case Else
   End Select
   '2017/4/17 END
   
'   '92.03.05 nick   寫在另外一個 function
'   '92/03/05 nick  邱小姐說請款數據應該和 財務系統 相同，收款不管
'   StrMenu1            '請款
'   'StrTotle 'Modify By Sindy 2012/10/11 Mark
'   GRD1.Clear
'   StrMenu2            '收款
   
'   StrTotle2
'   lbl1(0).Caption = Format(str(IntTo1), "###,###,###.00")
'   lbl1(1).Caption = Format(str(IntTo2), "###,###,###")
'   'lbl1(2).Caption = Format(str(IntTo4), "###,###,###.00")
'   lbl1(2).Caption = Format(str(IntTo4), "###,###,###.0")
   lbl1(3).Caption = Format(str(IntTo3), "###,###,##0") '###,###,###.00
   '2007/11/27 modify by sonia
   'lbl1(4).Caption = Format(str(IntTo3 / 1000), "###,###,###.00000")
   lbl1(4).Caption = Format(str(IntTo5), "###,###,###.000") '###,###,###.00000
   '2007/11/27 end
   Me.Enabled = True
End Sub

Sub StrTotle2()
Dim tmpIntTo2 As Double  '2007/11/27 add by sonia
   
'   IntTo3 = 0: IntTo5 = 0
'   With GRD1
'      For i = 1 To .Rows - 1
'         .row = i
'         .col = 4
'         IntTo3 = IntTo3 + Val(.Text)
'         '2007/11/27 add by sonia
'         tmpIntTo2 = Val(.Text)
'         .col = 5 '規費
'         IntTo5 = IntTo5 + Format(str((tmpIntTo2 - Val(.Text)) / 1000), "###,###,###.00000")
'         '2007/11/27 end
'      Next i
'   End With
   
   'Modify By Sindy 2017/4/7
   IntTo3 = 0: IntTo5 = 0
   With grd2
      For i = 1 To .Rows - 1
         .row = i
         .col = 2 '台幣
         If .Text <> "" Then
            IntTo3 = IntTo3 + Val(CDbl(.Text))
         End If
         .col = 4 '點數
         If .Text <> "" Then
            IntTo5 = IntTo5 + Val(CDbl(.Text))
         End If
      Next i
   End With
   '2017/4/7 END
End Sub

'Sub StrTotle()
'Dim tmpIntTo2 As Double
'
'   IntTo1 = 0
'   IntTo2 = 0
'   IntTo4 = 0
'   With Grd1
'      For i = 1 To .Rows - 1
'         .row = i
'         .col = 3 '外幣金額
'         IntTo1 = IntTo1 + Val(.Text)
'         .col = 4 '台幣金額
'         IntTo2 = IntTo2 + Val(.Text)
'         tmpIntTo2 = Val(.Text)
'         .col = 5 '規費
'         'Modify By Cheng 2003/12/16
'         '2007/11/19 modify by sonia
'         'IntTo4 = IntTo4 + Val(.Text)
'         IntTo4 = IntTo4 + Format(str((tmpIntTo2 - Val(.Text)) / 1000), "###,###,###.0")
'      Next i
'   End With
'   'IntTo4 = Format(str((IntTo2 - IntTo4) / 1000), "###,###,###.0")  2007/11/19 cancel by sonia
'End Sub

Sub StrMenu1()            '請款
   strSQL1k0 = "": strSQLCP = ""
   '2007/11/19 modify by sonia
   'If Len(frm040205.txt1(0)) <> 0 Then
   If frm040205.txt1(0) <> "ALL" Then
   '2007/11/19 end
      strSQL1k0 = strSQL1k0 & " and a1k13 in (" & GetAddStr(frm040205.txt1(0)) & ") "
   End If
   If Trim(frm040205.txt1(0)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(0) & frm040205.txt1(0) 'Add By Sindy 2010/9/28
   End If
   'pub_QL05 = pub_QL05 & ";" & frm040205.Label1(1) & "請款&收款" 'Add By Sindy 2010/9/28
   If Len(Trim(frm040205.txt1(2))) <> 0 Then
      strSQL1k0 = strSQL1k0 & " and A1K02>=" & Val(frm040205.txt1(2)) & " "
   End If
   If Len(Trim(frm040205.txt1(3))) <> 0 Then
      strSQL1k0 = strSQL1k0 & " AND A1K02<=" & Val(frm040205.txt1(3)) & " "
   End If
   If Len(Trim(frm040205.txt1(2))) <> 0 Or Len(Trim(frm040205.txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(2) & Trim(frm040205.txt1(2)) & "-" & Trim(frm040205.txt1(3)) 'Add By Sindy 2010/9/28
   End If
   If Len(frm040205.txt1(4)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10>='" & frm040205.txt1(4) & "' "
   End If
   If Len(frm040205.txt1(5)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10<='" & frm040205.txt1(5) & "z' "
   End If
   If Len(Trim(frm040205.txt1(4))) <> 0 Or Len(Trim(frm040205.txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(4) & Trim(frm040205.txt1(4)) & "-" & Trim(frm040205.txt1(5))  'Add By Sindy 2010/9/28
   End If
   'Add By cheng 2002/02/15
   '案件性質
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
      strSQLCP = strSQLCP & " and CP12||''>='" & frm040205.txt1(10) & "' "
   End If
   If Len(frm040205.txt1(11)) <> 0 Then
      strSQLCP = strSQLCP & " and CP12||''<='" & frm040205.txt1(11) & "' "
   End If
   If Len(frm040205.txt1(10)) <> 0 Or Len(frm040205.txt1(11)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(9) & Trim(frm040205.txt1(10)) & "-" & Trim(frm040205.txt1(11))  'Add By Sindy 2010/9/28
   End If
   'End 2003/12/04
   'Modify By Cheng 2002/02/15加入案件性質的條件
   '92.4.2 不抓CASEPROGRESS否則資料會重覆
   '2005/5/17 MODIFY BY SONIA 外幣金額改抓A1K08
   '2007/11/19 modify by sonia 銷帳也不抓,同一請款單只抓收文日收文號最大者
   'strSQL = "SELECT A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11,2) AS 台幣金額,A1K09 AS 規費,'' AS 結清,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),A1K03) AS 代理人,'' AS 翻譯費,A1K13,A1K06, NULL " & _
   '         " FROM ACC1K0,fagent,nation, Caseprogress " & _
   '         " WHERE (A1K12=0 OR A1K12 IS NULL) and " & SQLNewFag("A1K03", "fa") & " " & _
   '         " and fa10=na01(+) And a1k01=CP60 And a1k13=CP01 And a1k14=CP02 And a1k15=CP03 And a1k16=CP04 " & strSQL1 & _
   '         " Group By A1K01, A1K02 ,A1K18 , A1K08 , ROUND(A1K11,2) ,A1K09 ,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),A1K03) ,A1K13,A1K06 "
   'Modify By Sindy 2012/10/11 +a1k31,改a1k08 - a1k31及a1k06,a1k10相關sql
'   strSql = "SELECT A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K08 AS 外幣金額,ROUND(A1K11-(nvl(a1k06,0)*a1k10),0) AS 台幣金額,A1K09 AS 規費,'' AS 結清,nvl(NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),A1K03) AS 代理人,'' AS 翻譯費,A1K13,A1K06, NULL FROM fagent,nation, Caseprogress, " & _
'            " (SELECT A1K01,A1K02,A1K03,A1K08,A1K11,A1K09,A1K06,a1k13,a1k18,a1k10,MAX(CP05||CP09) CP FROM ACC1K0,CASEPROGRESS Where (A1K12 = 0 Or A1K12 Is Null) " & strSQL1k0 & " And A1K25 Is Null And A1K01 = CP60 GROUP BY A1K01,A1K02,A1K03,A1K08,A1K11,A1K09,A1K06,a1k13,a1k18,a1k10) NEW " & _
'            " WHERE CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " and fa10=na01(+) " & strSQLCP & _
'            " Group By A1K01, A1K02 ,A1K18 , A1K08 , ROUND(A1K11-(nvl(a1k06,0)*a1k10),0) ,A1K09 ,nvl(NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),A1K03) ,A1K13,A1K06 "
   '2007/11/19 end
   SetGrd2 'Add By Sindy 2012/10/11
   strSql = "SELECT A1K18 AS 幣別,sum(A1K08 - nvl(a1k31,0)) AS 外幣金額,sum(ROUND(A1K11-nvl(a1k06,0),0)) AS 台幣金額,sum(A1K09) AS 規費,' ' AS 請款點數 FROM fagent,nation, Caseprogress, " & _
            " (SELECT A1K01,A1K02,A1K03,A1K08,A1K11,A1K09,A1K06,a1k13,a1k18,a1k10,MAX(CP05||CP09) CP,a1k31 FROM ACC1K0,CASEPROGRESS Where (A1K12 = 0 Or A1K12 Is Null) " & strSQL1k0 & " And A1K25 Is Null And A1K01 = CP60 GROUP BY A1K01,A1K02,A1K03,A1K08,A1K11,A1K09,A1K06,a1k13,a1k18,a1k10,a1k31) NEW " & _
            " WHERE CP09=SUBSTR(NEW.CP,9,9) AND " & SQLNewFag("A1K03", "fa") & " and fa10=na01(+) " & strSQLCP & _
            " Group By A1K18 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/28
   Else
      'ShowNoData
      'Me.Hide
      InsertQueryLog (0) 'Add By Sindy 2010/9/28
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   'strSQL = "SELECT A1K01 AS 單據編號," & SqlDateT("A1K02") & " AS 單據日期,A1K18 AS 幣別,A1K11/A1K10 AS 外幣金額,A1K11 AS 台幣金額,A1K09 AS 規費,'' AS 結清,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),A1K03) AS 代理人,A1L05 AS 翻譯費,A1K13,A1K06 FROM ACC1K0,acc1l0,fagent,nation WHERE a1l04='80' and a1k01=a1l01(+)  AND (A1K12=0 OR A1K12 IS NULL) and " & SQLNewFag("A1K03", "fa") & " and fa10=na01(+) " & strSQL1
   'Modify By Sindy 2012/10/11 Mark and 修改如下
'   Set Grd1.Recordset = adoRecordset
'   Me.Enabled = False
'   '抓結清
'   With Grd1
'      'Grd1.MousePointer = flexArrowHourGlass
'      Grd1.Visible = False
'      For i = 1 To .Rows - 1
'         .row = i
'         .col = 3 '外幣金額
'         .CellAlignment = flexAlignRightCenter
'         .col = 4 '台幣金額
'         .CellAlignment = flexAlignRightCenter
'         '台幣金額
'         Int1 = Val(.Text)
'         .col = 5 '規費
'         .CellAlignment = flexAlignRightCenter
'         .col = 10 '折讓金額
'         .CellAlignment = flexAlignRightCenter
'         Int2 = Val(.Text)
'         .col = 0
'         If GetPrjTotleByEnd(.Text) >= Int1 - Int2 Then
'            .col = 6 '結清
'            .CellAlignment = flexAlignRightCenter
'            .Text = "Y"
'         Else
'            .col = 6 '結清
'            .CellAlignment = flexAlignRightCenter
'             .Text = ""
'         End If
'         .col = 0
'         CheckOC
'         'DoEvents
''2007/11/19 cancel by sonia 此畫面無翻譯費
''         strSQL = "SELECT A1L05 FROM ACC1L0 WHERE a1l04='80' and a1l01='" & Trim(.Text) & "' "
''         adoRecordset.CursorLocation = adUseClient
''         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
''         If adoRecordset.RecordCount <> 0 Then
''            .col = 8 '翻譯費
''            .CellAlignment = flexAlignRightCenter
''            .Text = CheckStr(adoRecordset.Fields(0))
''         Else
''            .col = 8 '翻譯費
''            .CellAlignment = flexAlignRightCenter
''            .Text = ""
''         End If
''2007/11/19 end
'      Next i
'      'Grd1.MousePointer = flexDefault
'      Grd1.Visible = True
'   End With
   Set grd2.Recordset = adoRecordset
   SetGrd2
   Me.Enabled = False
   Dim dblNT As Double, dblFee As Double
   With grd2
      grd2.Visible = False
      For i = 1 To .Rows - 1
         dblNT = 0: dblFee = 0
         .row = i
         .col = 1 '外幣金額
         .Text = Format(.Text, "###,###,###.00")
         .col = 2 '台幣金額
         dblNT = .Text
         .Text = Format(.Text, "###,###,##0")
         .col = 3 '規費
         dblFee = .Text
         .col = 4 '請款點數
         .Text = Format(str((dblNT - dblFee) / 1000), "###,###,###.000")
      Next i
      grd2.Visible = True
   End With
   '2012/10/11 End
   CheckOC
   Me.Enabled = True
End Sub

'Add By Sindy 2012/10/11
Private Sub SetGrd2()
   With grd2
      .Visible = False
      .Cols = 5
      .row = 0
      .col = 0: .ColWidth(0) = 800: .Text = "幣別"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      
      .col = 1: .ColWidth(1) = 1500: .Text = "外幣金額"
      .CellAlignment = flexAlignRightCenter
      .ColAlignment(1) = flexAlignRightCenter
      
      .col = 2: .ColWidth(2) = 1500: .Text = "台幣金額"
      .CellAlignment = flexAlignRightCenter
      .ColAlignment(2) = flexAlignRightCenter
      
      .col = 3: .ColWidth(3) = 0: .Text = "規費"
      .CellAlignment = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      
      .col = 4: .ColWidth(4) = 1500
      'Modify By Sindy 2017/4/17
      Select Case frm040205.txt1(1)
      Case "1"
         .Text = "請款點數"
      Case "2"
         .Text = "收款點數"
      End Select
      '2017/4/17 END
      .CellAlignment = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      
      .Visible = True
   End With
End Sub

Sub StrMenu2()           '收款
Dim straccSales As String  '2007/11/27 add by sonia
Dim str1P0Sales As String  '2007/11/27 add by sonia
Dim strSalesArea As String '2007/11/27 add by sonia
Dim strAccSystem As String '2007/11/27 add by sonia

   strSQL1k0 = "": strSQLCP = "": strSalesArea = ""
   strAccSystem = "": straccSales = "": str1P0Sales = ""
   '2007/11/26 add by sonia
   If frm040205.txt1(0) <> "ALL" Then
      strSQL1k0 = " and a1k13 in (" & GetAddStr(frm040205.txt1(0)) & ") "
      strAccSystem = " and substr(ax214, 1, Length(ax214) - 9) in (" & GetAddStr(frm040205.txt1(0)) & ") "
   End If
   '國籍
   If Len(frm040205.txt1(4)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10>='" & frm040205.txt1(4) & "' "
   End If
   If Len(frm040205.txt1(5)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10<='" & frm040205.txt1(5) & "z' "
   End If
   '2007/11/26 end
   If Len(frm040205.txt1(7)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10>='" & frm040205.txt1(7) & "' "
   End If
   If Len(frm040205.txt1(8)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10<='" & frm040205.txt1(8) & "' "
   End If
   'Add by Morgan 2003/12/04
   '智權人員
   If Len(frm040205.txt1(9)) <> 0 Then
      strSQLCP = strSQLCP & " and CP13||''='" & frm040205.txt1(9) & "' "
   End If
   '業務區
   If Len(frm040205.txt1(10)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''>='" & frm040205.txt1(10) & "' "
   End If
   If Len(frm040205.txt1(11)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''<='" & frm040205.txt1(11) & "' "
   End If
   'End 2003/12/04
   
   '2007/11/27 add by sonia 非個人時再加印財務調整傳票
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
   '2007/11/27 end
   
   'Modify By Cheng 2002/02/15 '多加入案件性質的條件
   '92.4.2 不抓CASEPROGRESS否則資料會重覆
   'Modify By Cheng 2003/12/16 '加智權人員欄位(CP13)
   '2005/5/3 MODIFY BY SONIA 抓欄位時取消CP13,會造成一請款單多收文號時資料重覆
   '2007/11/27 modify by sonia同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
   'strSQL = "SELECT A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Y06 AS 外幣金額,A0Y04 AS 台幣金額,'' AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'', A1K01, CP01, CP02, CP03, CP04, NULL  FROM ACC0Y0,ACC1K0,acc0z0, CaseProgress " & _
   '         " WHERE A1K01=CP60 And A1K13=CP01 And A1K14=CP02 And A1K15=CP03 And A1K16=CP04 And A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND a0z01=A0Y01(+) and a0z02=a1k01(+) " & strSQLCP & _
   '         " Group By A0Y01 ,A0Y02 ,A0Y03 ,A0Y06 ,A0Y04 ,A0Y07, A1K13, A1K01, CP01, CP02, CP03, CP04 "
   '2011/4/12 modify by sonia 同一案號不同請款單同時收款故要加a1k01,另因同時收款傳票會有二組故要加distinct
   'Modify By Sindy 2017/4/17
'   strSql = "SELECT A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Z04 AS 外幣金額,round(A0Z04*A0Y04,2) AS 台幣金額,decode(round(A0Z04*A0Y04,2)-a1k30,0,A1K09,0) AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'',  CP01, CP02, CP03, CP04, NULL, a1k01 " & _
'            " FROM fagent,CaseProgress,acc1p0, acc021, " & _
'            "(SELECT distinct A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,MAX(CP05||CP09) CP,a1k01 FROM ACC0Y0,ACC1K0,acc0z0,CASEPROGRESS " & _
'            " WHERE A0Y02>=" & Val(frm040205.Txt1(2)) & " AND A0Y02<=" & Val(frm040205.Txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A0Z02=CP60(+)" & strSalesArea & strSQL1k0 & "GROUP BY A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,a1k01) NEW " & _
'            " WHERE cp09 in substr(new.cp,9,9) and " & SQLNewFag("A0Y07", "fa") & strSQLCP & "and new.a0y01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+)" & str1P0Sales & _
'            "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'            "union " & _
'            "select '' AS 單據編號,A0205 AS 單據日期,'' AS 幣別,0 AS 外幣金額,ax207-ax206 AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,substr(ax214, 1, Length(ax214) - 9) ,'' , substr(ax214, 1, Length(ax214) - 9), substr(ax214, Length(ax214) - 8, 6), substr(ax214, Length(ax214) - 2, 1), substr(ax214, Length(ax214) - 1, 2),null , null " & _
'            "from acc020,acc021,acc1p0 where a0205>=" & Val(frm040205.Txt1(2)) & " and a0205<=" & Val(frm040205.Txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & "and a0202=a1p22(+) and 'F'=a1p02(+) and a1p04 is null union " & _
'            "select '' AS 單據編號,A0205 AS 單據日期,'' AS 幣別,0 AS 外幣金額,ax207-ax206 AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,substr(ax214, 1, Length(ax214) - 9) ,'' , substr(ax214, 1, Length(ax214) - 9), substr(ax214, Length(ax214) - 8, 6), substr(ax214, Length(ax214) - 2, 1), substr(ax214, Length(ax214) - 1, 2),null , null " & _
'            "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress where a0205>=" & Val(frm040205.Txt1(2)) & " and a0205<=" & Val(frm040205.Txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
'            "and ax202=a1p22(+) and 'F'=a1p02(+) and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null "
   SetGrd2 'Add By Sindy 2017/4/17
   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
'   strSql = "SELECT 幣別,sum(外幣金額) AS 外幣金額,sum(台幣金額) AS 台幣金額,sum(規費) AS 規費,' ' AS 收款點數 from (" & _
'            "SELECT A0Y01 AS 單據編號,A0Y02 AS 單據日期,A0Y03 AS 幣別,A0Z04 AS 外幣金額,round(A0Z04*A0Y04,2) AS 台幣金額,decode(round(A0Z04*A0Y04,2)-a1k30,0,A1K09,0) AS 規費,'' AS 結清,A0Y07 AS 代理人,'' AS 翻譯費,A1K13,'',  CP01, CP02, CP03, CP04, NULL, a1k01 " & _
'            " FROM fagent,CaseProgress,acc1p0, acc021, " & _
'            "(SELECT distinct A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,MAX(CP05||CP09) CP,a1k01 FROM ACC0Y0,ACC1K0,acc0z0,CASEPROGRESS " & _
'            " WHERE A0Y02>=" & Val(frm040205.txt1(2)) & " AND A0Y02<=" & Val(frm040205.txt1(3)) & " AND A0Y01=a0z01(+) and a0z02=a1k01(+) AND A0Z02=CP60(+)" & strSalesArea & strSQL1k0 & "GROUP BY A0Y01,A0Y02,A0Y03,A0Y04,A0Y06,A0Y07,A1K03,A1K09,A1K13,A1K14,A1K15,A1K16,A1K29,a0z04,a1k30,a1k01) NEW " & _
'            " WHERE cp09 in substr(new.cp,9,9) and " & SQLNewFag("A0Y07", "fa") & strSQLCP & "and new.a0y01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+)" & str1P0Sales & _
'            "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'            "union " & _
'            "select '' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,ax207-ax206 AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,substr(ax214, 1, Length(ax214) - 9) ,'' , substr(ax214, 1, Length(ax214) - 9), substr(ax214, Length(ax214) - 8, 6), substr(ax214, Length(ax214) - 2, 1), substr(ax214, Length(ax214) - 1, 2),null , null " & _
'            "from acc020,acc021,acc1p0 where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & "and a0202=a1p22(+) and 'F'=a1p02 and a1p04 is null union " & _
'            "select '' AS 單據編號,A0205 AS 單據日期,'NTD' AS 幣別,0 AS 外幣金額,ax207-ax206 AS 台幣金額,0 AS 規費,'' AS 結清,'' AS 代理人,'' AS 翻譯費,substr(ax214, 1, Length(ax214) - 9) ,'' , substr(ax214, 1, Length(ax214) - 9), substr(ax214, Length(ax214) - 8, 6), substr(ax214, Length(ax214) - 2, 1), substr(ax214, Length(ax214) - 1, 2),null , null " & _
'            "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress where a0205>=" & Val(frm040205.txt1(2)) & " and a0205<=" & Val(frm040205.txt1(3)) & " and a0202=ax202 and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
'            "and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
'            ") group by 幣別"
   'Modify By Sindy 2018/7/20 抓資料改到 accrpt0205
   strSql = "select R21218 AS 幣別,sum(R21207) AS 外幣金額,sum(R21208) AS 台幣金額,sum(R21206) AS 規費,sum(R21209) AS 收款點數" & _
            " from accrpt0205 where R21201='" & strUserNum & "'" & _
            " group by R21218"
   CheckOC
'2007/11/27 cancel by sonia
'   If Len(frm040205.txt1(0)) <> 0 Then
'      strTemp = Split(frm040205.txt1(0), ",")
'   End If
'   'If Len(frm040205.txt1(0)) <> 0 Then
'   '   strSQL = strSQL & " and a1k13 in (" & GetAddStr(frm040205.txt1(0)) & ") "
'   'End If
'2007/11/27 end
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'2007/11/27 cancel by sonia
'      adoRecordset.MoveFirst
'      Do While adoRecordset.EOF = False
'         s = 0
'         If Len(frm040205.txt1(0)) <> 0 Then
'            If Not IsNull(adoRecordset.Fields(9)) Then
'               For i = 0 To UBound(strTemp)
'                  If strTemp(i) = adoRecordset.Fields(9) Then
'                     s = 1
'                  End If
'               Next i
'            End If
'         End If
'         If Len(frm040205.txt1(4)) <> 0 And s = 1 Then
'             'Modify By Cheng 2003/05/29
'      '        If Val(GetPrjNationNumber(adoRecordset.Fields(7))) >= Val(frm040205.txt1(4)) Then
'            If "" & GetPrjNationNumber(adoRecordset.Fields(7)) >= frm040205.txt1(4) Then
'                s = 1
'            Else
'                s = 0
'            End If
'         End If
'         If Len(frm040205.txt1(5)) <> 0 And s = 1 Then
'             'Modify By Cheng 2003/05/29
'      '        If Val(GetPrjNationNumber(adoRecordset.Fields(7))) <= Val(frm040205.txt1(5)) Then
'            If "" & GetPrjNationNumber(adoRecordset.Fields(7)) <= frm040205.txt1(5) & "z" Then
'                s = 1
'            Else
'                s = 0
'            End If
'         End If
'         If s = 0 Then
'             adoRecordset.Delete
'         Else
'             'adoRecordset.Fields(3) = adoRecordset.Fields(3) * GetPrjUSTotleByNick(adoRecordset.Fields(0))
'             'adoRecordset.Fields(4) = adoRecordset.Fields(3) * adoRecordset.Fields(4)
'             'adoRecordset.Fields(7) = GetPrjName2(adoRecordset.Fields(7))
'         End If
'         adoRecordset.MoveNext
'      Loop
   Else
'       'ShowNoData
'       'Me.Hide
      Exit Sub
   End If
   
'   Set GRD1.Recordset = adoRecordset
   Set grd2.Recordset = adoRecordset
   SetGrd2
   Me.Enabled = False
   Dim dblNT As Double, dblFee As Double
   With grd2
      grd2.Visible = False
      For i = 1 To .Rows - 1
         dblNT = 0: dblFee = 0
         .row = i
         .col = 1 '外幣金額
         .Text = Format(.Text, "###,###,###.00")
         .col = 2 '台幣金額
         dblNT = .Text
         .Text = Format(.Text, "###,###,##0")
         .col = 3 '規費
         dblFee = .Text
         .col = 4 '收款點數
         'Modify By Sindy 2018/7/23
         '.Text = Format(str((dblNT - dblFee) / 1000), "###,###,###.000")
         .Text = Format(.Text, "###,###,###.000")
         '2018/7/23 END
      Next i
      grd2.Visible = True
   End With
   '2012/10/11 End
   CheckOC
   
   'Add By Sindy 2018/7/20
   strSql = "select sum(R21208) AS 台幣金額,sum(R21209) AS 收款點數" & _
            " from accrpt0205 where R21201='" & strUserNum & "'"
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      IntTo3 = Format(Val(adoRecordset.Fields("台幣金額")), "###,###,##0") '台幣
      IntTo5 = Format(Val(adoRecordset.Fields("收款點數")), "###,###,###.000") 'Format(str((IntTo3 - Val(adoRecordset.Fields("規費"))) / 1000), "###,###,###.000") '點數
   End If
   CheckOC
   '2018/7/20 END
   
   Me.Enabled = True
'   Dim IntTestTemp As Double     '台幣金額
'   Dim IntTestTemp2 As Double   '外幣金額
'   Set Grd1.Recordset = adoRecordset
'   For i = 1 To Grd1.Rows - 1
'      Grd1.row = i
'      Grd1.col = 0
'      strSQL = Grd1.Text
'   '   strSQL = "SELECT A0Z04,A0Y04 FROM ACC0Z0,ACC0Y0 WHERE A0Z01='" & grd1.Text & "' AND A0Y01=A0Z01 "
'      strSQL = "SELECT A0Z04,A0Y04 FROM ACC0Z0,ACC0Y0, ACC1K0, (Select CP01, CP02, CP03, CP04, CP60 From CaseProgress Where " & ChgCaseprogress(Me.Grd1.TextMatrix(i, 12) & Me.Grd1.TextMatrix(i, 13) & Me.Grd1.TextMatrix(i, 14) & Me.Grd1.TextMatrix(i, 15)) & " And CP60='" & Me.Grd1.TextMatrix(i, 11) & "' Group By CP01, CP02, CP03, CP04, CP60 ) C1 WHERE A0Z01='" & Grd1.Text & "' AND A0Y01=A0Z01 And A0Z02=A1K01 And A1K01=C1.CP60 And A1K13=C1.CP01 And A1K14=C1.CP02 And A1K15=C1.CP03 And A1K16=C1.CP04 And A1K01='" & Me.Grd1.TextMatrix(i, 11) & "' And C1.CP01||C1.CP02||C1.CP03||C1.CP04='" & Me.Grd1.TextMatrix(i, 12) & Me.Grd1.TextMatrix(i, 13) & Me.Grd1.TextMatrix(i, 14) & Me.Grd1.TextMatrix(i, 15) & "' "
'      CheckOC3
'      AdoRecordSet3.CursorLocation = adUseClient
'      AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
'         IntTestTemp = 0
'         IntTestTemp2 = 0
'         AdoRecordSet3.MoveFirst
'         Do While AdoRecordSet3.EOF = False
'            IntTestTemp = IntTestTemp + (Val(CheckStr(AdoRecordSet3.Fields(0))) * Val(CheckStr(AdoRecordSet3.Fields(1))))
'            IntTestTemp2 = IntTestTemp2 + Val(CheckStr(AdoRecordSet3.Fields(0)))
'            AdoRecordSet3.MoveNext
'         Loop
'         Grd1.col = 3
'         Grd1.Text = Format(str(IntTestTemp2), "0.00")
'         Grd1.CellAlignment = flexAlignRightCenter
'         Grd1.col = 4
'         Grd1.Text = Format(str(IntTestTemp), "0.00")
'         Grd1.CellAlignment = flexAlignRightCenter
'      Else
'         Grd1.col = 3
'         Grd1.Text = "0.00"
'         Grd1.CellAlignment = flexAlignRightCenter
'         Grd1.col = 4
'         Grd1.Text = "0.00"
'         Grd1.CellAlignment = flexAlignRightCenter
'      End If
'      CheckOC3
'      Grd1.col = 7
'      Grd1.Text = GetPrjName2(Grd1.Text)
'      Grd1.CellAlignment = flexAlignRightCenter
'   Next i
'2007/11/27 end
   CheckOC
End Sub
