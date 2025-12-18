VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100111_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人收/發文量查詢"
   ClientHeight    =   5730
   ClientLeft      =   30
   ClientTop       =   990
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9320
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7272
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8496
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
      Height          =   4740
      Left            =   48
      TabIndex        =   2
      Top             =   960
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   8361
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      HighLight       =   0
      ScrollBars      =   2
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
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
      Height          =   4740
      Left            =   4800
      TabIndex        =   4
      Top             =   936
      Width           =   4476
      _ExtentX        =   7885
      _ExtentY        =   8361
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      HighLight       =   0
      ScrollBars      =   2
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
      _Band(0).Cols   =   3
   End
   Begin MSForms.Label lbl2 
      Height          =   300
      Left            =   5580
      TabIndex        =   9
      Top             =   444
      Width           =   2232
      Size            =   "3937;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   1
      Left            =   4770
      TabIndex        =   8
      Top             =   450
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "機關來文"
      Height          =   180
      Left            =   4770
      TabIndex        =   7
      Top             =   690
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "接洽及內部收文單"
      Height          =   180
      Left            =   75
      TabIndex        =   6
      Top             =   690
      Width           =   1440
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   1152
      TabIndex        =   5
      Top             =   408
      Width           =   2232
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收文期間："
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   3
      Top             =   405
      Width           =   900
   End
End
Attribute VB_Name = "frm100111_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(承辦人姓名lbl1(1)改lbl2,grdDataList1及grdDataList2改Fonts)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSql As String, strTemp As Variant, StrTest As String, StrOkorNo As String
Dim i As Integer, j As Integer, s As Integer, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
   grdDataList1.Cols = 3
   grdDataList1.row = 0
   grdDataList1.col = 0: grdDataList1.Text = "案件性質"
   grdDataList1.ColWidth(0) = 1800
   grdDataList1.CellAlignment = flexAlignCenterCenter
   grdDataList1.col = 1: grdDataList1.Text = "數量"
   grdDataList1.ColWidth(1) = 600
   grdDataList1.CellAlignment = flexAlignCenterCenter
   grdDataList1.col = 2: grdDataList1.Text = ""
   grdDataList1.ColWidth(2) = 0
   grdDataList1.CellAlignment = flexAlignCenterCenter
   grdDataList2.Cols = 3
   grdDataList2.row = 0
   grdDataList2.col = 0: grdDataList2.Text = "案件性質"
   grdDataList2.ColWidth(0) = 1800
   grdDataList2.CellAlignment = flexAlignCenterCenter
   grdDataList2.col = 1: grdDataList2.Text = "數量"
   grdDataList2.ColWidth(1) = 600
   grdDataList2.CellAlignment = flexAlignCenterCenter
   grdDataList2.col = 2: grdDataList2.Text = ""
   grdDataList2.ColWidth(2) = 0
   grdDataList2.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
   '92.04.16 nick 以下無效
   Select Case Index
      Case 0
         Me.Hide
      Case 1
         bolToEndByNick = True
        Unload Me
        Exit Sub
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   Screen.MousePointer = vbHourglass
      MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
   Screen.MousePointer = vbHourglass
End Sub

Sub StrMenu()
   lbl1(0).Caption = frm100111_1.txt1(2) + "-" + frm100111_1.txt1(3)
   'DoEvents
   'modify by sonia 2022/1/20 承辦人姓名lbl1(1)改lbl2
   lbl2.Caption = frm100111_1.lbl1.Caption
   
   Me.Enabled = False
   '寫入暫存檔
   strSQL1 = ""
   'StrSQL2 = ""
   'StrSQL3 = ""
   'StrSQL4 = ""
   'StrSQL5 = ""
   If frm100111_1.txt1(1) = "1" Then
      If Len(Trim(frm100111_1.txt1(2))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm100111_1.txt1(2))) & " "
      End If
      If Len(Trim(frm100111_1.txt1(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(frm100111_1.txt1(3))) & " "
      End If
      If Len(Trim(frm100111_1.txt1(2))) <> 0 Or Len(Trim(frm100111_1.txt1(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";收文" & frm100111_1.Label3 & frm100111_1.txt1(2) & "-" & frm100111_1.txt1(3) 'Add By Sindy 2010/11/4
      End If
   Else
      If Len(Trim(frm100111_1.txt1(2))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(frm100111_1.txt1(2))) & " "
      End If
      If Len(Trim(frm100111_1.txt1(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(frm100111_1.txt1(3))) & " "
      End If
      If Len(Trim(frm100111_1.txt1(2))) <> 0 Or Len(Trim(frm100111_1.txt1(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";發文" & frm100111_1.Label3 & frm100111_1.txt1(2) & "-" & frm100111_1.txt1(3) 'Add By Sindy 2010/11/4
      End If
      Label3(0) = "發文期間："
   End If
   If Len(Trim(frm100111_1.txt1(4))) <> 0 Then
       strSQL1 = strSQL1 + " AND CP10>='" & frm100111_1.txt1(4) & "' "
   End If
   If Len(Trim(frm100111_1.txt1(5))) <> 0 Then
       strSQL1 = strSQL1 + " AND CP10<='" & frm100111_1.txt1(5) & "' "
   End If
   If Len(Trim(frm100111_1.txt1(4))) <> 0 Or Len(Trim(frm100111_1.txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100111_1.Label5(0) & frm100111_1.txt1(4) & "-" & frm100111_1.txt1(5) 'Add By Sindy 2010/11/4
   End If
   
   strSQL2 = strSQL1
   StrSQL3 = strSQL1
   StrSQL4 = strSQL1
   strSQL5 = strSQL1
   If Len(Trim(frm100111_1.txt1(6))) <> 0 Then
      'Modify By Cheng 2002/03/14
   '   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(frm100111_1.txt1(6), 1) & ") "
   '   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(frm100111_1.txt1(6), 2) & ") "
   '   StrSQL3 = StrSQL3 & " and cp01 in (" & SQLGrpStr(frm100111_1.txt1(6), 3) & ") "
   '   StrSQL4 = StrSQL4 & " and cp01 in (" & SQLGrpStr(frm100111_1.txt1(6), 4) & ") "
   '   StrSQL5 = StrSQL5 & " and cp01 in (" & SQLGrpStr(frm100111_1.txt1(6), 5) & ") "
      strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(IIf(frm100111_1.txt1(6).Text <> "ALL", frm100111_1.txt1(6).Text, GetAllSysKind(frm100111_1.txt1(6))), 1) & ") "
      strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(IIf(frm100111_1.txt1(6).Text <> "ALL", frm100111_1.txt1(6).Text, GetAllSysKind(frm100111_1.txt1(6))), 2) & ") "
      StrSQL3 = StrSQL3 & " and cp01 in (" & SQLGrpStr(IIf(frm100111_1.txt1(6).Text <> "ALL", frm100111_1.txt1(6).Text, GetAllSysKind(frm100111_1.txt1(6))), 3) & ") "
      StrSQL4 = StrSQL4 & " and cp01 in (" & SQLGrpStr(IIf(frm100111_1.txt1(6).Text <> "ALL", frm100111_1.txt1(6).Text, GetAllSysKind(frm100111_1.txt1(6))), 4) & ") "
      strSQL5 = strSQL5 & " and cp01 in (" & SQLGrpStr(IIf(frm100111_1.txt1(6).Text <> "ALL", frm100111_1.txt1(6).Text, GetAllSysKind(frm100111_1.txt1(6))), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Left(frm100111_1.Label5(1), 5) & frm100111_1.txt1(6) 'Add By Sindy 2010/11/4
   End If
   
   If Len(frm100111_1.txt1(0)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100111_1.Label1 & frm100111_1.txt1(0) & frm100111_1.lbl1 'Add By Sindy 2010/11/4
   End If
   If frm100111_1.txt1(7) = "N" Then
      pub_QL05 = pub_QL05 & ";" & Left(frm100111_1.Label5(2), 11) & frm100111_1.txt1(7) 'Add By Sindy 2010/11/4
   End If
   
   cnnConnection.Execute "delete from r100111 where id='" & strUserNum & "' "
   CheckOC
   'strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,CP09,CP10 FROM CASEPROGRESS WHERE CP14='" & frm100111_1.Txt1(0) & "' AND CP26 IS NULL AND CP57 IS NULL "
   'Modify By Cheng 2002/01/14
   '若不統計不計件之案件, 才設定CP26 IS NULL 的條件
                   strSql = "SELECT CP09,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),'" & strUserNum & "' FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP WHERE CP14='" & frm100111_1.txt1(0) & "' " & IIf(frm100111_1.txt1(7).Text = "N", " AND CP26 IS NULL ", " ") & " AND CP57 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
   strSql = strSql & " union all select CP09,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP WHERE CP14='" & frm100111_1.txt1(0) & "' " & IIf(frm100111_1.txt1(7).Text = "N", " AND CP26 IS NULL ", " ") & " AND CP57 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
   strSql = strSql & " union all select CP09,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10),'" & strUserNum & "' FROM CASEPROGRESS,LAWCASE,CASEPROPERTYMAP WHERE CP14='" & frm100111_1.txt1(0) & "' " & IIf(frm100111_1.txt1(7).Text = "N", " AND CP26 IS NULL ", " ") & " AND CP57 IS NULL AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3
   strSql = strSql & " union all select CP09,NVL(CPM03,CP10),'" & strUserNum & "' FROM CASEPROGRESS,HIRECASE,CASEPROPERTYMAP WHERE CP14='" & frm100111_1.txt1(0) & "' " & IIf(frm100111_1.txt1(7).Text = "N", " AND CP26 IS NULL ", " ") & " AND CP57 IS NULL AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4
   strSql = strSql & " union all select CP09,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP WHERE CP14='" & frm100111_1.txt1(0) & "' " & IIf(frm100111_1.txt1(7).Text = "N", " AND CP26 IS NULL ", " ") & " AND CP57 IS NULL AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5
   strSql = " insert into r100111 " & strSql
   cnnConnection.Execute strSql
   CheckOC
   strSql = "select * from r100111 where id='" & strUserNum & "' "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/4
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/11/4
       Me.Enabled = True
       ShowNoData
       Screen.MousePointer = vbDefault
       '92.04.18 nick
       'Me.Hide
       tmpBol = fnCancelNowFormAndShowParentForm(Me)
       Exit Sub
   End If
   CheckOC
   
   '開始找資料
   'A,B
   strSql = "SELECT R05002 AS 案件性質,COUNT(*) AS 數量 FROM R100111 WHERE R05001<'C' and id='" & strUserNum & "' GROUP BY R05002 "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grdDataList1.Recordset = adoRecordset
   CheckOC
   strSql = "SELECT R05002 AS 案件性質,COUNT(*) AS 數量 FROM R100111 WHERE R05001>'C' and id='" & strUserNum & "' GROUP BY R05002 "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grdDataList2.Recordset = adoRecordset
   CheckOC
   SetDataListWidth
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100111_2 = Nothing
End Sub
