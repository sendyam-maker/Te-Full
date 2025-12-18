VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010013_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所收文量查詢"
   ClientHeight    =   5730
   ClientLeft      =   5445
   ClientTop       =   3390
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5550
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   1845
      Left            =   3750
      TabIndex        =   16
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3254
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3510
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4740
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
      Height          =   4815
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   8493
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "各系統類別筆數："
      Height          =   180
      Left            =   3780
      TabIndex        =   17
      Top             =   3630
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "PS：查詢數量包含              收文無費用者及          國外來函請款者         (例：證書費)"
      ForeColor       =   &H00000000&
      Height          =   780
      Index           =   1
      Left            =   3750
      TabIndex        =   15
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "所　別："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   2640
      TabIndex        =   14
      Top             =   555
      Width           =   735
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   3480
      TabIndex        =   13
      Top             =   555
      Width           =   735
      ForeColor       =   0
      VariousPropertyBits=   27
      Size            =   "1296;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   6
      Left            =   4170
      TabIndex        =   12
      Top             =   1620
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2275;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Caption         =   "案件性質："
      Height          =   180
      Left            =   3810
      TabIndex        =   11
      Top             =   1380
      Width           =   975
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   5
      Left            =   4170
      TabIndex        =   10
      Top             =   1140
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2275;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "申請國家："
      Height          =   180
      Left            =   3810
      TabIndex        =   9
      Top             =   900
      Width           =   975
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   4
      Left            =   4170
      TabIndex        =   8
      Top             =   2535
      Visible         =   0   'False
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2275;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   3
      Left            =   4170
      TabIndex        =   7
      Top             =   2055
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2275;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "總計："
      Height          =   180
      Left            =   3795
      TabIndex        =   6
      Top             =   2325
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   3795
      TabIndex        =   5
      Top             =   1845
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   2
      Top             =   555
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2117;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "收文日期："
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   555
      Width           =   975
   End
End
Attribute VB_Name = "frm010013_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 lbl1()
'Memo By Sonia 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

'2004/1/27
Private Sub SetDataListWidth()
'    edit by  nick 2004/10/08
'    grdDataList1.Cols = 3
'    grdDataList1.Row = 0
'    grdDataList1.Col = 0: grdDataList1.Text = "業務區"
'    grdDataList1.ColWidth(0) = 1600
'    grdDataList1.CellAlignment = flexAlignCenterCenter
'    grdDataList1.Col = 1: grdDataList1.Text = "數量"
'    grdDataList1.ColWidth(1) = 1000
'    grdDataList1.CellAlignment = flexAlignCenterCenter
'    'add by nick 2004/10/08
'    grdDataList1.Col = 2: grdDataList1.Text = "點數"
'    grdDataList1.ColWidth(2) = 1000
'    grdDataList1.CellAlignment = flexAlignCenterCenter
    grdDataList1.Cols = 4
    grdDataList1.row = 0
    grdDataList1.col = 0: grdDataList1.Text = "部門"
    grdDataList1.ColWidth(0) = 1000
    grdDataList1.CellAlignment = flexAlignCenterCenter
    grdDataList1.col = 1: grdDataList1.Text = "智權人員"
    grdDataList1.ColWidth(1) = 800
    grdDataList1.CellAlignment = flexAlignCenterCenter
    grdDataList1.col = 2: grdDataList1.Text = "數量"
    grdDataList1.ColWidth(2) = 700
    grdDataList1.CellAlignment = flexAlignCenterCenter
    grdDataList1.col = 3: grdDataList1.Text = "點數"
    grdDataList1.ColWidth(3) = 700
    grdDataList1.CellAlignment = flexAlignCenterCenter
End Sub

'2004/1/27
Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
        Case 0
            frm010013_1.Show
        Case 1
            Unload frm010013_1
    End Select
    Unload Me
End Sub

'2004/1/27
Private Sub Form_Load()
    MoveFormToCenter Me
    SetDataListWidth
End Sub

'2004/1/27
Sub StrMenu()
    Me.Enabled = False
    '讀出資料
    If DoTemp = False Then
       frm010013_1.Show
       Screen.MousePointer = vbDefault
       Unload Me
       Exit Sub
    End If
    '顯示表單資料
    '收文日
    'edit by nick 2004/10/13
    If frm010013_1.Option1(0).Value = True Then
        Label3.Caption = "收文月份："
        lbl1(0).Caption = frm010013_1.txt1(0)
    Else
        Label3.Caption = "收文日期："
        lbl1(0).Caption = frm010013_1.txt1(1)
    End If
    '所別
    lbl1(1).Caption = frm010013_1.lbl1(2)
    '申請國家
    Me.lbl1(5).Caption = ""
    If frm010013_1.txt1(9).Text <> "" Or frm010013_1.txt1(10).Text <> "" Then
        Me.lbl1(5).Caption = frm010013_1.txt1(9).Text & "－" & frm010013_1.txt1(10).Text
    End If
    '案件性質
    Me.lbl1(6).Caption = ""
    If frm010013_1.txt1(13).Text <> "" Or frm010013_1.txt1(14).Text <> "" Then
        Me.lbl1(6).Caption = frm010013_1.txt1(13).Text & "－" & frm010013_1.txt1(14).Text
    End If
    '智權人員
    lbl1(3).Caption = ""
    If frm010013_1.lbl1(1) <> "" Then
        lbl1(3).Caption = frm010013_1.lbl1(1)
    End If
    grdDataList1.Visible = True
    Me.Enabled = True
End Sub

Function DoTemp() As Boolean
Dim strSql As String, strCon As String
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
'add by nickc 2008/03/12
Dim m_str As String, m_strcon As String, m_str1 As String, m_str2 As String, m_str3 As String, m_str4 As String, m_str5 As String
Dim m_rs As New ADODB.Recordset

'add by nick 2004/10/13
cnnConnection.Execute "delete from r010013_1 where id='" & strUserNum & "' "
cnnConnection.Execute "delete from r010013_2 where id='" & strUserNum & "' "
    
    frm010013_1.Hide
    'edit by nick 2004/10/08
    'strSQL = "SELECT T4.A0902, COUNT(*) FROM CASEPROGRESS T1, STAFF T2, ACC090 T4 WHERE T1.CP09<'B' AND T1.CP13=T2.ST01 AND T1.CP12=T4.A0901(+)"
    '2005/6/2 MODIFY BY SONIA 加入C來函有費用者
    'strSQL = "SELECT T1.cp12 as cp12,T1.CP13 as cp13, COUNT(*),to_char(sum(T1.CP18),'999999D9'),'" & strUserNum & "' FROM CASEPROGRESS T1,staff T2 WHERE T1.CP09<'B'  and t1.cp13=t2.st01 "
    strSql = "SELECT T1.cp12 as cp12,T1.CP13 as cp13, COUNT(*),to_char(sum(T1.CP18),'999999D9'),'" & strUserNum & "' FROM CASEPROGRESS T1,staff T2 WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0))  and t1.cp13=t2.st01 "
    'add by nickc 2008/03/12
    m_str = "select t1.cp01,sum(decode(substr(t1.CP09,1,1),'A',1,0)),sum(decode(substr(t1.CP09,1,1),'C',1,0)) from caseprogress t1,staff t2 where (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0))  and t1.cp13=t2.st01 "
    '2005/6/2 END
'edit by nick 2004/10/08
'    strSQL1 = "SELECT T1.CP12 FROM CASEPROGRESS T1, STAFF T2, PATENT T3" & _
'                " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01 AND T1.CP01=T3.PA01 AND T1.CP02=T3.PA02 AND T1.CP03=T3.PA03 AND T1.CP04=T3.PA04"
'    strSQL2 = "SELECT T1.CP12 FROM CASEPROGRESS T1, STAFF T2, TRADEMARK T3" & _
'                " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01 AND T1.CP01=T3.TM01 AND T1.CP02=T3.TM02 AND T1.CP03=T3.TM03 AND T1.CP04=T3.TM04"
'    StrSQL3 = "SELECT T1.CP12 FROM CASEPROGRESS T1, STAFF T2, LAWCASE T3" & _
'                " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01 AND T1.CP01=T3.LC01 AND T1.CP02=T3.LC02 AND T1.CP03=T3.LC03 AND T1.CP04=T3.LC04"
'    StrSQL4 = "SELECT T1.CP12 FROM CASEPROGRESS T1, STAFF T2, HIRECASE T3" & _
'                " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01 AND T1.CP01=T3.HC01 AND T1.CP02=T3.HC02 AND T1.CP03=T3.HC03 AND T1.CP04=T3.HC04"
'    strSQL5 = "SELECT T1.CP12 FROM CASEPROGRESS T1, STAFF T2, SERVICEPRACTICE T3" & _
'                " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01 AND T1.CP01=T3.SP01 AND T1.CP02=T3.SP02 AND T1.CP03=T3.SP03 AND T1.CP04=T3.SP04"
    '2005/6/2 MODIFY BY SONIA 加入C來函有費用者
    'strSQL1 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, PATENT T3" & _
    '            " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.PA01(+) AND T1.CP02=T3.PA02(+) AND T1.CP03=T3.PA03(+) AND T1.CP04=T3.PA04(+) and t1.cp01 in (select sk01 from systemkind where sk02=1) "
    'strSQL2 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, TRADEMARK T3" & _
    '            " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.TM01(+) AND T1.CP02=T3.TM02(+) AND T1.CP03=T3.TM03(+) AND T1.CP04=T3.TM04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=2) "
    'StrSQL3 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, LAWCASE T3" & _
    '            " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.LC01(+) AND T1.CP02=T3.LC02(+) AND T1.CP03=T3.LC03(+) AND T1.CP04=T3.LC04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=3) "
    'StrSQL4 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, HIRECASE T3" & _
    '            " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.HC01(+) AND T1.CP02=T3.HC02(+) AND T1.CP03=T3.HC03(+) AND T1.CP04=T3.HC04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=4) "
    'strSQL5 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, SERVICEPRACTICE T3" & _
    '            " WHERE T1.CP09<'B' AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.SP01(+) AND T1.CP02=T3.SP02(+) AND T1.CP03=T3.SP03(+) AND T1.CP04=T3.SP04(+)  and t1.cp01 in (select sk01 from systemkind where sk02 in (5,6,7,8)) "
    strSQL1 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, PATENT T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.PA01(+) AND T1.CP02=T3.PA02(+) AND T1.CP03=T3.PA03(+) AND T1.CP04=T3.PA04(+) and t1.cp01 in (select sk01 from systemkind where sk02=1) "
    strSQL2 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, TRADEMARK T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.TM01(+) AND T1.CP02=T3.TM02(+) AND T1.CP03=T3.TM03(+) AND T1.CP04=T3.TM04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=2) "
    StrSQL3 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, LAWCASE T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.LC01(+) AND T1.CP02=T3.LC02(+) AND T1.CP03=T3.LC03(+) AND T1.CP04=T3.LC04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=3) "
    StrSQL4 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, HIRECASE T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.HC01(+) AND T1.CP02=T3.HC02(+) AND T1.CP03=T3.HC03(+) AND T1.CP04=T3.HC04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=4) "
    strSQL5 = "SELECT T1.CP12 as CP12,T1.CP13 as CP13,T1.CP18 as CP18 FROM CASEPROGRESS T1, STAFF T2, SERVICEPRACTICE T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.SP01(+) AND T1.CP02=T3.SP02(+) AND T1.CP03=T3.SP03(+) AND T1.CP04=T3.SP04(+)  and t1.cp01 in (select sk01 from systemkind where sk02 in (5,6,7,8)) "
    'add by nickc 2008/03/12
    m_str1 = "SELECT T1.CP01 as CP01,T1.CP09 as CP09 FROM CASEPROGRESS T1, STAFF T2, PATENT T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.PA01(+) AND T1.CP02=T3.PA02(+) AND T1.CP03=T3.PA03(+) AND T1.CP04=T3.PA04(+) and t1.cp01 in (select sk01 from systemkind where sk02=1) "
    m_str2 = "SELECT T1.CP01 as CP01,T1.CP09 as CP09 FROM CASEPROGRESS T1, STAFF T2, TRADEMARK T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.TM01(+) AND T1.CP02=T3.TM02(+) AND T1.CP03=T3.TM03(+) AND T1.CP04=T3.TM04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=2) "
    m_str3 = "SELECT T1.CP01 as CP01,T1.CP09 as CP09 FROM CASEPROGRESS T1, STAFF T2, LAWCASE T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.LC01(+) AND T1.CP02=T3.LC02(+) AND T1.CP03=T3.LC03(+) AND T1.CP04=T3.LC04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=3) "
    m_str4 = "SELECT T1.CP01 as CP01,T1.CP09 as CP09 FROM CASEPROGRESS T1, STAFF T2, HIRECASE T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.HC01(+) AND T1.CP02=T3.HC02(+) AND T1.CP03=T3.HC03(+) AND T1.CP04=T3.HC04(+)  and t1.cp01 in (select sk01 from systemkind where sk02=4) "
    m_str5 = "SELECT T1.CP01 as CP01,T1.CP09 as CP09 FROM CASEPROGRESS T1, STAFF T2, SERVICEPRACTICE T3" & _
                " WHERE (T1.CP09<'B' OR (T1.CP09>'C' AND T1.CP16 IS NOT NULL AND T1.CP16>0)) AND T1.CP13=T2.ST01(+) AND T1.CP01=T3.SP01(+) AND T1.CP02=T3.SP02(+) AND T1.CP03=T3.SP03(+) AND T1.CP04=T3.SP04(+)  and t1.cp01 in (select sk01 from systemkind where sk02 in (5,6,7,8)) "
    
    '2005/6/2 END
    '組合條件
    strCon = ""
    strCon1 = ""
    'add by nickc 2008/03/12
    m_strcon = ""
    
    'edit by nick 2004/10/08
    If frm010013_1.Option1(0).Value = True Then
        If Len(Trim(frm010013_1.txt1(0))) > 0 Then
            strCon = strCon + " AND T1.CP05>=" & Val(ChangeTStringToWString(frm010013_1.txt1(0) & "01")) & " AND T1.CP05<=" & Val(ChangeTStringToWString(frm010013_1.txt1(0) & "31")) & " "
            'add by nickc 2008/03/12
            m_strcon = m_strcon + " AND T1.CP05>=" & Val(ChangeTStringToWString(frm010013_1.txt1(0) & "01")) & " AND T1.CP05<=" & Val(ChangeTStringToWString(frm010013_1.txt1(0) & "31")) & " "
        End If
    Else
        '收文日
        If Len(Trim(frm010013_1.txt1(1))) > 0 Then
           strCon = strCon + " AND T1.CP05=" & Val(ChangeTStringToWString(frm010013_1.txt1(1))) & " "
           'add by nickc 2008/03/12
           m_strcon = m_strcon + " AND T1.CP05=" & Val(ChangeTStringToWString(frm010013_1.txt1(1))) & " "
        End If
    End If
    '所別
    If Len(Trim(frm010013_1.lbl1(0))) > 0 Then
        strCon = strCon + " AND T2.ST06='" & frm010013_1.lbl1(0) & "'"
        'add by nickc 2008/03/12
        m_strcon = m_strcon + " AND T2.ST06='" & frm010013_1.lbl1(0) & "'"
    End If
    '智權人員
    If Len(Trim(frm010013_1.txt1(8))) > 0 Then
        strCon = strCon + " AND T1.CP13='" & frm010013_1.txt1(8) & "'"
        'add by nickc 2008/03/12
        m_strcon = m_strcon + " AND T1.CP13='" & frm010013_1.txt1(8) & "'"
    End If
    '案件性質
    If Len(Trim(frm010013_1.txt1(13))) <> 0 Then
        strCon = strCon + " AND T1.CP10>='" & frm010013_1.txt1(13) & "' "
        'add by nickc 2008/03/12
        m_strcon = m_strcon + " AND T1.CP10>='" & frm010013_1.txt1(13) & "' "
    End If
    If Len(Trim(frm010013_1.txt1(14))) <> 0 Then
        strCon = strCon + " AND T1.CP10<='" & frm010013_1.txt1(14) & "' "
        'add by nickc 2008/03/12
        m_strcon = m_strcon + " AND T1.CP10<='" & frm010013_1.txt1(14) & "' "
    End If
    '申請國家
    If Len(Trim(frm010013_1.txt1(9))) = 0 And Len(Trim(frm010013_1.txt1(10))) = 0 Then
        'edit by nick 2004/10/13
        'strSQL = strSQL & strCon & " Group by T4.A0902 ORDER BY 1"
        strSql = strSql & strCon & " Group by T1.cp12,T1.cp13 "
        'add by nickc 2008/03/12
        m_str = m_str & m_strcon & " Group by T1.cp01 "
    Else
        strSQL1 = strSQL1 & strCon
        strSQL2 = strSQL2 & strCon
        StrSQL3 = StrSQL3 & strCon
        StrSQL4 = StrSQL4 & strCon
        strSQL5 = strSQL5 & strCon
        'add by nickc 2008/03/12
        m_str1 = m_str1 & m_strcon
        m_str2 = m_str2 & m_strcon
        m_str3 = m_str3 & m_strcon
        m_str4 = m_str4 & m_strcon
        m_str5 = m_str5 & m_strcon
        If Len(Trim(frm010013_1.txt1(9))) > 0 Then
            strSQL1 = strSQL1 & " AND PA09>='" & frm010013_1.txt1(9) & "'"
            strSQL2 = strSQL2 & " AND TM10>='" & frm010013_1.txt1(9) & "'"
            StrSQL3 = StrSQL3 & " AND LC15>='" & frm010013_1.txt1(9) & "'"
            strSQL5 = strSQL5 & " AND SP09>='" & frm010013_1.txt1(9) & "'"
            'add by nickc 2008/03/12
            m_str1 = m_str1 & " AND PA09>='" & frm010013_1.txt1(9) & "'"
            m_str2 = m_str2 & " AND TM10>='" & frm010013_1.txt1(9) & "'"
            m_str3 = m_str3 & " AND LC15>='" & frm010013_1.txt1(9) & "'"
            m_str5 = m_str5 & " AND SP09>='" & frm010013_1.txt1(9) & "'"
        End If
        If Len(Trim(frm010013_1.txt1(10))) > 0 Then
            strSQL1 = strSQL1 & " AND PA09<='" & frm010013_1.txt1(10) & "'"
            strSQL2 = strSQL2 & " AND TM10<='" & frm010013_1.txt1(10) & "'"
            StrSQL3 = StrSQL3 & " AND LC15<='" & frm010013_1.txt1(10) & "'"
            strSQL5 = strSQL5 & " AND SP09<='" & frm010013_1.txt1(10) & "'"
            'add by nickc 2008/03/12
            m_str1 = m_str1 & " AND PA09<='" & frm010013_1.txt1(10) & "'"
            m_str2 = m_str2 & " AND TM10<='" & frm010013_1.txt1(10) & "'"
            m_str3 = m_str3 & " AND LC15<='" & frm010013_1.txt1(10) & "'"
            m_str5 = m_str5 & " AND SP09<='" & frm010013_1.txt1(10) & "'"
        End If
        'edit by nick 2004/10/08
        'strSQL = "SELECT T4.A0902,COUNT(*) FROM ( " & strSQL1 & " UNION ALL " & strSQL2 & " UNION ALL " & StrSQL3 & " UNION ALL " & StrSQL4 & " UNION ALL " & strSQL5 & " ) X, ACC090 T4 WHERE X.CP12=T4.A0901(+) GROUP BY T4.A0902 ORDER BY 1"
        strSql = "SELECT X.cp12,X.cp13,COUNT(*),to_char(sum(X.CP18),'999999D9'),'" & strUserNum & "' FROM ( " & strSQL1 & " UNION ALL " & strSQL2 & " UNION ALL " & StrSQL3 & " UNION ALL " & StrSQL4 & " UNION ALL " & strSQL5 & " ) X  GROUP BY X.cp12,X.cp13 "
        'add by nickc 2008/03/12
        m_str = "SELECT X.cp01,sum(decode(substr(X.CP09,1,1),'A',1,0)),sum(decode(substr(X.CP09,1,1),'C',1,0)) FROM ( " & strSQL1 & " UNION ALL " & strSQL2 & " UNION ALL " & StrSQL3 & " UNION ALL " & StrSQL4 & " UNION ALL " & strSQL5 & " ) X  GROUP BY X.cp01 "
    End If
    Dim TableDataItem As Long
    Dim tmpR054001 As String
    Dim tmpR054002 As String
    'add by nick 2004/10/13
    cnnConnection.Execute "insert into r010013_1 (" & strSql & ")"
    strSql = "select r054001,r054002,a0902,st02,sum(nvl(r054003,0)),sum(nvl(r054004,0)) from r010013_1,staff,acc090 where r054001=a0901(+) and r054002=st01(+) and id='" & strUserNum & "' group by r054001,r054002,st02,a0902 order by 1,2 "
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    TableDataItem = 1
    If adoRecordset.RecordCount > 0 Then
        adoRecordset.MoveFirst
        tmpR054001 = "=="
        tmpR054002 = "=="
        Do While Not adoRecordset.EOF
            If CheckStr(adoRecordset.Fields(0).Value) <> tmpR054001 Then
                If tmpR054001 <> "==" And tmpR054002 <> "==" Then
                    cnnConnection.Execute "insert into r010013_2 (R055001,r055002,r055003,r055004,r055005,id)  (select " & TableDataItem & ",'部門小計','',sum(nvl(r054003,0)),sum(nvl(r054004,0)),'" & strUserNum & "' from r010013_1 where id='" & strUserNum & "' and r054001='" & tmpR054001 & "' group by  " & TableDataItem & ",'部門小計' ) "
                End If
                tmpR054001 = CheckStr(adoRecordset.Fields(0).Value)
                tmpR054002 = CheckStr(adoRecordset.Fields(1).Value)
                cnnConnection.Execute "insert into r010013_2 (R055001,r055002,r055003,r055004,r055005,id) values ( " & TableDataItem & ",'" & CheckStr(adoRecordset.Fields(2).Value) & "','" & CheckStr(adoRecordset.Fields(3).Value) & "',0" & CheckStr(adoRecordset.Fields(4).Value) & ",0" & CheckStr(adoRecordset.Fields(5).Value) & ",'" & strUserNum & "') "
            ElseIf CheckStr(adoRecordset.Fields(1).Value) <> tmpR054002 Then
                tmpR054002 = CheckStr(adoRecordset.Fields(1).Value)
                cnnConnection.Execute "insert into r010013_2 (R055001,r055002,r055003,r055004,r055005,id) values ( " & TableDataItem & ",'','" & CheckStr(adoRecordset.Fields(3).Value) & "'," & CheckStr(adoRecordset.Fields(4).Value) & "," & CheckStr(adoRecordset.Fields(5).Value) & ",'" & strUserNum & "') "
            Else
                cnnConnection.Execute "insert into r010013_2 (R055001,r055002,r055003,r055004,r055005,id) values (" & TableDataItem & ",'',''," & CheckStr(adoRecordset.Fields(4).Value) & "," & CheckStr(adoRecordset.Fields(5).Value) & ",'" & strUserNum & "') "
            End If
            TableDataItem = TableDataItem + 1
            adoRecordset.MoveNext
        Loop
    End If
    '補最後的小計
    cnnConnection.Execute "insert into r010013_2 (R055001,r055002,r055003,r055004,r055005,id)  (select " & TableDataItem & ",'部門小計','',sum(nvl(r054003,0)),sum(nvl(r054004,0)),'" & strUserNum & "' from r010013_1 where id='" & strUserNum & "' and r054001='" & tmpR054001 & "' group by  " & TableDataItem & ",'部門小計' ) "
    '補最後的總計
    TableDataItem = TableDataItem + 1
    cnnConnection.Execute "insert into r010013_2 (R055001,r055002,r055003,r055004,r055005,id)  (select " & TableDataItem & ",'總計','',sum(nvl(r054003,0)),sum(nvl(r054004,0)),'" & strUserNum & "' from r010013_1 where id='" & strUserNum & "' group by  " & TableDataItem & ",'總計' ) "
    
    strSql = "select r055002,r055003,r055004,r055005,r055001 from r010013_2 where id='" & strUserNum & "' order by r055001 "
    lbl1(4).Caption = ""
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount > 0 Then
        With grdDataList1
        Do While Not adoRecordset.EOF
            .TextMatrix(.Rows - 1, 0) = "" & adoRecordset.Fields(0).Value
            .TextMatrix(.Rows - 1, 1) = "" & adoRecordset.Fields(1).Value
            'add by nick 2004/10/08
            .TextMatrix(.Rows - 1, 2) = "" & adoRecordset.Fields(2).Value
            .TextMatrix(.Rows - 1, 3) = "" & adoRecordset.Fields(3).Value
            lbl1(4) = Val(lbl1(4)) + Val(.TextMatrix(.Rows - 1, 1))
            adoRecordset.MoveNext
            .Rows = .Rows + 1
        Loop
        .Rows = .Rows - 1
        End With
        
        'add by nickc 2008/03/12
        If m_rs.State = 1 Then m_rs.Close
        m_rs.CursorLocation = adUseClient
        m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
        Set grd2.Recordset = m_rs
        SetDataListWidth2
    Else
        ShowNoData
        Screen.MousePointer = vbDefault
        DoTemp = False
        Exit Function
    End If
    CheckOC
    DoTemp = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frm010013_2 = Nothing
End Sub

Private Sub SetDataListWidth2()
    grd2.Cols = 3
    grd2.row = 0
    grd2.col = 0: grd2.Text = "系統別"
    grd2.ColWidth(0) = 700
    grd2.CellAlignment = flexAlignCenterCenter
    grd2.col = 1: grd2.Text = "接洽單筆數"
    grd2.ColWidth(1) = 700
    grd2.CellAlignment = flexAlignCenterCenter
    grd2.col = 2: grd2.Text = "國外來函請款筆數"
    grd2.ColWidth(2) = 700
    grd2.CellAlignment = flexAlignCenterCenter
End Sub

