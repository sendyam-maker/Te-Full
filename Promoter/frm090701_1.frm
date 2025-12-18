VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090701_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員達成情形"
   ClientHeight    =   8880
   ClientLeft      =   -2640
   ClientTop       =   1350
   ClientWidth     =   15000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15000
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   13380
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   8250
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   14552
      _Version        =   393216
      Rows            =   3
      Cols            =   14
      FixedRows       =   2
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm090701_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Private Sub cmdOK_Click()
Me.Hide
frm090701.Show
Unload Me
End Sub

Private Sub Combo1_Click()
Process
End Sub

Private Sub Form_Load()
'select r102001,r102002,sum(r102003),sum(r102004),sum(r102005),sum(r102006),sum(r102007),sum(r102008),sum(r102009),sum(r102010),sum(r102011),sum(r102012),sum(r102013),sum(r102014) from r090608 group by r102001,r102002
MoveFormToCenter Me
SetGrd1
'StrMenu
Process
End Sub

Sub Process()
Dim Cnt As Integer '明細筆數
Dim ii As Integer

'Modify By Cheng 2003/08/01
'strSQL = "select ST02," & SQLSum2("r102002") & "," & SQLSum2("r102003") & "," & SQLSum2("r102004") & "," & SQLSum2("r102005") & "," & _
'         SQLSum2("r102006") & "," & SQLSum2("r102007") & "," & SQLSum2("r102008") & "," & SQLSum2("r102009") & "," & SQLSum2("r102010") & "," & _
'         SQLSum2("r102011") & "," & SQLSum2("r102012") & "," & SQLSum2("r102013") & "," & SQLSum2("r102014") & "," & SQLSum2("r102015") & "," & _
'         SQLSum2("r102016") & "," & SQLSum2("r102017") & "," & SQLSum2("r102018") & ",R102001 from r090701,STAFF WHERE R102001=ST01(+) AND ID='" & strUserNum & "'  group by r102001,ST02 "
'edit by nickc 2005/03/24
'StrSql = "select ST02," & SQLSum2("r102002") & "," & SQLSum2("r102003") & "," & SQLSum2("r102004") & "," & SQLSum2("r102005") & "," & _
'         SQLSum2("r102006") & "," & SQLSum2("r102007") & "," & SQLSum2("r102008") & "," & SQLSum2("r102009") & "," & SQLSum2("r102010") & "," & _
'         SQLSum2("r102011") & "," & SQLSum2("r102012") & "," & SQLSum2("r102013") & "," & SQLSum2("r102014") & "," & SQLSum2("r102015") & "," & _
'         SQLSum2("r102016") & "," & SQLSum2("r102017") & "," & SQLSum2("r102018") & "," & SQLSum2("r102019") & "," & SQLSum2("r102020") & ",R102001, ST06 from r090701,STAFF WHERE R102001<>'99999' AND R102001=ST01(+) AND ID='" & strUserNum & "'  group by r102001,ST02, ST06 "
strSql = "select " & SQLSum2("r102002") & "," & SQLSum2("r102003") & "," & SQLSum2("r102004") & "," & SQLSum2("r102005") & "," & _
         SQLSum2("r102006") & "," & SQLSum2("r102007") & "," & SQLSum2("r102008") & "," & SQLSum2("r102009") & "," & SQLSum2("r102010") & "," & _
         SQLSum2("r102011") & "," & SQLSum2("r102012") & "," & SQLSum2("r102013") & "," & SQLSum2("r102014") & "," & SQLSum2("r102015") & "," & _
         SQLSum2("r102016") & "," & SQLSum2("r102017") & "," & SQLSum2("r102018") & "," & SQLSum2("r102019") & "," & SQLSum2("r102020") & ",R102001, ST06,st02 from r090701,STAFF WHERE R102001<>'99999' AND R102001=ST01(+) AND ID='" & strUserNum & "'  group by r102001,ST02, ST06 "

strSql = strSql & " Having (nvl(Sum(R102002),0)+nvl(Sum(R102003),0)+nvl(Sum(R102004),0)+nvl(Sum(R102005),0)+nvl(Sum(R102006),0)+nvl(Sum(R102007),0)+nvl(Sum(R102008),0)+nvl(Sum(R102009),0)+nvl(Sum(R102010),0)+nvl(Sum(R102011),0)+nvl(Sum(R102012),0)+nvl(Sum(R102013),0)+nvl(Sum(R102014),0)+nvl(Sum(R102015),0)+nvl(Sum(R102016),0)+nvl(Sum(R102017),0)+nvl(Sum(R102018),0)+nvl(Sum(R102019),0)+nvl(Sum(R102020),0))  > 0 "
Select Case Val(frm090701.txt1(9))
Case 1
     pub_QL05 = pub_QL05 & ";" & frm090701.Label1(6) & "1.點數 %" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY " & SQLSum2("R102008") & " DESC "
Case 2
     pub_QL05 = pub_QL05 & ";" & frm090701.Label1(6) & "2.件數 %" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY " & SQLSum2("R102009") & " DESC "
Case 3
     pub_QL05 = pub_QL05 & ";" & frm090701.Label1(6) & "3.平均 %" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY " & SQLSum2("R102011") & " DESC "
Case 4
     pub_QL05 = pub_QL05 & ";" & frm090701.Label1(6) & "4.繪圖人員" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY ST06, R102001 "
Case Else
End Select
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        Cnt = .RecordCount
        Set grd1.Recordset = adoRecordset
        'add by nickc 2005/03/24
        For ii = 0 To grd1.Rows - 1
            grd1.TextMatrix(ii, 0) = grd1.TextMatrix(ii, 22)
        Next ii
        'Add By Cheng 2004/03/12
        '加合計
        CheckOC
        strSql = ""
        strSql = "select '合  計'," & SQLSum2("r102002") & "," & SQLSum2("r102003") & "," & SQLSum2("r102004") & "," & SQLSum2("r102005") & "," & _
                 SQLSum2("r102006") & "," & SQLSum2("r102007") & "," & SQLSum2("r102008") & "," & SQLSum2("r102009") & "," & SQLSum2("r102010") & "," & _
                 SQLSum2("r102011") & "," & SQLSum2("r102012") & "," & SQLSum2("r102013") & "," & SQLSum2("r102014") & "," & SQLSum2("r102015") & "," & _
                 SQLSum2("r102016") & "," & SQLSum2("r102017") & "," & SQLSum2("r102018") & "," & SQLSum2("r102019") & "," & SQLSum2("r102020") & ", '', '' from r090701,STAFF WHERE R102001<>'99999' AND R102001=ST01(+) AND ID='" & strUserNum & "'  group by '合  計' "
        strSql = strSql & " Having (nvl(Sum(R102002),0)+nvl(Sum(R102003),0)+nvl(Sum(R102004),0)+nvl(Sum(R102005),0)+nvl(Sum(R102006),0)+nvl(Sum(R102007),0)+nvl(Sum(R102008),0)+nvl(Sum(R102009),0)+nvl(Sum(R102010),0)+nvl(Sum(R102011),0)+nvl(Sum(R102012),0)+nvl(Sum(R102013),0)+nvl(Sum(R102014),0)+nvl(Sum(R102015),0)+nvl(Sum(R102016),0)+nvl(Sum(R102017),0)+nvl(Sum(R102018),0)+nvl(Sum(R102019),0)+nvl(Sum(R102020),0)) > 0 "
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            Me.grd1.AddItem "", Me.grd1.Rows
            For ii = 0 To Me.grd1.Cols - 2
                Select Case ii
                Case 7, 8, 9, 10
                    Me.grd1.TextMatrix(Me.grd1.Rows - 1, ii) = "" & Format((Val("" & .Fields(ii).Value) / Cnt), "0.00")
                Case Else
                    Me.grd1.TextMatrix(Me.grd1.Rows - 1, ii) = "" & .Fields(ii).Value
                End Select
            Next ii
        End If
        'End
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
SetGrd1
CheckOC
End Sub

Private Sub SetGrd1()
With grd1
    'Modify By Cheng 2003/08/01
'    .Cols = 18
    'edit by nickc 2005/03/24
    '.Cols = 20
    .Cols = 21
    'Add By Cheng 2003/06/05
    .row = 0
    .col = 0:   .Text = "繪圖人員"
    .ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "目標"
    .ColWidth(1) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "目標"
    .ColWidth(2) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "目標"
    .ColWidth(3) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "目標達成"
    .ColWidth(4) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "目標達成"
    .ColWidth(5) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "目標達成"
    .ColWidth(6) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "達成率%"
    .ColWidth(7) = 750
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "達成率%"
    .ColWidth(8) = 750
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "達成率%"
    .ColWidth(9) = 750
    .CellAlignment = flexAlignCenterCenter
    .col = 10:  .Text = "達成率%"
    .ColWidth(10) = 750
    .CellAlignment = flexAlignCenterCenter
    'edit by nickc 2005/04/13
    If frm090701.txt1(10) = "1" Then
         .col = 11:  .Text = "其他新案"
         .ColWidth(11) = 700
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = "其他新案"
         .ColWidth(12) = 700
         .CellAlignment = flexAlignCenterCenter
         .col = 13:  .Text = "其他舊案"
         .ColWidth(13) = 700
         .CellAlignment = flexAlignCenterCenter
         .col = 14:  .Text = "其他舊案"
         .ColWidth(14) = 700
         .CellAlignment = flexAlignCenterCenter
     Else
         .col = 11:  .Text = "提供圖檔(0.6)"
         .ColWidth(11) = 1400
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = "轉換案(0.4)"
         .ColWidth(12) = 1400
         .CellAlignment = flexAlignCenterCenter
         .col = 13:  .Text = "其他新舊案"
         .ColWidth(13) = 1300
         .CellAlignment = flexAlignCenterCenter
         .col = 14:  .Text = "其他新舊案"
         .ColWidth(14) = 0
         .CellAlignment = flexAlignCenterCenter
     End If
    .col = 15:  .Text = "完　　成"
    .ColWidth(15) = 900
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "完　　成"
'    .ColWidth(16) = 1200
    .ColWidth(16) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "完　　成"
    .ColWidth(17) = 900
    .CellAlignment = flexAlignCenterCenter
    'Add By Cheng 2003/08/01
    .col = 18:  .Text = "逾　　時"
    .ColWidth(18) = 900
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "逾　　時"
    .ColWidth(19) = 900
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = ""
    .ColWidth(20) = 0
    .CellAlignment = flexAlignCenterCenter
    
    .row = 1
    .col = 0:   .Text = "繪圖人員"
    .ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "點數"
    .ColWidth(1) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "件數"
    .ColWidth(2) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "張數"
    .ColWidth(3) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "點數"
    .ColWidth(4) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "件數"
    .ColWidth(5) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "張數"
    .ColWidth(6) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "點數"
    .ColWidth(7) = 750
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "件數"
    .ColWidth(8) = 750
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "張數"
    .ColWidth(9) = 750
    .CellAlignment = flexAlignCenterCenter
    .col = 10:  .Text = "平均"
    .ColWidth(10) = 750
    .CellAlignment = flexAlignCenterCenter
    'edit by nickc 2005/04/13
    If frm090701.txt1(10) = "1" Then
         .col = 11:  .Text = "點數"
         .ColWidth(11) = 700
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = "件數"
         .ColWidth(12) = 700
         .CellAlignment = flexAlignCenterCenter
         .col = 13:  .Text = "點數"
         .ColWidth(13) = 700
         .CellAlignment = flexAlignCenterCenter
         .col = 14:  .Text = "件數"
         .ColWidth(14) = 700
         .CellAlignment = flexAlignCenterCenter
    Else
         .col = 11:  .Text = "件數"
         .ColWidth(11) = 1400
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = "件數"
         .ColWidth(12) = 1400
         .CellAlignment = flexAlignCenterCenter
         .col = 13:  .Text = "件數"
         .ColWidth(13) = 1300
         .CellAlignment = flexAlignCenterCenter
         .col = 14:  .Text = "件數"
         .ColWidth(14) = 0
         .CellAlignment = flexAlignCenterCenter
    End If
    .col = 15:  .Text = "草圖件數"
    .ColWidth(15) = 900
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "ACAD 件數"
'    .ColWidth(16) = 1200
    .ColWidth(16) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "墨圖件數"
    .ColWidth(17) = 900
    .CellAlignment = flexAlignCenterCenter
    'Add By Cheng 2003/08/01
    .col = 18:  .Text = "草圖件數"
    .ColWidth(18) = 900
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "墨圖件數 "
    .ColWidth(19) = 900
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = ""
    .ColWidth(20) = 0
    .CellAlignment = flexAlignCenterCenter
    'Add By Cheng 2003/06/05
    '標題合併顯示
    .MergeCells = flexMergeRestrictRows
    .MergeRow(0) = True
    .MergeCol(0) = True
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090701_1 = Nothing
End Sub
