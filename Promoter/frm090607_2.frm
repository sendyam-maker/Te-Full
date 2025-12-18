VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090607_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收文高低標查詢(合併)-新申請案"
   ClientHeight    =   5730
   ClientLeft      =   1845
   ClientTop       =   2235
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7944
      TabIndex        =   6
      Top             =   72
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "高標"
      Height          =   300
      Index           =   0
      Left            =   910
      TabIndex        =   5
      Top             =   864
      Width           =   2600
   End
   Begin VB.CommandButton cmd 
      Caption         =   "介於"
      Height          =   300
      Index           =   1
      Left            =   3510
      TabIndex        =   4
      Top             =   864
      Width           =   2600
   End
   Begin VB.CommandButton cmd 
      Caption         =   "低標"
      Height          =   300
      Index           =   2
      Left            =   6110
      TabIndex        =   3
      Top             =   864
      Width           =   2600
   End
   Begin VB.CommandButton cmd 
      Caption         =   "低標"
      Height          =   300
      Index           =   5
      Left            =   6410
      TabIndex        =   2
      Top             =   3384
      Width           =   2760
   End
   Begin VB.CommandButton cmd 
      Caption         =   "介於"
      Height          =   300
      Index           =   4
      Left            =   3650
      TabIndex        =   1
      Top             =   3384
      Width           =   2760
   End
   Begin VB.CommandButton cmd 
      Caption         =   "高標"
      Height          =   300
      Index           =   3
      Left            =   890
      TabIndex        =   0
      Top             =   3384
      Width           =   2760
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2184
      Left            =   60
      TabIndex        =   7
      Top             =   1152
      Width           =   9216
      _ExtentX        =   16245
      _ExtentY        =   3863
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      MergeCells      =   1
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
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   2040
      Left            =   60
      TabIndex        =   8
      Top             =   3660
      Width           =   9216
      _ExtentX        =   16245
      _ExtentY        =   3598
      _Version        =   393216
      ScrollTrack     =   -1  'True
      HighLight       =   2
      MergeCells      =   1
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
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      Caption         =   "年度："
      Height          =   180
      Index           =   0
      Left            =   168
      TabIndex        =   12
      Top             =   600
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "月份："
      Height          =   180
      Index           =   1
      Left            =   2076
      TabIndex        =   11
      Top             =   600
      Width           =   660
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   600
      Width           =   1032
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   2844
      TabIndex        =   9
      Top             =   600
      Width           =   996
   End
End
Attribute VB_Name = "frm090607_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grd1改字型=新細明體-ExtB、grd2改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'Modified by Morgan 2013/4/19 原105條件要再加125
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Double, SavDay1 As String, SavDay2 As String, q As Integer
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, k As Integer, SeekA As Integer, SeekB As Integer, SeekC As Integer
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, SeekD As Integer

Private Sub SetGrd1()
With grd1
    .Cols = 16
    .row = 0
    .col = 0:   .Text = "智權人員"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    For i = 0 To 2
        .col = 1 + (i * 5): .Text = "發明"
        .ColWidth(1 + (i * 5)) = 500
        .CellAlignment = flexAlignCenterCenter
        .col = 2 + (i * 5):  .Text = "新型"
        .ColWidth(2 + (i * 5)) = 500
        .CellAlignment = flexAlignCenterCenter
        .col = 3 + (i * 5):  .Text = "設計"
        .ColWidth(3 + (i * 5)) = 600
        .CellAlignment = flexAlignCenterCenter
        .col = 4 + (i * 5):  .Text = "再審"
        .ColWidth(4 + (i * 5)) = 500
        .CellAlignment = flexAlignCenterCenter
        .col = 5 + (i * 5):  .Text = "小計"
        .ColWidth(5 + (i * 5)) = 500
        .CellAlignment = flexAlignCenterCenter
    Next i
End With
End Sub

Private Sub SetGrd2()
With grd2
    .Cols = 13
    .Rows = 7
    .row = 0
    .col = 0:   .Text = ""
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    For i = 0 To 2
        .col = 1 + (i * 4): .Text = "發明"
        .ColWidth(1 + (i * 4)) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 2 + (i * 4):  .Text = "新型"
        .ColWidth(2 + (i * 4)) = 750
        .CellAlignment = flexAlignCenterCenter
        .col = 3 + (i * 4):  .Text = "設計"
        .ColWidth(3 + (i * 4)) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 4 + (i * 4):  .Text = "再審"
        .ColWidth(4 + (i * 4)) = 600
        .CellAlignment = flexAlignCenterCenter
    Next i
    .col = 0
    .row = 1
    .Text = "合計件"
    .CellAlignment = flexAlignCenterCenter
    .row = 2
    .Text = "合計點"
    .CellAlignment = flexAlignCenterCenter
    .row = 3
    .Text = "平均點"
    .CellAlignment = flexAlignCenterCenter
    .row = 4
    .Text = "標準價"
    .CellAlignment = flexAlignCenterCenter
    .row = 5
    .Text = "差距"
    .CellAlignment = flexAlignCenterCenter
    .row = 6
    .Text = "分析"
    .CellAlignment = flexAlignCenterCenter
    .col = 1
    .row = 6
    .Text = "介於差"
    .CellAlignment = flexAlignCenterCenter
    .col = 4
    .row = 6
    .Text = "綜合"
    .CellAlignment = flexAlignCenterCenter
    .col = 7
    .row = 6
    .Text = "平均差"
    .CellAlignment = flexAlignCenterCenter
End With
End Sub


Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
Process
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetGrd1
SetGrd2
lbl1(0).Caption = (frm090607.txt1(3) \ 100) & "-" & (frm090607.txt1(4) \ 100)
lbl1(1).Caption = Right(frm090607.txt1(3), 2) & "-" & Right(frm090607.txt1(4), 2)
'StrMenu
Process
End Sub

'Sub StrMenu()
'strSQL = "SELECT DISTINCT R099001 FROM R090607_1 WHERE ID='" & strUserNum & "' "
'CheckOC
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        K = 0
'        Do While .EOF = False
'            Combo1.AddItem CheckStr(.Fields(0)), K
'            K = K + 1
'            .MoveNext
'        Loop
'    End If
'E 'nd With
'CheckOC
'COMBO1.TEXT = Combo1.List(0)
'End Sub

Sub Process()
grd1.Visible = False
grd2.Visible = False
'911030 nick 薛說改為件數
'***** start
'Modify By Cheng 2003/07/07
'strSQL = "SELECT R099002,SUM(R099003),SUM(R099004),SUM(R099005),SUM(R099006),SUM(R099007),SUM(R099008),SUM(R099009),SUM(R099010),SUM(R099011),SUM(R099012),SUM(R099013),SUM(R099014),SUM(R099015),SUM(R099016),SUM(R099017) FROM R090607_1 WHERE ID='" & strUserNum & "' GROUP BY R099002 "
strSql = "SELECT R099002,count(decode(R099003,0,null,r099003)),count(decode(R099004,0,null,r099004)),count(decode(R099005,0,null,r099005)),count(decode(R099006,0,null,r099006)),count(decode(R099007,0,null,r099007)),count(decode(R099008,0,null,r099008)),count(decode(R099009,0,null,r099009)),count(decode(R099010,0,null,r099010))," & _
         " count(decode(R099011,0,null,r099011)),count(decode(R099012,0,null,r099012)),count(decode(R099013,0,null,r099013)),count(decode(R099014,0,null,r099014)),count(decode(R099015,0,null,r099015)),count(decode(R099016,0,null,r099016)),count(decode(R099017,0,null,r099017)) FROM R090607_1 WHERE ID='" & strUserNum & "' GROUP BY R099002 "
'2010/7/14 還原 BY SONIA不計件也要計算
'strSql = "SELECT R099002,count(decode(R099003,0,null,r099003)),count(decode(R099004,0,null,r099004)),count(decode(R099005,0,null,r099005)),count(decode(R099006,0,null,r099006)),count(decode(R099007,0,null,r099007)),count(decode(R099008,0,null,r099008)),count(decode(R099009,0,null,r099009)),count(decode(R099010,0,null,r099010))," & _
         " count(decode(R099011,0,null,r099011)),count(decode(R099012,0,null,r099012)),count(decode(R099013,0,null,r099013)),count(decode(R099014,0,null,r099014)),count(decode(R099015,0,null,r099015)),count(decode(R099016,0,null,r099016)),count(decode(R099017,0,null,r099017)) FROM R090607_1 WHERE ID='" & strUserNum & "' And R099019 Is Null GROUP BY R099002 "
'***** end
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset
    Else
        Set grd1.Recordset = adoRecordset
    End If
End With
CheckOC
SetGrd1
grd2.Clear
SetGrd2
grd2.row = 1
'Modify By Cheng 2003/07/07
strSql = "SELECT COUNT(R099003) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099003 IS NOT NULL AND R099003<> 0 "
'2010/7/14 還原 BY SONIA不計件也要計算
'strSql = "SELECT COUNT(R099003) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099003 IS NOT NULL AND R099003<> 0 And R099019 Is Null "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 1
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 1
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099004) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099004 IS NOT NULL AND R099004<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099004) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099004 IS NOT NULL AND R099004<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 2
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 2
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099005) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099005 IS NOT NULL AND R099005<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099005) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099005 IS NOT NULL AND R099005<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 3
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 3
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099006) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099006 IS NOT NULL AND R099006<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099006) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099006 IS NOT NULL AND R099006<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 4
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 4
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099008) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099008 IS NOT NULL AND R099008<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099008) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099008 IS NOT NULL AND R099008<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 5
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 5
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099009) fROM R090607_1 WHERE ID='" & strUserNum & "' AND R099009 IS NOT NULL AND R099009<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099009) fROM R090607_1 WHERE ID='" & strUserNum & "' AND R099009 IS NOT NULL AND R099009<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 6
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 6
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099010) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099010 IS NOT NULL AND R099010<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099010) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099010 IS NOT NULL AND R099010<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 7
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 7
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099011) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099011 IS NOT NULL AND R099011<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099011) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099011 IS NOT NULL AND R099011<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 8
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 8
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099013) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099013 IS NOT NULL AND R099013<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099013) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099013 IS NOT NULL AND R099013<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 9
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 9
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099014) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099014 IS NOT NULL AND R099014<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099014) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099014 IS NOT NULL AND R099014<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 10
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 10
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099015) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099015 IS NOT NULL AND R099015<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099015) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099015 IS NOT NULL AND R099015<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 11
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 11
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
    'Modify By Cheng 2003/07/07
    strSql = "SELECT COUNT(R099016) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099016 IS NOT NULL AND R099016<> 0 "
   '2010/7/14 還原 BY SONIA不計件也要計算
    'strSql = "SELECT COUNT(R099016) FROM R090607_1 WHERE ID='" & strUserNum & "' AND R099016 IS NOT NULL AND R099016<> 0 And R099019 Is Null "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.row = 1
        grd2.col = 12
        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
    Else
        grd2.col = 12
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
End With
'Modify By Cheng 2003/07/14
'strSQL = "select CF13 FROM CASEFEE WHERE CF01='" & frm090607.Txt1(0) & "' AND CF02='" & frm090607.Txt1(1) & "' AND SUBSTR(CF03,1,3)='101' "
'CheckOC
'grd2.Row = 4
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        grd2.col = 1
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 5
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 9
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'    Else
'        grd2.col = 1
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 5
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 9
'        grd2.Text = Format(0, "#########0.00")
'    End If
'    CheckOC
'    strSQL = "select CF13 FROM CASEFEE WHERE CF01='" & frm090607.Txt1(0) & "' AND CF02='" & frm090607.Txt1(1) & "' AND SUBSTR(CF03,1,3)='102' "
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        grd2.col = 2
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 6
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 10
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'    Else
'        grd2.col = 2
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 6
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 10
'        grd2.Text = Format(0, "#########0.00")
'    End If
'    CheckOC
'    strSQL = "select CF13 FROM CASEFEE WHERE CF01='" & frm090607.Txt1(0) & "' AND CF02='" & frm090607.Txt1(1) & "' AND SUBSTR(CF03,1,3)='103' "
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        grd2.col = 3
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 7
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 11
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'    Else
'        grd2.col = 3
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 7
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 11
'        grd2.Text = Format(0, "#########0.00")
'    End If
'    CheckOC
'    strSQL = "select CF13 FROM CASEFEE WHERE CF01='" & frm090607.Txt1(0) & "' AND CF02='" & frm090607.Txt1(1) & "' AND SUBSTR(CF03,1,3)='107' "
'    .CursorLocation = adUseClient
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        grd2.col = 4
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 8
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'        grd2.col = 12
'        grd2.Text = Format(Val(CheckStr(.Fields(0))), "#########0.00")
'    Else
'        grd2.col = 4
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 8
'        grd2.Text = Format(0, "#########0.00")
'        grd2.col = 12
'        grd2.Text = Format(0, "#########0.00")
'    End If
'    CheckOC
'End With
'標準價
strSql = "Select Sum(Decode(R099003,0,0,Nvl(CP33,0))), Sum(Decode(R099003,0,0,Nvl(1,0))), Sum(Decode(R099008,0,0,Nvl(CP33,0))), Sum(Decode(R099008,0,0,Nvl(1,0))), Sum(Decode(R099013,0,0,Nvl(CP33,0))), Sum(Decode(R099013,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE ID='" & strUserNum & "' AND R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='101' Or (CP10='104' And PA08='1' )) "
CheckOC
grd2.row = 4
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.col = 1
        If Val("" & .Fields(1).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(0))) / Val(.Fields(1).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 5
        If Val("" & .Fields(3).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(2))) / Val(.Fields(3).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 9
        If Val("" & .Fields(5).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(4))) / Val(.Fields(5).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
    Else
        grd2.col = 1
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 5
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 9
        grd2.Text = Format(0, "#########0.00")
    End If
End With
CheckOC
strSql = "Select Sum(Decode(R099004,0,0,Nvl(CP33,0))), Sum(Decode(R099004,0,0,Nvl(1,0))), Sum(Decode(R099009,0,0,Nvl(CP33,0))), Sum(Decode(R099009,0,0,Nvl(1,0))), Sum(Decode(R099014,0,0,Nvl(CP33,0))), Sum(Decode(R099014,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE ID='" & strUserNum & "' AND R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='102' Or (CP10='104' And PA08='2' )) "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.col = 2
        If Val("" & .Fields(1).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(0))) / Val(.Fields(1).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 6
        If Val("" & .Fields(3).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(2))) / Val(.Fields(3).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 10
        If Val("" & .Fields(5).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(4))) / Val(.Fields(5).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
    Else
        grd2.col = 2
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 6
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 10
        grd2.Text = Format(0, "#########0.00")
    End If
End With
CheckOC
strSql = "Select Sum(Decode(R099005,0,0,Nvl(CP33,0))), Sum(Decode(R099005,0,0,Nvl(1,0))), Sum(Decode(R099010,0,0,Nvl(CP33,0))), Sum(Decode(R099010,0,0,Nvl(1,0))), Sum(Decode(R099015,0,0,Nvl(CP33,0))), Sum(Decode(R099015,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE ID='" & strUserNum & "' AND R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='103' Or CP10='105' Or CP10='125') "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.col = 3
        If Val("" & .Fields(1).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(0))) / Val(.Fields(1).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 7
        If Val("" & .Fields(3).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(2))) / Val(.Fields(3).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 11
        If Val("" & .Fields(5).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(4))) / Val(.Fields(5).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
    Else
        grd2.col = 3
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 7
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 11
        grd2.Text = Format(0, "#########0.00")
    End If
End With
CheckOC
strSql = "Select Sum(Decode(R099006,0,0,Nvl(CP33,0))), Sum(Decode(R099006,0,0,Nvl(1,0))), Sum(Decode(R099011,0,0,Nvl(CP33,0))), Sum(Decode(R099011,0,0,Nvl(1,0))), Sum(Decode(R099016,0,0,Nvl(CP33,0))), Sum(Decode(R099016,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE ID='" & strUserNum & "' AND R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='107' ) "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        grd2.col = 4
        If Val("" & .Fields(1).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(0))) / Val(.Fields(1).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 8
        If Val("" & .Fields(3).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(2))) / Val(.Fields(3).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
        grd2.col = 12
        If Val("" & .Fields(5).Value) <> 0 Then
            grd2.Text = Format(Val(CheckStr(.Fields(4))) / Val(.Fields(5).Value), "#########0.00")
        Else
            grd2.Text = Format(0, "#########0.00")
        End If
    Else
        grd2.col = 4
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 8
        grd2.Text = Format(0, "#########0.00")
        grd2.col = 12
        grd2.Text = Format(0, "#########0.00")
    End If
    CheckOC
End With
strSql = "SELECT SUM(R099003),SUM(R099004),SUM(R099005),SUM(R099006),SUM(R099008),SUM(R099009),SUM(R099010),SUM(R099011),SUM(R099013),SUM(R099014),SUM(R099015),SUM(R099016) FROM R090607_1 WHERE ID='" & strUserNum & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            grd2.row = 2
            For i = 0 To 11
                grd2.col = i + 1
                grd2.Text = Format(Val(CheckStr(.Fields(i))), "#########0.00")
            Next i
            .MoveNext
        Loop
    End If
End With
CheckOC
k = 0
s = 0
With grd2
    For i = 1 To 12
        .row = 1
        .col = i
        If .Text = 0 Then
            .row = 3
            .col = i
            .Text = Format(0, "#########0.00")
        Else
            k = Val(.Text)
            .row = 2
            s = Val(.Text)
            .row = 3
            .Text = Format(s / k, "#########0.00")
        End If
        .row = 3
        q = Val(.Text)
        .row = 4
        s = Val(.Text)
        .row = 5
        .Text = Format(q - s, "#########0.00")
    Next i
    k = 0
    s = 0
    SeekA = 0
    SeekB = 0
    SeekC = 0
    For i = 1 To 4
        .col = i
        .row = 1
        k = Val(.Text)
        .row = 5
        s = Val(.Text)
        SeekA = SeekA + (k * s)
    Next i
    For i = 5 To 8
        .col = i
        .row = 1
        k = Val(.Text)
        .row = 5
        s = Val(.Text)
        SeekB = SeekB + (k * s)
    Next i
    For i = 9 To 12
        .col = i
        .row = 1
        k = Val(.Text)
        .row = 5
        s = Val(.Text)
        SeekC = SeekC + (k * s)
    Next i
    .row = 6
    .col = 2
    .Text = Format(SeekB, "#########0.00")
    .col = 5
    .Text = Format(SeekA + SeekB + SeekC, "#########0.00")
    SeekD = 0
    For i = 1 To 12
        .row = 1
        .col = i
        SeekD = SeekD + Val(.Text)
    Next i
    If SeekD <> 0 Then
        .row = 6
        .col = 8
        .Text = Format((SeekA + SeekB + SeekC) / SeekD, "#########0.00")
    Else
        .row = 6
        .col = 8
        .Text = Format(0, "#########0.00")
    End If
End With
grd1.Visible = True
grd2.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm090607.Show
   Set frm090607_2 = Nothing
End Sub

