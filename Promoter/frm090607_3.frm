VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090607_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收文高低標查詢(各區)-爭議/救濟案"
   ClientHeight    =   5730
   ClientLeft      =   1440
   ClientTop       =   2310
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4848
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   564
      Width           =   4380
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7824
      TabIndex        =   6
      Top             =   36
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
      Left            =   6120
      TabIndex        =   2
      Top             =   3384
      Width           =   2610
   End
   Begin VB.CommandButton cmd 
      Caption         =   "介於"
      Height          =   300
      Index           =   4
      Left            =   3500
      TabIndex        =   1
      Top             =   3384
      Width           =   2610
   End
   Begin VB.CommandButton cmd 
      Caption         =   "高標"
      Height          =   300
      Index           =   3
      Left            =   880
      TabIndex        =   0
      Top             =   3384
      Width           =   2610
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2184
      Left            =   96
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
      TabIndex        =   9
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
      Left            =   264
      TabIndex        =   14
      Top             =   588
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "月份："
      Height          =   180
      Index           =   1
      Left            =   2172
      TabIndex        =   13
      Top             =   588
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別："
      Height          =   180
      Index           =   2
      Left            =   3936
      TabIndex        =   12
      Top             =   588
      Width           =   912
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   972
      TabIndex        =   11
      Top             =   588
      Width           =   780
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   2952
      TabIndex        =   10
      Top             =   588
      Width           =   780
   End
End
Attribute VB_Name = "frm090607_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grd1改字型=新細明體-ExtB、grd2改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Double, SavDay1 As String, SavDay2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, k As Integer, SeekA As Single, SeekB As Single, SeekC As Single
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, SeekD As Integer, q As Single
Dim kk As Single, ss As Single, qq As Single

Private Sub SetGrd1()
With grd1
    .Cols = 13
    .row = 0
    .col = 0:   .Text = "智權人員"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    For i = 0 To 2
        .col = 1 + (i * 4): .Text = "爭議"
        .ColWidth(1 + (i * 4)) = 650
        .CellAlignment = flexAlignCenterCenter
        .col = 2 + (i * 4):  .Text = "救濟"
        .ColWidth(2 + (i * 4)) = 650
        .CellAlignment = flexAlignCenterCenter
        .col = 3 + (i * 4):  .Text = "其他"
        .ColWidth(3 + (i * 4)) = 650
        .CellAlignment = flexAlignCenterCenter
        .col = 4 + (i * 4):  .Text = "小計"
        .ColWidth(4 + (i * 4)) = 650
        .CellAlignment = flexAlignCenterCenter
    Next i
End With
End Sub

Private Sub SetGrd2()
With grd2
    .Cols = 10
    .Rows = 7
    .row = 0
    .col = 0:   .Text = ""
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    For i = 0 To 2
        .col = 1 + (i * 3): .Text = "爭議"
        .ColWidth(1 + (i * 3)) = 866
        .CellAlignment = flexAlignCenterCenter
        .col = 2 + (i * 3):  .Text = "救濟"
        .ColWidth(2 + (i * 3)) = 866
        .CellAlignment = flexAlignCenterCenter
        .col = 3 + (i * 3):  .Text = "其他"
        .ColWidth(3 + (i * 3)) = 868
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
StrMenu
Process
End Sub

Sub StrMenu()
strSql = "SELECT DISTINCT R100001,nvl(a0902,A0903) FROM R090607_2,acc090 WHERE R100001=a0901(+) and ID='" & strUserNum & "' order by R100001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        k = 0
        Do While .EOF = False
            Combo1.AddItem CheckStr(.Fields(0)) & "==>" & CheckStr(.Fields(1).Value), k
            k = k + 1
            .MoveNext
        Loop
    End If
End With
CheckOC
Combo1.Text = Combo1.List(0)
End Sub

Sub Process()
grd1.Visible = False
grd2.Visible = False
'911030 nick 薛說改為件數
'***** start
'strSQL = "SELECT R100002,SUM(R100003),SUM(R100004),SUM(R100005),SUM(R100006),SUM(R100007),SUM(R100008),SUM(R100009),SUM(R100010),SUM(R100011),SUM(R100012),SUM(R100013),SUM(R100014) FROM R090607_2 WHERE ID='" & strUserNum & "' AND R100001='" & Mid(Combo1.Text, 1, 3) & "' GROUP BY R100002 "
strSql = "SELECT R100002,count(decode(R100003,0,null,R100003)),count(decode(R100004,0,null,R100004)),count(decode(R100005,0,null,R100005)),count(decode(R100006,0,null,R100006)),count(decode(R100007,0,null,R100007)),count(decode(R100008,0,null,R100008)),count(decode(R100009,0,null,R100009)),count(decode(R100010,0,null,R100010))," & _
         " count(decode(R100011,0,null,R100011)),count(decode(R100012,0,null,R100012)),count(decode(R100013,0,null,R100013)),count(decode(R100014,0,null,R100014)) FROM R090607_2 WHERE ID='" & strUserNum & "' AND R100001='" & Mid(Combo1.Text, 1, 3) & "' GROUP BY R100002 "
'***** end


CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset
    End If
End With
CheckOC
SetGrd1
grd2.Clear
SetGrd2
grd2.row = 1
strSql = "SELECT COUNT(R100003) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100003 IS NOT NULL AND r100003<> 0 "
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
    strSql = "SELECT COUNT(r100004) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100004 IS NOT NULL AND r100004<> 0 "
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
    strSql = "SELECT COUNT(r100005) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100005 IS NOT NULL AND r100005<> 0 "
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
    strSql = "SELECT COUNT(r100007) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100007 IS NOT NULL AND r100007<> 0 "
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
    strSql = "SELECT COUNT(r100008) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100008 IS NOT NULL AND r100008<> 0 "
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
    strSql = "SELECT COUNT(r100009) fROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100009 IS NOT NULL AND r100009<> 0 "
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
    strSql = "SELECT COUNT(r100011) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100011 IS NOT NULL AND r100011<> 0 "
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
    strSql = "SELECT COUNT(r100012) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100012 IS NOT NULL AND r100012<> 0 "
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
    strSql = "SELECT COUNT(r100013) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' AND r100013 IS NOT NULL AND r100013<> 0 "
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
End With
grd2.row = 4
grd2.col = 1
grd2.Text = Format(0, "#########0.00")
grd2.col = 4
grd2.Text = Format(0, "#########0.00")
grd2.col = 7
grd2.Text = Format(0, "#########0.00")
grd2.col = 2
grd2.Text = Format(0, "#########0.00")
grd2.col = 5
grd2.Text = Format(0, "#########0.00")
grd2.col = 8
grd2.Text = Format(0, "#########0.00")
grd2.col = 3
grd2.Text = Format(0, "#########0.00")
grd2.col = 6
grd2.Text = Format(0, "#########0.00")
grd2.col = 9
grd2.Text = Format(0, "#########0.00")

strSql = "SELECT SUM(r100003),SUM(r100004),SUM(r100005),SUM(r100007),SUM(r100008),SUM(r100009),SUM(r100011),SUM(r100012),SUM(r100013) FROM R090607_2 WHERE ID='" & strUserNum & "' AND r100001='" & Mid(Combo1.Text, 1, 3) & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            grd2.row = 2
            For i = 0 To 8
                grd2.col = i + 1
                grd2.Text = Format(Val(CheckStr(.Fields(i))), "#########0.00")
            Next i
            .MoveNext
        Loop
    End If
End With
CheckOC
kk = 0
ss = 0
With grd2
    For i = 1 To 9
        .row = 1
        .col = i
        If .Text = 0 Then
            .row = 3
            .col = i
            .Text = Format(0, "#########0.00")
        Else
            kk = Val(.Text)
            .row = 2
            ss = Val(.Text)
            .row = 3
            .Text = Format(ss / kk, "#########0.00")
        End If
        .row = 3
        qq = .Text
        .row = 4
        ss = Val(.Text)
        .row = 5
        .Text = Format(qq - ss, "#########0.00")
    Next i
    kk = 0
    ss = 0
    SeekA = 0
    SeekB = 0
    SeekC = 0
    For i = 1 To 2
        .col = i
        .row = 1
        kk = Val(.Text)
        .row = 5
        ss = Val(.Text)
        SeekA = SeekA + (kk * ss)
    Next i
    For i = 4 To 5
        .col = i
        .row = 1
        kk = Val(.Text)
        .row = 5
        ss = Val(.Text)
        SeekB = SeekB + (kk * ss)
    Next i
    For i = 7 To 8
        .col = i
        .row = 1
        kk = Val(.Text)
        .row = 5
        ss = Val(.Text)
        SeekC = SeekC + (kk * ss)
    Next i
    .row = 6
    .col = 2
    .Text = Format(SeekB, "#########0.00")
    .col = 5
    .Text = Format(SeekA + SeekB + SeekC, "#########0.00")
    SeekD = 0
    For i = 1 To 2
        .row = 1
        .col = i
        SeekD = SeekD + Val(.Text)
    Next i
    For i = 4 To 5
        .row = 1
        .col = i
        SeekD = SeekD + Val(.Text)
    Next i
    For i = 7 To 8
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
   Set frm090607_3 = Nothing
End Sub

