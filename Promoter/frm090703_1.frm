VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090703_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員每日分案情形查詢"
   ClientHeight    =   8880
   ClientLeft      =   -2445
   ClientTop       =   1320
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
      Left            =   13320
      TabIndex        =   0
      Top             =   72
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   8100
      Left            =   0
      TabIndex        =   1
      Top             =   645
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   14288
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorSel    =   16777088
      ScrollTrack     =   -1  'True
      FocusRect       =   0
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
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   72
      TabIndex        =   2
      Top             =   396
      Width           =   2676
   End
End
Attribute VB_Name = "frm090703_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/07 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim i As Integer

Private Sub cmdOK_Click()
frm090703.Show
frm090703.Enabled = True
Screen.MousePointer = vbDefault
Unload Me
End Sub

Private Sub Form_Load()
lbl1.Caption = frm090703.txt1(5) & " 年 " & frm090703.txt1(6) & " 月 "
MoveFormToCenter Me
Process
SetGrd1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090703_1 = Nothing
End Sub

Sub Process()
'Modify By Cheng 2003/07/17
'strSQL = "SELECT R105001," & SQLSum("R105002") & "," & SQLSum("R105003") & "," & SQLSum("R105004") & "," & SQLSum("R105005") & "," & SQLSum("R105006") & "," & SQLSum("R105007") & "," & SQLSum("R105008") & "," & SQLSum("R105009") & "," & SQLSum("R105010") & "," & SQLSum("R105011") & "," & SQLSum("R105012") & "," & SQLSum("R105013") & "," & SQLSum("R105014") & "," & SQLSum("R105015") & "," & SQLSum("R105016") & "," & SQLSum("R105017") & "," & SQLSum("R105018") & "," & SQLSum("R105019") & "," & SQLSum("R105020") & "," & SQLSum("R105021") & "," & SQLSum("R105022") & "," & SQLSum("R105023") & "," & SQLSum("R105024") & "," & SQLSum("R105025") & "," & SQLSum("R105026") & "," & SQLSum("R105027") & "," & SQLSum("R105028") & "," & SQLSum("R105029") & "," & SQLSum("R105030") & "," & SQLSum("R105031") & "," & SQLSum("R105032") & " FROM R090703 WHERE ID='" & strUserNum & "' GROUP BY R105001"
'Modify By Cheng 2004/02/18
'加合計欄
'strSQL = "SELECT ST02," & SQLSum("R105002") & "," & SQLSum("R105003") & "," & SQLSum("R105004") & "," & SQLSum("R105005") & "," & SQLSum("R105006") & "," & SQLSum("R105007") & "," & SQLSum("R105008") & "," & SQLSum("R105009") & "," & SQLSum("R105010") & "," & SQLSum("R105011") & "," & SQLSum("R105012") & "," & SQLSum("R105013") & "," & SQLSum("R105014") & "," & SQLSum("R105015") & "," & SQLSum("R105016") & "," & SQLSum("R105017") & "," & SQLSum("R105018") & "," & SQLSum("R105019") & "," & SQLSum("R105020") & "," & SQLSum("R105021") & "," & SQLSum("R105022") & "," & SQLSum("R105023") & "," & SQLSum("R105024") & "," & SQLSum("R105025") & "," & SQLSum("R105026") & "," & SQLSum("R105027") & "," & SQLSum("R105028") & "," & SQLSum("R105029") & "," & SQLSum("R105030") & "," & SQLSum("R105031") & "," & SQLSum("R105032") & ", R105001, ST06 " & _
'                " FROM R090703, Staff WHERE R105001=ST01(+) And ID='" & strUserNum & "' GROUP BY R105001, ST02, ST06 Order By ST06, R105001 "
strSql = "SELECT ST02, Sum(Nvl(R105002,0)) + Sum(Nvl(R105003,0)) + Sum(Nvl(R105004,0)) + Sum(Nvl(R105005,0)) + Sum(Nvl(R105006,0)) + Sum(Nvl(R105007,0)) + Sum(Nvl(R105008,0)) + Sum(Nvl(R105009,0)) + Sum(Nvl(R105010,0)) + Sum(Nvl(R105011,0)) + Sum(Nvl(R105012,0)) + Sum(Nvl(R105013,0)) + Sum(Nvl(R105014,0)) + Sum(Nvl(R105015,0)) + Sum(Nvl(R105016,0)) + Sum(Nvl(R105017,0)) + Sum(Nvl(R105018,0)) + Sum(Nvl(R105019,0)) + Sum(Nvl(R105020,0)) + Sum(Nvl(R105021,0)) + Sum(Nvl(R105022,0)) + Sum(Nvl(R105023,0)) + Sum(Nvl(R105024,0)) + Sum(Nvl(R105025,0)) + Sum(Nvl(R105026,0)) + Sum(Nvl(R105027,0)) + Sum(Nvl(R105028,0)) + Sum(Nvl(R105029,0)) + Sum(Nvl(R105030,0)) + Sum(Nvl(R105031,0)) + Sum(Nvl(R105032,0)), " & _
                SQLSum("R105002") & "," & SQLSum("R105003") & "," & SQLSum("R105004") & "," & SQLSum("R105005") & "," & SQLSum("R105006") & "," & SQLSum("R105007") & "," & SQLSum("R105008") & "," & SQLSum("R105009") & "," & SQLSum("R105010") & "," & SQLSum("R105011") & "," & SQLSum("R105012") & "," & SQLSum("R105013") & "," & SQLSum("R105014") & "," & SQLSum("R105015") & "," & SQLSum("R105016") & "," & SQLSum("R105017") & "," & SQLSum("R105018") & "," & SQLSum("R105019") & "," & SQLSum("R105020") & "," & SQLSum("R105021") & "," & SQLSum("R105022") & "," & SQLSum("R105023") & "," & SQLSum("R105024") & "," & SQLSum("R105025") & "," & SQLSum("R105026") & "," & SQLSum("R105027") & "," & SQLSum("R105028") & "," & SQLSum("R105029") & "," & SQLSum("R105030") & "," & SQLSum("R105031") & "," & SQLSum("R105032") & ", R105001, ST06 " & _
                " FROM R090703, Staff WHERE R105001=ST01(+) And ID='" & strUserNum & "' GROUP BY R105001, ST02, ST06 Order By ST06, R105001 "
'End
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/20
        Set grd1.Recordset = adoRecordset
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/20
    End If
End With
CheckOC
End Sub

Private Sub SetGrd1()
With grd1
    'Modify By Cheng 2004/02/18
'    .Cols = 32
    .Cols = 33
    'End
    .row = 0: .col = 0: .Text = "繪圖人員"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .row = 0: .col = 1: .Text = "合計"
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    For i = 2 To Me.grd1.Cols - 1
        .col = i: .Text = i - 1
        .ColWidth(i) = 420
        .CellAlignment = flexAlignCenterCenter
    Next i
End With
End Sub
