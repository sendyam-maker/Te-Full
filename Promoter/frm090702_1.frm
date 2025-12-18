VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090702_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員工作量查詢"
   ClientHeight    =   5715
   ClientLeft      =   -2070
   ClientTop       =   1515
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "逾本所期限(&T)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6504
      TabIndex        =   1
      Top             =   24
      Width           =   1500
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8028
      TabIndex        =   0
      Top             =   24
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4620
      Left            =   30
      TabIndex        =   2
      Top             =   750
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8149
      _Version        =   393216
      Rows            =   3
      Cols            =   1
      FixedRows       =   2
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "備註: 括弧內為舊制算法"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   7290
      TabIndex        =   5
      Top             =   5430
      Width           =   1890
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   396
      TabIndex        =   4
      Top             =   468
      Width           =   1776
   End
   Begin VB.Label lbl2 
      Height          =   180
      Left            =   75
      TabIndex        =   3
      Top             =   5415
      Width           =   4035
   End
End
Attribute VB_Name = "frm090702_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'Modified by Morgan 2016/4/28 改新制統計,加修改圖式欄位
Option Explicit


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Me.Hide
     frm090702_2.Show
Case 1
     Me.Hide
     If frm090702.ObjForm = 1 Then
        frm090702.Show
        Unload Me
        Exit Sub
     Else
        frm090702_2.Show
        Unload Me
        Exit Sub
     End If
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetGrd1
'Modify By Cheng 2004/03/30
'lbl1.Caption = Mid(GetTaiwanTodayDate, 1, 2) & " 年 " & Mid(GetTaiwanTodayDate, 3, 2) & " 月 "
lbl1.Caption = Left(strSrvDate(1), 4) - 1911 & " 年 " & Mid(strSrvDate(1), 5, 2) & " 月 "
'End
StrMenu
'Modify By Cheng 2004/03/05
'LBL2.Caption = "合計：" & str(grd1.Rows - 1)
lbl2.Caption = "合計：" & str(grd1.Rows - 2)
'End
'Process
End Sub

Sub StrMenu()
'Modify By Cheng 2003/07/16
'strSQL = "select r103001,sum(r103002),sum(r103003),sum(r103004),sum(r103005),sum(r103006),sum(r103007),sum(r103008),sum(r103009),sum(r103010) from r090702_1 where id='" & strUserNum & "' group by r103001 order by r103001 "
strSql = "select ST02,sum(r103002),sum(r103003),sum(r103012)||'('||sum(r103004)||')',sum(r103013)||'('||sum(r103005)||')',sum(r103014)||'('||sum(r103006)||')',sum(r103015)||'('||sum(r103007)||')',sum(r103011),sum(r103016)||'('||sum(r103008)||')',sum(r103017)||'('||sum(r103009)||')',sum(r103010), ST06, r103001 from r090702_1, Staff where R103001=ST01 And id='" & strUserNum & "' group by ST02, ST06, r103001 order by ST06, r103001 "
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
SetGrd1
End Sub

Private Sub SetGrd1()
With grd1
    .Cols = 11
    .row = 0
    .col = 0:   .Text = " "
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "逾時件數"
    .ColWidth(1) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "逾時件數"
    .ColWidth(2) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "承辦量"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "承辦量"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "可辦量"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "可辦量"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "可辦量"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = " "
    .ColWidth(8) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = " "
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = " "
    .ColWidth(10) = 800
    .CellAlignment = flexAlignCenterCenter
    
    .row = 1
    .col = 0:   .Text = "繪圖人員"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "草圖"
    .ColWidth(1) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "墨圖"
    .ColWidth(2) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "草圖"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "墨圖"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "草圖"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "墨圖"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "修改圖式(件)"
    .ColWidth(7) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "分案量"
    .ColWidth(8) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "發文量"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "發文點數"
    .ColWidth(10) = 800
    .CellAlignment = flexAlignCenterCenter

    '標題合併顯示
    .MergeCells = flexMergeRestrictRows
    .MergeRow(0) = True
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090702_1 = Nothing
End Sub

