VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090610_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人每日分案情形查詢"
   ClientHeight    =   5730
   ClientLeft      =   -3330
   ClientTop       =   1005
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   8016
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4968
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   9288
      _ExtentX        =   16378
      _ExtentY        =   8758
      _Version        =   393216
      BackColorSel    =   8388608
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
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
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   108
      TabIndex        =   0
      Top             =   468
      Width           =   2172
   End
End
Attribute VB_Name = "frm090610_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim i As Integer
Private Sub cmdOK_Click()
frm090610.Show
frm090610.Enabled = True
'frm090610.MousePointer = vbDefault
Unload Me
End Sub

Private Sub Form_Load()
lbl1.Caption = frm090610.txt1(7) & " 年 " & frm090610.txt1(8) & " 月 "
MoveFormToCenter Me
Process
SetGrd1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090610_1 = Nothing
End Sub

Sub Process()
Dim StrSQLa As String

'Modify By Cheng 2003/10/01
'加當月個人合計
'strSQL = "SELECT R105001,SUM(R105002),SUM(R105003),SUM(R105004),SUM(R105005),SUM(R105006),SUM(R105007),SUM(R105008),SUM(R105009),SUM(R105010),SUM(R105011),SUM(R105012),SUM(R105013),SUM(R105014),SUM(R105015),SUM(R105016),SUM(R105017),SUM(R105018),SUM(R105019),SUM(R105020),SUM(R105021),SUM(R105022),SUM(R105023),SUM(R105024),SUM(R105025),SUM(R105026),SUM(R105027),SUM(R105028),SUM(R105029),SUM(R105030),SUM(R105031),SUM(R105032),R105033 FROM R090610 WHERE ID='" & strUserNum & "' GROUP BY R105001,R105033 "
StrSQLa = "SUM(Nvl(R105002,0))+SUM(Nvl(R105003,0))+SUM(Nvl(R105004,0))+SUM(Nvl(R105005,0))+SUM(Nvl(R105006,0))+SUM(Nvl(R105007,0))+SUM(Nvl(R105008,0))+SUM(Nvl(R105009,0))+SUM(Nvl(R105010,0))+SUM(Nvl(R105011,0))+SUM(Nvl(R105012,0))+SUM(Nvl(R105013,0))+SUM(Nvl(R105014,0))+SUM(Nvl(R105015,0))+SUM(Nvl(R105016,0))+SUM(Nvl(R105017,0))+SUM(Nvl(R105018,0))+SUM(Nvl(R105019,0))+SUM(Nvl(R105020,0))+SUM(Nvl(R105021,0))+SUM(Nvl(R105022,0))+SUM(Nvl(R105023,0))+SUM(Nvl(R105024,0))+SUM(Nvl(R105025,0))+SUM(Nvl(R105026,0))+SUM(Nvl(R105027,0))+SUM(Nvl(R105028,0))+SUM(Nvl(R105029,0))+SUM(Nvl(R105030,0))+SUM(Nvl(R105031,0))+SUM(Nvl(R105032,0))"
strSql = "SELECT Nvl(ST02, R105001)," & StrSQLa & ",SUM(R105002),SUM(R105003),SUM(R105004),SUM(R105005),SUM(R105006),SUM(R105007),SUM(R105008),SUM(R105009),SUM(R105010),SUM(R105011),SUM(R105012),SUM(R105013),SUM(R105014),SUM(R105015),SUM(R105016),SUM(R105017),SUM(R105018),SUM(R105019),SUM(R105020),SUM(R105021),SUM(R105022),SUM(R105023),SUM(R105024),SUM(R105025),SUM(R105026),SUM(R105027),SUM(R105028),SUM(R105029),SUM(R105030),SUM(R105031),SUM(R105032),R105033, ST06, R105001 FROM R090610, Staff WHERE R105001=ST01(+) And ID='" & strUserNum & "' GROUP BY Nvl(ST02, R105001), R105033, ST06, R105001 Order By ST06, R105001, R105033 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        Set grd1.Recordset = adoRecordset
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
End Sub

Private Sub SetGrd1()
With grd1
    .Cols = 35
    .row = 0
    .col = 0
    .ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "承辦人"
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "合計"
    .ColWidth(2) = 800
    .CellAlignment = flexAlignCenterCenter
    For i = 2 To 32
        .col = i + 1: .Text = str(i - 1)
        .ColWidth(i + 1) = 600
        .CellAlignment = flexAlignCenterCenter
    Next i
    .col = 34
    .ColWidth(34) = 0
    For i = 1 To .Rows - 1
        .row = i
        .col = 34
        Select Case Val(.Text)
        Case 0
              .col = 0
              .Text = ""
        Case 1
              .col = 0
              .Text = "設計"
        Case 2
              .col = 0
              .Text = "非設計"
        Case Else
        End Select
    Next i
End With
End Sub


