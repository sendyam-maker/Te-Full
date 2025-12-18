VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090711_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護_其他國內案本所案號"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9090
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(U)"
      Height          =   390
      Left            =   7890
      TabIndex        =   0
      Top             =   30
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5325
      Left            =   30
      TabIndex        =   1
      Top             =   465
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9393
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
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
End
Attribute VB_Name = "frm090711_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (grd1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String
Dim RS090201 As New ADODB.Recordset
Dim s As Integer

Private Sub cmdOK_Click()
frm090711.Show
Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090711_5 = Nothing
End Sub

'傳入本所案號
'用本所案號串 caseMap 的  cm05~08 且 cm10='0' (國外案)，再用 cm01~04 串案件進度檔串基本檔
Sub StrMenu(strText As String)
strSql = "SELECT na03,CM05||'-'||CM06||'-'||CM07||'-'||CM08,pa05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,patent,nation WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=pa01(+) and cm06=pa02(+) and cm07=pa03(+) and cm08=pa04(+) and pa09=na01(+) and cm05 in (" & SQLGrpStr("", 1) & ") "
Set RS090201 = New ADODB.Recordset
RS090201.CursorLocation = adUseClient
RS090201.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If RS090201.RecordCount <> 0 Then
    Set GRD1.Recordset = RS090201
    SetGrd1
Else
    s = MsgBox("沒有國內案關聯資料！", , "沒有資料！")
    frm090711.Show
    Unload Me
End If
Set RS090201 = Nothing
End Sub

Sub SetGrd1()
With GRD1
    .Visible = False
    .Cols = 4
    .row = 0
    .col = 0:   .Text = "申請國家"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "本所案號"
    .ColWidth(1) = 1600
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "案件名稱"
    .ColWidth(2) = 3300
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "承辦人"
    .ColWidth(3) = 1000
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End With
End Sub



