VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090711_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護_其他多國案本所案號"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9210
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(U)"
      Height          =   390
      Left            =   8010
      TabIndex        =   0
      Top             =   30
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5355
      Left            =   30
      TabIndex        =   1
      Top             =   435
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9446
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
Attribute VB_Name = "frm090711_4"
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
Set frm090711_4 = Nothing
End Sub

'傳入本所案號
'用本所案號串 caseRelation 的  cr01~08 且 cp21='Y' (國外案)，再用 cr05~04 串案件進度檔串基本檔
Sub StrMenu(strText As String)
'Modified by Morgan 2013/3/20 不必再限制只抓子案--瓊玉
'strSql = "SELECT na03,Cr05||'-'||Cr06||'-'||Cr07||'-'||Cr08,pa05,ST02 FROM caseRelation,CASEPROGRESS,STAFF,patent,nation WHERE cr01='" & SystemNumber(strText, 1) & "' and cr02='" & SystemNumber(strText, 2) & "' and cr03='" & SystemNumber(strText, 3) & "' and cr04='" & SystemNumber(strText, 4) & "' and cp21='Y' and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and cr05 in (" & SQLGrpStr("", 1) & ") and pa09=na01(+)  "
strSql = "SELECT na03,Cr05||'-'||Cr06||'-'||Cr07||'-'||Cr08,pa05,ST02 FROM caseRelation,CASEPROGRESS,STAFF,patent,nation WHERE cr01='" & SystemNumber(strText, 1) & "' and cr02='" & SystemNumber(strText, 2) & "' and cr03='" & SystemNumber(strText, 3) & "' and cr04='" & SystemNumber(strText, 4) & "' and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and cr05 in (" & SQLGrpStr("", 1) & ") and pa09=na01(+)  "
Set RS090201 = New ADODB.Recordset
RS090201.CursorLocation = adUseClient
RS090201.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If RS090201.RecordCount <> 0 Then
    Set GRD1.Recordset = RS090201
    SetGrd1
Else
    s = MsgBox("沒有多國案關聯資料！", , "沒有資料！")
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



