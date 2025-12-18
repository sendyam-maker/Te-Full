VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090613_2_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件處理時間_其他國內外案本所案號"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6210
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(U)"
      Height          =   390
      Left            =   5010
      TabIndex        =   1
      Top             =   30
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4125
      Left            =   15
      TabIndex        =   0
      Top             =   435
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   7276
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
Attribute VB_Name = "frm090613_2_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/08 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'create by nickc 2006/04/27  copy from frm090201_2_1
Option Explicit

Dim strSql As String
Dim RS090201 As New ADODB.Recordset
Dim s As Integer

Private Sub cmdOK_Click()
frm090613_2.Show
Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090613_2_1 = Nothing
End Sub

'傳入本所案號
'用本所案號串 caseMap 的  cm05~08 且 cm10='0' (國外案)，再用 cm01~04 串案件進度檔串基本檔
Sub StrMenu(strText As String)
strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,pa05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,patent                                          WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and cm01 in (" & SQLGrpStr("", 1) & ") "
strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,tm05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,trademark       WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=tm01(+) and cm02=tm02(+) and cm03=tm03(+) and cm04=tm04(+) and cm01 in (" & SQLGrpStr("", 2) & ") "
strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,lc05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,lawcase            WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=lc01(+) and cm02=lc02(+) and cm03=lc03(+) and cm04=lc04(+) and cm01 in (" & SQLGrpStr("", 3) & ") "
strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,hc06,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,hirecase           WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=hc01(+) and cm02=hc02(+) and cm03=hc03(+) and cm04=hc04(+) and cm01 in (" & SQLGrpStr("", 4) & ") "
strSql = strSql & " union all SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,sp05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,servicepractice WHERE cm05='" & SystemNumber(strText, 1) & "' and cm06='" & SystemNumber(strText, 2) & "' and cm07='" & SystemNumber(strText, 3) & "' and cm08='" & SystemNumber(strText, 4) & "' and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm01=sp01(+) and cm02=sp02(+) and cm03=sp03(+) and cm04=sp04(+) and cm01 in (" & SQLGrpStr("", 5) & ")  "
'add by nick 2005/02/17 陳玲玲填請做單要國內外皆可
strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,pa05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,patent             WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=pa01(+) and cm06=pa02(+) and cm07=pa03(+) and cm08=pa04(+) and cm05 in (" & SQLGrpStr("", 1) & ") "
strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,tm05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,trademark       WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=tm01(+) and cm06=tm02(+) and cm07=tm03(+) and cm08=tm04(+) and cm05 in (" & SQLGrpStr("", 2) & ") "
strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,lc05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,lawcase            WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=lc01(+) and cm06=lc02(+) and cm07=lc03(+) and cm08=lc04(+) and cm05 in (" & SQLGrpStr("", 3) & ") "
strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,hc06,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,hirecase           WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=hc01(+) and cm06=hc02(+) and cm07=hc03(+) and cm08=hc04(+) and cm05 in (" & SQLGrpStr("", 4) & ") "
strSql = strSql & " union all SELECT CM05||'-'||CM06||'-'||CM07||'-'||CM08,sp05,ST02 FROM CASEMAP,CASEPROGRESS,STAFF,servicepractice WHERE cm01='" & SystemNumber(strText, 1) & "' and cm02='" & SystemNumber(strText, 2) & "' and cm03='" & SystemNumber(strText, 3) & "' and cm04='" & SystemNumber(strText, 4) & "' and cm10='0' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05=sp01(+) and cm06=sp02(+) and cm07=sp03(+) and cm08=sp04(+) and cm05 in (" & SQLGrpStr("", 5) & ")  order by 1 "


Set RS090201 = New ADODB.Recordset
RS090201.CursorLocation = adUseClient
RS090201.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If RS090201.RecordCount <> 0 Then
    Set grd1.Recordset = RS090201
    SetGrd1
Else
    s = MsgBox("沒有國內外案關聯資料！", , "沒有資料！")
    frm090613_2.Show
    Unload Me
End If
Set RS090201 = Nothing
End Sub

Sub SetGrd1()
With grd1
    .Visible = False
    .Cols = 3
    .row = 0
    .col = 0:   .Text = "本所案號"
    .ColWidth(0) = 1600
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "案件名稱"
    .ColWidth(1) = 3300
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "承辦人"
    .ColWidth(2) = 1000
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End With
End Sub


