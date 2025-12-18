VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090107_2 
   Caption         =   "查名人查覆統計"
   ClientHeight    =   5820
   ClientLeft      =   -2952
   ClientTop       =   1080
   ClientWidth     =   9168
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   9168
   Begin VB.CommandButton CmdCancel 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   7078
      TabIndex        =   0
      Top             =   70
      Width           =   1150
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   0
      Left            =   8256
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4860
      Left            =   72
      TabIndex        =   2
      Top             =   552
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   8573
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   5
   End
   Begin VB.Label LblTotal 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      Height          =   252
      Left            =   6840
      TabIndex        =   7
      Top             =   5480
      Width           =   960
   End
   Begin VB.Label LblSubTotal 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   5340
      TabIndex        =   6
      Top             =   5480
      Width           =   960
   End
   Begin VB.Label LblSubTotal 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   5480
      Width           =   960
   End
   Begin VB.Label LblSubTotal 
      Height          =   252
      Index           =   1
      Left            =   2328
      TabIndex        =   4
      Top             =   5480
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "總計："
      Height          =   252
      Left            =   156
      TabIndex        =   3
      Top             =   5480
      Width           =   600
   End
End
Attribute VB_Name = "frm090107_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer, intK As Integer
Dim SubTotal(1 To 4) As Integer
Dim strSql As String, strTemp As Variant, StrTest As String
Dim Rs As New ADODB.Recordset
Dim bolIsTMA As Boolean 'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)

Private Sub GridHead()
   With MSHFlexGrid1
      .row = 0
      .col = 0:       .Text = "查名人"
      .ColWidth(0) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 1:       .Text = "中文"
      .ColWidth(1) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 2:       .Text = "英文"
      .ColWidth(2) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 3:       .Text = "圖形"
      .ColWidth(3) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 4:       .Text = "小　計"
      .ColWidth(4) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 5:       .Text = ""
      .ColWidth(5) = 0
   End With
End Sub

Private Sub cmdCancel_Click(Index As Integer)
   Me.Hide
End Sub

Private Sub cmdExit_Click(Index As Integer)
   bolToEndByNick = True
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Height = 6990
   Me.Width = 9285
   MoveFormToCenter Me
   bolToEndByNick = False
End Sub

Sub GridData()
Dim SubSQL As String, strCondition As String, strTemp As String, strTemp1 As String

   'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)
   If DBDATE(frm090107_1.Txtdata(1)) >= 查名單網中系統啟用日 Or DBDATE(frm090107_1.Txtdata(2)) >= 查名單網中系統啟用日 Then
      bolIsTMA = True
   Else
      bolIsTMA = False
   End If
   'end 2024//11/15
   
   Me.Enabled = False
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   SubSQL = "": strCondition = ""
   For i = 1 To 4
      SubTotal(i) = 0
   Next i
   If frm090107_1.Txtdata(0) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         strCondition = strCondition + " AND TMA10 = '" & frm090107_1.Txtdata(0) & "'"
      Else
      'end 2024/11/15
         strCondition = strCondition + " AND TMQ10 = '" & frm090107_1.Txtdata(0) & "'"
      End If
      pub_QL05 = pub_QL05 & ";" & frm090107_1.Label1 & frm090107_1.Txtdata(0) & frm090107_1.LblTmq10NM 'Add By Sindy 2010/12/14
   End If
   If frm090107_1.Txtdata(1) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14>=" & Val(ChangeTStringToWString(frm090107_1.Txtdata(1))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ11>=" & Val(ChangeTStringToWString(frm090107_1.Txtdata(1))) & ""
      End If
   End If
   If frm090107_1.Txtdata(2) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14<=" & Val(ChangeTStringToWString(frm090107_1.Txtdata(2))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ11<=" & Val(ChangeTStringToWString(frm090107_1.Txtdata(2))) & ""
      End If
   End If
   If frm090107_1.Txtdata(1) <> Empty Or frm090107_1.Txtdata(2) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & frm090107_1.Label2 & frm090107_1.Txtdata(1) & "-" & frm090107_1.Txtdata(2) 'Add By Sindy 2010/12/14
   End If
   
   If Len(SubSQL) <> 0 Then
      SubSQL = " WHERE " & Mid(SubSQL, 5)
   End If
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      strSql = "SELECT NVL(ST02, TMA10) AS 查名人, SUM(NVL(TMA36, 0)) AS 中文, SUM(NVL(TMA37, 0)) AS 英文, SUM(NVL(TMA38, 0)) AS 圖形, SUM(NVL(TMA36, 0) + NVL(TMA37, 0) + NVL(TMA38, 0)) AS 小計, TMA10 " & _
               "FROM TMQAPPFORM, STAFF " & SubSQL & strCondition & " AND TMA10 = ST01(+) AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' GROUP BY TMA10, NVL(ST02, TMA10)"
   Else
   'end 2024/11/15
      strSql = "SELECT NVL(ST02, TMQ10) AS 查名人, SUM(NVL(TMQ07, 0)) AS 中文, SUM(NVL(TMQ08, 0)) AS 英文, SUM(NVL(TMQ09, 0)) AS 圖形, SUM(NVL(TMQ07, 0) + NVL(TMQ08, 0) + NVL(TMQ09, 0)), TMQ10 FROM TRADEMARKQUERY, STAFF " & SubSQL & strCondition & " AND TMQ10 = ST01(+) GROUP BY TMQ10, NVL(ST02, TMQ10)"
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not Rs.RecordCount > 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/14
      Me.Enabled = False
      Exit Sub
   End If
   InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/14
   intK = Rs.RecordCount
   Set MSHFlexGrid1.Recordset = Rs
   GridHead
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .row = i
         For j = 1 To 4
         .col = j
         .CellAlignment = 7
         SubTotal(j) = SubTotal(j) + CLng(.Text)
         .Text = Format(.Text, "###,###,###,###")
         Next j
         DoEvents
      Next i
   End With
   For i = 1 To 3
      LblSubTotal(i).Caption = Format(SubTotal(i), "###,###,###,###")
   Next i
   LblTotal.Caption = Format(SubTotal(4), "###,###,###,###")
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090107_2 = Nothing
End Sub
