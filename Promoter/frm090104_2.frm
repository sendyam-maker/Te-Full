VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090104_2 
   Caption         =   "查名人查覆明細"
   ClientHeight    =   5820
   ClientLeft      =   180
   ClientTop       =   540
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
      Left            =   7100
      TabIndex        =   0
      Top             =   72
      Width           =   1150
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   0
      Left            =   8268
      TabIndex        =   1
      Top             =   72
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5148
      Left            =   72
      TabIndex        =   2
      Top             =   600
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   9081
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   11
   End
End
Attribute VB_Name = "frm090104_2"
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
Dim strSql As String, strTemp As Variant, StrTest As String
Dim Rs As New ADODB.Recordset
Dim bolIsTMA As Boolean 'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)

Private Sub GridHead()
   With MSHFlexGrid1
      .row = 0
      .col = 0:       .Text = "委查單號"
      .ColWidth(0) = 1000
      .CellAlignment = flexAlignCenterCenter
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         .col = 1:       .Text = "類別組群"
      Else
      'end 2024/11/15
         .col = 1:       .Text = "組群"
      End If
      .ColWidth(1) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 2:       .Text = "委查人"
      .ColWidth(2) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 3:       .Text = "委查日"
      .ColWidth(3) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 4:       .Text = "收件日"
      .ColWidth(4) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 5:       .Text = "中文"
      .ColWidth(5) = 500
      .CellAlignment = flexAlignCenterCenter
      .col = 6:       .Text = "英文"
      .ColWidth(6) = 500
      .CellAlignment = flexAlignCenterCenter
      .col = 7:       .Text = "圖形"
      .ColWidth(7) = 500
      .CellAlignment = flexAlignCenterCenter
      .col = 8:       .Text = "查名人"
      .ColWidth(8) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 9:       .Text = "查覆日期"
      .ColWidth(9) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 10:       .Text = "期限日期"
      .ColWidth(10) = 900
      .CellAlignment = flexAlignCenterCenter
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
   Me.Height = 6555
   Me.Width = 9285
   MoveFormToCenter Me
   bolToEndByNick = False
End Sub

Sub GridData()
Dim OrderStr As String, SubSQL As String, strCondition As String, strTemp As String, strTemp1 As String

   'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)
   If DBDATE(frm090104_1.Txtdata(1)) >= 查名單網中系統啟用日 Or DBDATE(frm090104_1.Txtdata(2)) >= 查名單網中系統啟用日 Then
      bolIsTMA = True
   Else
      bolIsTMA = False
   End If
   'end 2024//11/15
   
   If frm090104_1.OptBtn(0).Value Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         OrderStr = "委查單號"
      Else
      'end 2024/11/15
         OrderStr = "TMQ01"
      End If
      pub_QL05 = pub_QL05 & ";" & frm090104_1.Frame1.Caption & frm090104_1.OptBtn(0).Caption 'Add By Sindy 2010/12/13
   Else
      If frm090104_1.OptBtn(1).Value Then
         'Added by Lydia 2024/11/15 查名單(網中)
         If bolIsTMA = True Then
            OrderStr = "委查人, 委查單號"
         Else
         'end 2024/11/15
            OrderStr = "TMQ02, TMQ01"
         End If
         pub_QL05 = pub_QL05 & ";" & frm090104_1.Frame1.Caption & frm090104_1.OptBtn(1).Caption 'Add By Sindy 2010/12/13
      Else
         If frm090104_1.OptBtn(2).Value Then
            'Added by Lydia 2024/11/15 查名單(網中)
            If bolIsTMA = True Then
               OrderStr = "委查日, 委查單號"
            Else
            'end 2024/11/15
               OrderStr = "TMQ04, TMQ01"
            End If
            pub_QL05 = pub_QL05 & ";" & frm090104_1.Frame1.Caption & frm090104_1.OptBtn(2).Caption 'Add By Sindy 2010/12/13
         Else
            If frm090104_1.OptBtn(3).Value Then
               'Added by Lydia 2024/11/15 查名單(網中)
               If bolIsTMA = True Then
                  OrderStr = "收件日, 委查單號"
               Else
               'end 2024/11/15
                  OrderStr = "TMQ05, TMQ01"
               End If
               pub_QL05 = pub_QL05 & ";" & frm090104_1.Frame1.Caption & frm090104_1.OptBtn(3).Caption 'Add By Sindy 2010/12/13
            End If
         End If
      End If
   End If
   Me.Enabled = False
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   SubSQL = "": strCondition = ""
   If frm090104_1.Txtdata(0) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         strCondition = strCondition + " AND TMA10 = '" & frm090104_1.Txtdata(0) & "'"
      Else
      'end 2024/11/15
         strCondition = strCondition + " AND TMQ10 = '" & frm090104_1.Txtdata(0) & "'"
      End If
      pub_QL05 = pub_QL05 & ";" & frm090104_1.Label1 & frm090104_1.Txtdata(0) & frm090104_1.LblTmq10NM 'Add By Sindy 2010/12/13
   End If
   If frm090104_1.Txtdata(1) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14>=" & Val(ChangeTStringToWString(frm090104_1.Txtdata(1))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ11>=" & Val(ChangeTStringToWString(frm090104_1.Txtdata(1))) & ""
      End If
   End If
   If frm090104_1.Txtdata(2) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14<=" & Val(ChangeTStringToWString(frm090104_1.Txtdata(2))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ11<=" & Val(ChangeTStringToWString(frm090104_1.Txtdata(2))) & ""
      End If
   End If
   If frm090104_1.Txtdata(1) <> Empty Or frm090104_1.Txtdata(2) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & frm090104_1.Label2 & frm090104_1.Txtdata(1) & "-" & frm090104_1.Txtdata(2) 'Add By Sindy 2010/12/13
   End If
   
   If Len(SubSQL) <> 0 Then
      SubSQL = " WHERE " & Mid(SubSQL, 5)
   End If
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      strSql = " SELECT TMA01 AS 委查單號, " & PUB_GetTMAforClass & " AS 類別組群, NVL(S1.ST02, TMA08) 委查人, TO_CHAR(TMA04,'YYYYMMDD') AS 委查日, TMA09 AS 收件日, TMA36 AS 中文, TMA37 AS 英文, TMA38 AS 圖形, NVL(S2.ST02, TMA10) AS 查名人, TMA14 AS 查覆日期, NVL(TMA11,TMA12) AS 期限日期" & _
               " FROM TMQAPPFORM, STAFF S1, STAFF S2 " & SubSQL & strCondition & " AND TMA08 = S1.ST01(+) AND TMA10 = S2.ST01(+) AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' ORDER BY " & OrderStr
   Else
   'end 2024/11/15
      strSql = "SELECT TMQ01 AS 委查單號, TMQ03 AS 組群, NVL(S1.ST02, TMQ02) AS 委查人, TMQ04 AS 委查日, TMQ05 AS 收件日, TMQ07 AS 中文, TMQ08 AS 英文, TMQ09 AS 圖形, NVL(S2.ST02, TMQ10) AS 查名人, TMQ11 AS 查覆日期, TMQ06 AS 期限日期 FROM TRADEMARKQUERY, STAFF S1, STAFF S2 " & SubSQL & strCondition & " AND TMQ02 = S1.ST01(+) AND TMQ10 = S2.ST01(+) ORDER BY " & OrderStr
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not Rs.RecordCount > 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/13
      Me.Enabled = False
      Exit Sub
   End If
   InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/13
   intK = Rs.RecordCount
   Set MSHFlexGrid1.Recordset = Rs
   GridHead
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .row = i
         .col = 3
         .Text = ChangeWStringToTString(.Text)
         .CellAlignment = 4
         .col = 4
         .Text = ChangeWStringToTString(.Text)
         .CellAlignment = 4
         .col = 5
         .CellAlignment = 7
         .Text = Format(.Text, "#,###")
         .col = 6
         .CellAlignment = 7
         .Text = Format(.Text, "#,###")
         .col = 7
         .CellAlignment = 7
         .Text = Format(.Text, "#,###")
         .col = 9
         .Text = ChangeWStringToTString(.Text)
         .CellAlignment = 4
         .col = 10
         .Text = ChangeWStringToTString(.Text)
         .CellAlignment = 4
'         'FRM100.Tag = str(IntK) + "=" + str(i)
'         'FRM100.StrMenu
         DoEvents
'         'FRM100.Refresh
      Next i
   End With
   'Unload frm100
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090104_2 = Nothing
End Sub

