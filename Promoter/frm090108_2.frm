VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090108_2 
   Caption         =   "委查人委查統計"
   ClientHeight    =   5820
   ClientLeft      =   180
   ClientTop       =   636
   ClientWidth     =   9168
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   9168
   Begin VB.CommandButton CmdCancel 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   7042
      TabIndex        =   0
      Top             =   70
      Width           =   1150
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   0
      Left            =   8220
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4716
      Left            =   48
      TabIndex        =   2
      Top             =   600
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   8319
      _Version        =   393216
      Cols            =   8
      FixedRows       =   0
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
      _Band(0).Cols   =   8
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
      Left            =   8064
      TabIndex        =   7
      Top             =   5448
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
      Index           =   4
      Left            =   6780
      TabIndex        =   6
      Top             =   5460
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
      Left            =   5292
      TabIndex        =   5
      Top             =   5460
      Width           =   960
   End
   Begin VB.Label LblSubTotal 
      Height          =   252
      Index           =   2
      Left            =   3780
      TabIndex        =   4
      Top             =   5460
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "總計："
      Height          =   252
      Left            =   108
      TabIndex        =   3
      Top             =   5460
      Width           =   600
   End
End
Attribute VB_Name = "frm090108_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer, k As Integer, intK As Integer
'Modify By Cheng 2002/08/07
'Dim SubTotal(2 To 5) As Integer, GradeTotal(2 To 5) As Integer
Dim SubTotal(2 To 5) As Integer, SubTotalA(2 To 5) As Integer, GradeTotal(2 To 5) As Integer
Dim strSql As String, strTemp As Variant, StrTest As String, TmpKey1 As String, TmpKey1A As String
Dim Rs As New ADODB.Recordset
Dim bolIsTMA As Boolean 'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)

Private Sub GridHead()
   With MSHFlexGrid1
      .row = 0
      .col = 0:       .Text = "部　門"
      .ColWidth(0) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 1:       .Text = "委查人"
      .ColWidth(1) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 2:       .Text = "中文"
      .ColWidth(2) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 3:       .Text = "英文"
      .ColWidth(3) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 4:       .Text = "圖形"
      .ColWidth(4) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 5:       .Text = "小　計"
      .ColWidth(5) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 6:       .Text = ""
      .ColWidth(6) = 0
      .col = 7:       .Text = ""
      .ColWidth(7) = 0
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
   If DBDATE(frm090108_1.txtData(1)) >= 查名單網中系統啟用日 Or DBDATE(frm090108_1.txtData(2)) >= 查名單網中系統啟用日 Then
      bolIsTMA = True
   Else
      bolIsTMA = False
   End If
   'end 2024//11/15
   
   Me.Enabled = False
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   SubSQL = "": strCondition = "": strCondition = ""
   For i = 2 To 5
      SubTotal(i) = 0: GradeTotal(i) = 0
      'Add By Cheng 2002/08/07
      SubTotalA(i) = 0
   Next i

   If frm090108_1.txtData(4) <> Empty Then
      strCondition = strCondition + " AND ST03>='" & frm090108_1.txtData(4) & "'"
   End If
   If frm090108_1.txtData(5) <> Empty Then
      strCondition = strCondition + " AND ST03<='" & frm090108_1.txtData(5) & "'"
   End If
   If frm090108_1.txtData(4) <> Empty Or frm090108_1.txtData(5) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & frm090108_1.Label4 & frm090108_1.txtData(4) & "-" & frm090108_1.txtData(5) 'Add By Sindy 2010/12/14
   End If

   If frm090108_1.txtData(0) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         strCondition = strCondition + " AND TMA08 = '" & frm090108_1.txtData(0) & "'"
      Else
      'end 2024/11/15
         strCondition = strCondition + " AND TMQ02 = '" & frm090108_1.txtData(0) & "'"
      End If
      pub_QL05 = pub_QL05 & ";" & frm090108_1.Label1 & frm090108_1.txtData(0) & frm090108_1.LblTmq02NM 'Add By Sindy 2010/12/14
   End If
   If frm090108_1.txtData(1) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TO_CHAR(TMA04,'YYYYMMDD')>=" & Val(ChangeTStringToWString(frm090108_1.txtData(1))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ04>=" & Val(ChangeTStringToWString(frm090108_1.txtData(1))) & ""
      End If
   End If
   If frm090108_1.txtData(2) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TO_CHAR(TMA04,'YYYYMMDD')<=" & Val(ChangeTStringToWString(frm090108_1.txtData(2))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ04<=" & Val(ChangeTStringToWString(frm090108_1.txtData(2))) & ""
      End If
   End If
   If frm090108_1.txtData(1) <> Empty Or frm090108_1.txtData(2) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & frm090108_1.Label2 & frm090108_1.txtData(1) & "-" & frm090108_1.txtData(2) 'Add By Sindy 2010/12/14
   End If
   
   If Len(SubSQL) <> 0 Then
      SubSQL = " WHERE " & Mid(SubSQL, 5)
   End If
   
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      strSql = "SELECT DECODE(A0902, NULL, ' ', NVL(A0902, ST15)) AS 部門,NVL(ST02, TMA08) AS 委查人, SUM(NVL(TMA36, 0)) AS 中文, SUM(NVL(TMA37, 0)) AS 英文, SUM(NVL(TMA38, 0)) AS 圖形, SUM(NVL(TMA36, 0) + NVL(TMA37, 0) + NVL(TMA38, 0)) AS 小計, NVL(ST15, '   ') AS ST15, TMA08 " & _
               "FROM TMQAPPFORM, STAFF, ACC090 " & SubSQL & strCondition & " AND TMA08 = ST01(+) AND ST15 = A0901(+) AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' GROUP BY ST15, TMA08, DECODE(A0902, NULL, ' ', NVL(A0902, ST15)), NVL(ST02, TMA08)"
   Else
   'end 2024/11/15
      'Modify By Cheng 2002/08/07
   '   strsql = "SELECT DECODE(A0902, NULL, ' ', NVL(A0902, ST03)) AS 部門, NVL(ST02, TMQ02) AS 委查人, SUM(NVL(TMQ07, 0)) AS 中文, SUM(NVL(TMQ08, 0)) AS 英文, SUM(NVL(TMQ09, 0)) AS 圖形, SUM(NVL(TMQ07, 0) + NVL(TMQ08, 0) + NVL(TMQ09, 0)), NVL(ST03, ' '), TMQ02 FROM TRADEMARKQUERY, STAFF, ACC090 " & SubSQL & strCondition & strcondition & " AND TMQ02 = ST01(+) AND ST03 = A0901(+) GROUP BY ST03, TMQ02, DECODE(A0902, NULL, ' ', NVL(A0902, ST03)), NVL(ST02, TMQ02)"
      strSql = "SELECT DECODE(A0902, NULL, ' ', NVL(A0902, ST15)) AS 部門, NVL(ST02, TMQ02) AS 委查人, SUM(NVL(TMQ07, 0)) AS 中文, SUM(NVL(TMQ08, 0)) AS 英文, SUM(NVL(TMQ09, 0)) AS 圖形, SUM(NVL(TMQ07, 0) + NVL(TMQ08, 0) + NVL(TMQ09, 0)), NVL(ST15, '   '), TMQ02 FROM TRADEMARKQUERY, STAFF, ACC090 " & SubSQL & strCondition & strCondition & " AND TMQ02 = ST01(+) AND ST15 = A0901(+) GROUP BY ST15, TMQ02, DECODE(A0902, NULL, ' ', NVL(A0902, ST15)), NVL(ST02, TMQ02)"
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not Rs.RecordCount > 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/14
      Me.Enabled = False
      Exit Sub
   End If
   InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/14
   GridHead
   i = 1: intK = 0
   With MSHFlexGrid1
      .Rows = Rs.RecordCount: TmpKey1 = ""
      'Add By Cheng 2002/08/07
      TmpKey1A = ""
      Rs.MoveFirst
      Do
         intK = intK + 1
         If .Rows = i Then .Rows = .Rows + 1
         .row = i
         If TmpKey1 <> Empty And TmpKey1 <> Rs.Fields(6) Then
            .col = 1
            .Text = "小　計："
            For k = 2 To 5
               .col = k
               .Text = Format(SubTotal(k), "###,###,###,###")
            Next k
            i = i + 2
            If .Rows <= i Then .Rows = .Rows + 2
            .row = i
            For k = 2 To 5
               SubTotal(k) = 0
            Next k
            'Add By Cheng 2002/08/07
            'ST15前二碼不同者要小計
            If TmpKey1A <> Left(Rs.Fields(6), 2) Then
               .col = 1
               .Text = "合　計："
               For k = 2 To 5
                  .col = k
                  .Text = Format(SubTotalA(k), "###,###,###,###")
               Next k
               i = i + 2
               If .Rows <= i Then .Rows = .Rows + 2
               .row = i
               For k = 2 To 5
                  SubTotalA(k) = 0
               Next k
            End If
         
         End If
         .col = 0
         .Text = StrConv(MidB(StrConv(Rs.Fields(0), vbFromUnicode), 1, 10), vbUnicode)
         If TmpKey1 <> Empty And TmpKey1 = Rs.Fields(6) Then
            .Text = ""
         End If
         .col = 1
         .Text = Rs.Fields(1)
         For j = 2 To 5
            .col = j
            SubTotal(j) = SubTotal(j) + CLng(Rs.Fields(j))
            'Add By Cheng 2002/08/07
            SubTotalA(j) = SubTotalA(j) + CLng(Rs.Fields(j))
            GradeTotal(j) = GradeTotal(j) + CLng(Rs.Fields(j))
            .CellAlignment = 7
            .Text = Format(Rs.Fields(j), "###,###,###,###")
         Next j
         DoEvents
         TmpKey1 = Rs.Fields(6)
         'Add By Cheng 2002/08/07
         TmpKey1A = Left(Rs.Fields(6), 2)
         Rs.MoveNext
         i = i + 1
      Loop Until Rs.EOF
      If .Rows = i Then .Rows = .Rows + 2
      .row = i
      .col = 1
      .Text = "小　計："
      For k = 2 To 5
         .col = k
         .Text = Format(SubTotal(k), "###,###,###,###")
      Next k
      'Add By Cheng 2002/08/07
      i = i + 2
      If .Rows = i Then .Rows = .Rows + 1
      .row = i
      .col = 1
      .Text = "合　計："
      For k = 2 To 5
         .col = k
         .Text = Format(SubTotalA(k), "###,###,###,###")
      Next k
   
   End With
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   For i = 2 To 4
      LblSubTotal(i).Caption = Format(GradeTotal(i), "###,###,###,###")
   Next i
   LblTotal.Caption = Format(GradeTotal(5), "###,###,###,###")
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090108_2 = Nothing
End Sub
