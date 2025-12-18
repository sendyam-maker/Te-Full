VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090109_2 
   Caption         =   "商品類別委查統計"
   ClientHeight    =   5820
   ClientLeft      =   216
   ClientTop       =   660
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   9156
   Begin VB.CommandButton CmdCancel 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   7066
      TabIndex        =   0
      Top             =   70
      Width           =   1150
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   0
      Left            =   8244
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4716
      Left            =   48
      TabIndex        =   2
      Top             =   576
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   8319
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(8888：團體標章)"
      Height          =   180
      Left            =   270
      TabIndex        =   9
      Top             =   90
      Width           =   1380
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "(9999：證明標章)"
      Height          =   180
      Left            =   270
      TabIndex        =   8
      Top             =   330
      Width           =   1380
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
      Height          =   255
      Left            =   6790
      TabIndex        =   7
      Top             =   5450
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
      Height          =   255
      Index           =   3
      Left            =   5290
      TabIndex        =   6
      Top             =   5450
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
      Height          =   255
      Index           =   2
      Left            =   3790
      TabIndex        =   5
      Top             =   5450
      Width           =   960
   End
   Begin VB.Label LblSubTotal 
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   5450
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "總計："
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   5450
      Width           =   600
   End
End
Attribute VB_Name = "frm090109_2"
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
Dim SubTotal(1 To 4) As Variant
Dim strSql As String, strTemp(0 To 4) As String
Dim StrArray As Variant, ComArray As Variant
Dim Rs As New ADODB.Recordset
Dim bolIsTMA As Boolean 'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)

Private Sub GridHead()
   With MSHFlexGrid1
      .row = 0
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         .col = 0:       .Text = "商品類別組群"
      Else
      'end 2024/11/15
         .col = 0:       .Text = "商品組群"
      End If
      .ColWidth(0) = 1300
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
Dim SubSQL As String

   'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)
   If DBDATE(frm090109_1.Txtdata(1)) >= 查名單網中系統啟用日 Or DBDATE(frm090109_1.Txtdata(2)) >= 查名單網中系統啟用日 Then
      bolIsTMA = True
      Label12.Visible = False
      Label14.Visible = False
   Else
      bolIsTMA = False
      Label12.Visible = True
      Label14.Visible = True
   End If
   'end 2024//11/15
   
   Me.Enabled = False
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   cnnConnection.Execute "DELETE FROM R090109 WHERE ID='" & strUserNum & "' "
   SubSQL = ""
   For i = 1 To 4
      SubTotal(i) = 0
   Next i
   If frm090109_1.Txtdata(1) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14>=" & Val(ChangeTStringToWString(frm090109_1.Txtdata(1))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ11>=" & Val(ChangeTStringToWString(frm090109_1.Txtdata(1))) & ""
      End If
   End If
   If frm090109_1.Txtdata(2) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14<=" & Val(ChangeTStringToWString(frm090109_1.Txtdata(2))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " TMQ11<=" & Val(ChangeTStringToWString(frm090109_1.Txtdata(2))) & ""
      End If
   End If
   If frm090109_1.Txtdata(1) <> Empty Or frm090109_1.Txtdata(2) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & frm090109_1.Label2 & frm090109_1.Txtdata(1) & "-" & frm090109_1.Txtdata(2) 'Add By Sindy 2010/12/14
   End If
   If frm090109_1.Txtdata(0) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm090109_1.Label1 & frm090109_1.Txtdata(0) 'Add By Sindy 2010/12/14
   End If
   If frm090109_1.Txtdata(4) = "1" Then
      pub_QL05 = pub_QL05 & ";" & frm090109_1.Label8 & "1：群組" 'Add By Sindy 2010/12/14
   ElseIf frm090109_1.Txtdata(4) = "2" Then
      pub_QL05 = pub_QL05 & ";" & frm090109_1.Label8 & "2：中文筆數" 'Add By Sindy 2010/12/14
   ElseIf frm090109_1.Txtdata(4) = "3" Then
      pub_QL05 = pub_QL05 & ";" & frm090109_1.Label8 & "3：英文筆數" 'Add By Sindy 2010/12/14
   Else
      pub_QL05 = pub_QL05 & ";" & frm090109_1.Label8 & "4：圖形筆數" 'Add By Sindy 2010/12/14
   End If
   
   If Len(SubSQL) <> 0 Then
      SubSQL = " WHERE " & Mid(SubSQL, 5)
   End If
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      strSql = "SELECT " & PUB_GetTMAforClass & " AS 類別組群, SUM(NVL(TMA36, 0)) AS 中文, SUM(NVL(TMA37, 0)) AS 英文, SUM(NVL(TMA38, 0)) AS 圖形 " & _
               "FROM TMQAPPFORM " & SubSQL & " AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' GROUP BY " & PUB_GetTMAforClass & " ORDER  BY 1"
   Else
   'end 2024/11/15
      strSql = "SELECT TMQ03, NVL(TMQ07, 0), NVL(TMQ08, 0), NVL(TMQ09, 0) FROM TRADEMARKQUERY " & SubSQL & " ORDER BY TMQ03"
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not Rs.RecordCount > 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/14
      Me.Enabled = False
      Exit Sub
   Else
      With Rs
         .MoveFirst
         j = 0
         DoEvents
         Do While .EOF = False
            For i = 0 To 3
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            Insert_Temp
            j = j + 1
            DoEvents
            .MoveNext
         Loop
      End With
      Process
      If Not Rs.RecordCount > 0 Then
         Me.Enabled = True
         Exit Sub
      End If
   End If
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
'         'FRM100.Tag = str(IntK) + "=" + str(i)
 '        'FRM100.StrMenu
         DoEvents
  '       'FRM100.Refresh
      Next i
   End With
   'Unload frm100
   For i = 1 To 3
      LblSubTotal(i).Caption = Format(SubTotal(i), "###,###,###,###")
   Next i
   LblTotal.Caption = Format(SubTotal(4), "###,###,###,###")
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090109_2 = Nothing
End Sub

'寫入暫存檔
Private Sub Insert_Temp()
Dim BlnIsNew As Boolean, tmpArray As Variant
   'Modify By Cheng 2002/03/26
'   ComArray = Split(strTemp(0), ",")
   'Modified by Lydia 2019/03/07
   'ComArray = Split(strTemp(0), ".")
   strExc(0) = Replace(strTemp(0), ",", ".")
   ComArray = Split(strExc(0), ".")
   'end 2019/03/07
   For i = 0 To UBound(ComArray)
      BlnIsNew = False
      If Len(frm090109_1.Txtdata(0)) <> 0 Then
         StrArray = Split(frm090109_1.Txtdata(0), ",")
         tmpArray = Filter(StrArray, ComArray(i))
         For j = 0 To UBound(tmpArray)
            If tmpArray(j) <> Empty Then
               BlnIsNew = True
               j = UBound(tmpArray) + 1
            End If
         Next j
      Else
         BlnIsNew = True
      End If
      If BlnIsNew Then
         strSql = "INSERT INTO R090109 VALUES('" & ComArray(i) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",'" & strUserNum & "')"
         cnnConnection.Execute strSql
      End If
   Next i
End Sub

Private Sub Process()
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   'Modify By Cheng 2002/03/26
'   strSQL = "SELECT R001001, SUM(NVL(R001002, 0)), SUM(NVL(R001003, 0)), SUM(NVL(R001004, 0)), SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY R001001"
   '依群組排序
   If frm090109_1.Txtdata(4).Text = "1" Then
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 組群"
   '依中文筆數排序
   ElseIf frm090109_1.Txtdata(4).Text = "2" Then
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 中文筆數 DESC"
   '依英文筆數排序
   ElseIf frm090109_1.Txtdata(4).Text = "3" Then
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 英文筆數 DESC"
   '依圖形筆數排序
   Else
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 圖形筆數 DESC"
   End If
   
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/14
      cnnConnection.Execute "DELETE FROM R090109 WHERE ID='" & strUserNum & "' "
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/14
   End If
End Sub


