VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090613_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件處理時間統計查詢"
   ClientHeight    =   5724
   ClientLeft      =   -228
   ClientTop       =   1008
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   9324
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   7860
      TabIndex        =   10
      Top             =   20
      Width           =   1170
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "錯誤資料(&L)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6690
      TabIndex        =   9
      Top             =   20
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "收文-發文"
      Height          =   600
      Index           =   6
      Left            =   7860
      TabIndex        =   8
      Top             =   480
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "會稿完成-發文"
      Height          =   600
      Index           =   5
      Left            =   6690
      TabIndex        =   7
      Top             =   480
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "墨圖齊備-墨圖完稿"
      Height          =   600
      Index           =   4
      Left            =   5520
      TabIndex        =   6
      Top             =   480
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "會稿-會稿完成"
      Height          =   600
      Index           =   3
      Left            =   4350
      TabIndex        =   5
      Top             =   480
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "齊備-會稿"
      Height          =   600
      Index           =   2
      Left            =   3180
      TabIndex        =   4
      Top             =   480
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "草圖齊備-草圖完稿"
      Height          =   600
      Index           =   1
      Left            =   2010
      TabIndex        =   3
      Top             =   480
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "收文-齊備"
      Height          =   600
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1170
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4440
      Left            =   0
      TabIndex        =   1
      Top             =   1050
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7832
      _Version        =   393216
      Cols            =   1
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "游標指到會變黑色的欄位表示可點選顯示明細資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   390
      TabIndex        =   11
      Top             =   5520
      Width           =   4440
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   75
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frm090613_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/08 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim j As Integer, i As Integer
Dim m_iCol As Integer, m_iRow As Integer


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Me.Hide
     'Load frm090613_2
     'edit by nickc 2007/03/08
     'strSQL = "select DISTINCT r109017 from r090613 where id='" & strUserNum & "' and r109016='#' order by r109017 "
     strSql = "select DISTINCT r109017 from r090613 where id='" & strUserNum & "' and r109016 is not null order by r109017 "
     CheckOC
     frm090613_2.Combo1.Clear
     j = 0
     With adoRecordset
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            .MoveFirst
            Do While .EOF = False
                frm090613_2.Combo1.AddItem CheckStr(.Fields(0)), j
                j = j + 1
                .MoveNext
            Loop
            'add by nickc 2006/04/26
            frm090613_2.Combo1.Text = frm090613_2.Combo1.List(0)
            frm090613_2.StrMenu
            frm090613_2.Show
        Else
            Unload frm090613_2
            Me.Show
            MsgBox "沒有錯誤資料！", vbExclamation, "警告！"
        End If
     End With
     CheckOC
     'edit by nickc 2006/04/26 往上搬
     'frm090613_2.Combo1.Text = frm090613_2.Combo1.List(0)
     'frm090613_2.StrMenu
     'frm090613_2.Show
Case 1
     Me.Hide
     frm090613.Show
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetGrd1
'Modify By Sindy 2015/8/25
'If frm090613.Option1(0).Value = True Then
'    lbl1.Caption = frm090613.Txt1(5) & " 年 " & frm090613.Txt1(6) & " 月    收文"
'Else
'    lbl1.Caption = frm090613.Txt1(7) & " 年 " & frm090613.Txt1(8) & " 月    發文"
'End If
If frm090613.Option1(0).Value = True Then
   lbl1.Caption = ChangeTStringToTDateString(frm090613.Txt1(5)) & "~" & ChangeTStringToTDateString(frm090613.Txt1(6)) & "　　收文"
ElseIf frm090613.Option1(1).Value = True Then
   lbl1.Caption = ChangeTStringToTDateString(frm090613.Txt1(7)) & "~" & ChangeTStringToTDateString(frm090613.Txt1(8)) & "　　發文"
'Added by Morgan 2017/8/17
ElseIf frm090613.Option1(5).Value = True Then
   lbl1.Caption = ChangeTStringToTDateString(frm090613.Txt1(28)) & "~" & ChangeTStringToTDateString(frm090613.Txt1(29)) & "　　齊備"
'end 2017/8/17
ElseIf frm090613.Option1(2).Value = True Then
   lbl1.Caption = ChangeTStringToTDateString(frm090613.Txt1(21)) & "~" & ChangeTStringToTDateString(frm090613.Txt1(22)) & "　　完稿"
ElseIf frm090613.Option1(3).Value = True Then
   lbl1.Caption = ChangeTStringToTDateString(frm090613.Txt1(23)) & "~" & ChangeTStringToTDateString(frm090613.Txt1(24)) & "　　會完"
Else
   lbl1.Caption = ChangeTStringToTDateString(frm090613.Txt1(26)) & "~" & ChangeTStringToTDateString(frm090613.Txt1(27)) & "　　會稿"
End If
'2015/8/25 END
StrMenu
SetGrd1
End Sub

Sub StrMenu()
'StrSQL = "SELECT nvl(ST02,R109001),round(SUM(R109002)/SUM(R109003),2),SUM(R109003),round(SUM(R109004)/SUM(R109005),2),SUM(R109005),round(SUM(R109006)/SUM(R109007),2),SUM(R109007),round(SUM(R109008)/SUM(R109009),2),SUM(R109009),round(SUM(R109010)/SUM(R109011),2),SUM(R109011),round(SUM(R109012)/SUM(R109013),2),SUM(R109013),round(SUM(R109014)/SUM(R109015),2),SUM(R109015),R109001 FROM R090613,staff WHERE ID='" & strUserNum & "' and R109001=st01(+) and (r109016 is null or r109016='') GROUP BY R109001,nvl(ST02,R109001) ORDER BY R109001,nvl(ST02,R109001) "
'900821 邱小姐說錯誤資料也算
strSql = "SELECT nvl(ST02,R109001),DECODE(SUM(R109003),0,0,null,0,round(SUM(R109002)/SUM(R109003),2)),SUM(R109003),DECODE(SUM(R109005),0,0,null,0,round(SUM(R109004)/SUM(R109005),2)),SUM(R109005),DECODE(SUM(R109007),0,0,null,0,round(SUM(R109006)/SUM(R109007),2)),SUM(R109007),DECODE(SUM(R109009),0,0,null,0,round(SUM(R109008)/SUM(R109009),2)),SUM(R109009),DECODE(SUM(R109011),0,0,null,0,round(SUM(R109010)/SUM(R109011),2)),SUM(R109011),decode(SUM(R109013),0,0,null,0,round(SUM(R109012)/SUM(R109013),2)),SUM(R109013),decode(SUM(R109015),0,0,null,0,round(SUM(R109014)/SUM(R109015),2)),SUM(R109015),R109001 FROM R090613,staff WHERE ID='" & strUserNum & "' and R109001=st01(+) GROUP BY R109001,nvl(ST02,R109001) ORDER BY R109001,nvl(ST02,R109001) "
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
'add by nickc 2006/05/03 加入合計
Dim Cal1 As Double, Cal2 As Double, Cal3 As Double, Cal4 As Double, Cal5 As Double, Cal6 As Double, Cal7 As Double
Dim Cal8 As Double, Cal9 As Double, Cal10 As Double, Cal11 As Double, Cal12 As Double, Cal13 As Double, Cal14 As Double
Dim Cali As Integer
Cal1 = 0: Cal2 = 0: Cal3 = 0: Cal4 = 0: Cal5 = 0: Cal6 = 0: Cal7 = 0: Cal8 = 0: Cal9 = 0: Cal10 = 0: Cal11 = 0: Cal12 = 0: Cal13 = 0: Cal14 = 0
For Cali = 0 To grd1.Rows - 1
    Cal1 = Cal1 + (Val(grd1.TextMatrix(Cali, 1)) * Val(grd1.TextMatrix(Cali, 2)))
    Cal2 = Cal2 + Val(grd1.TextMatrix(Cali, 2))
    Cal3 = Cal3 + (Val(grd1.TextMatrix(Cali, 3)) * Val(grd1.TextMatrix(Cali, 4)))
    Cal4 = Cal4 + Val(grd1.TextMatrix(Cali, 4))
    Cal5 = Cal5 + (Val(grd1.TextMatrix(Cali, 5)) * Val(grd1.TextMatrix(Cali, 6)))
    Cal6 = Cal6 + Val(grd1.TextMatrix(Cali, 6))
    Cal7 = Cal7 + (Val(grd1.TextMatrix(Cali, 7)) * Val(grd1.TextMatrix(Cali, 8)))
    Cal8 = Cal8 + Val(grd1.TextMatrix(Cali, 8))
    Cal9 = Cal9 + (Val(grd1.TextMatrix(Cali, 9)) * Val(grd1.TextMatrix(Cali, 10)))
    Cal10 = Cal10 + Val(grd1.TextMatrix(Cali, 10))
    Cal11 = Cal11 + (Val(grd1.TextMatrix(Cali, 11)) * Val(grd1.TextMatrix(Cali, 12)))
    Cal12 = Cal12 + Val(grd1.TextMatrix(Cali, 12))
    Cal13 = Cal13 + (Val(grd1.TextMatrix(Cali, 13)) * Val(grd1.TextMatrix(Cali, 14)))
    Cal14 = Cal14 + Val(grd1.TextMatrix(Cali, 14))
Next Cali
grd1.Rows = grd1.Rows + 1
grd1.TextMatrix(grd1.Rows - 1, 0) = "合計"
If Cal2 <> 0 Then
    grd1.TextMatrix(grd1.Rows - 1, 1) = Format(Cal1 / Cal2, "0.00")
Else
    grd1.TextMatrix(grd1.Rows - 1, 1) = "0.00"
End If
grd1.TextMatrix(grd1.Rows - 1, 2) = Cal2
If Cal4 <> 0 Then
    grd1.TextMatrix(grd1.Rows - 1, 3) = Format(Cal3 / Cal4, "0.00")
Else
    grd1.TextMatrix(grd1.Rows - 1, 3) = "0.00"
End If
grd1.TextMatrix(grd1.Rows - 1, 4) = Cal4
If Cal6 <> 0 Then
    grd1.TextMatrix(grd1.Rows - 1, 5) = Format(Cal5 / Cal6, "0.00")
Else
    grd1.TextMatrix(grd1.Rows - 1, 5) = "0.00"
End If
grd1.TextMatrix(grd1.Rows - 1, 6) = Cal6
If Cal8 <> 0 Then
    grd1.TextMatrix(grd1.Rows - 1, 7) = Format(Cal7 / Cal8, "0.00")
Else
    grd1.TextMatrix(grd1.Rows - 1, 7) = "0.00"
End If
grd1.TextMatrix(grd1.Rows - 1, 8) = Cal8
If Cal10 <> 0 Then
    grd1.TextMatrix(grd1.Rows - 1, 9) = Format(Cal9 / Cal10, "0.00")
Else
    grd1.TextMatrix(grd1.Rows - 1, 9) = "0.00"
End If
grd1.TextMatrix(grd1.Rows - 1, 10) = Cal10
If Cal12 <> 0 Then
    grd1.TextMatrix(grd1.Rows - 1, 11) = Format(Cal11 / Cal12, "0.00")
Else
    grd1.TextMatrix(grd1.Rows - 1, 11) = "0.00"
End If
grd1.TextMatrix(grd1.Rows - 1, 12) = Cal12
If Cal14 <> 0 Then
    grd1.TextMatrix(grd1.Rows - 1, 13) = Format(Cal13 / Cal14, "0.00")
Else
    grd1.TextMatrix(grd1.Rows - 1, 13) = "0.00"
End If
grd1.TextMatrix(grd1.Rows - 1, 14) = Cal14
End Sub

Private Sub SetGrd1()
With grd1
    .Cols = 16
    .row = 0
    .col = 0
    If Val(frm090613.Txt1(9)) = 1 Then
        .Text = "承辦人"
    Else
        .Text = "智權人員"
    End If
    .ColWidth(0) = 800
    'edit by nickc 改靠右
    '.CellAlignment = flexAlignCenterCenter
    .ColAlignment(0) = flexAlignCenterCenter
    For i = 0 To 6
        .col = (i * 2) + 1: .Text = "平均天數"
        .ColWidth((i * 2) + 1) = 770
        'edit by nickc 改靠右
        '.CellAlignment = flexAlignCenterCenter
        .ColAlignment((i * 2) + 1) = flexAlignRightCenter
        .col = (i * 2) + 2:   .Text = "件數"
        .ColWidth((i * 2) + 2) = 400
        'edit by nickc 改靠右
        '.CellAlignment = flexAlignCenterCenter
        .ColAlignment((i * 2) + 2) = flexAlignRightCenter
    Next i
    .col = 15
    .Text = ""
    .ColWidth(15) = 0
    '.CellAlignment = flexAlignCenterCenter
    .ColAlignment(15) = flexAlignRightCenter
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090613_1 = Nothing
End Sub

'Add By Sindy 2015/9/30
Private Sub Grd1_Click()
   Dim iRow As Integer, iCol As Integer
   With grd1
      iRow = .MouseRow
      iCol = .MouseCol
      If iRow > 0 Then
         Select Case iCol
            Case 2, 4, 6, 8, 10, 12, 14
               .Enabled = False
               GetStatistic iRow, iCol
               .Enabled = True
         End Select
      End If
   End With
End Sub
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim iRow As Integer, iCol As Integer, lBackColor As Long
   
   With grd1
      iRow = .MouseRow
      iCol = .MouseCol
      If iRow < 0 Or iCol < 0 Then Exit Sub
      If iRow = m_iRow And iCol = m_iCol Then
         Exit Sub
      End If
      If m_iCol <> 0 Then
         .row = m_iRow: .col = m_iCol
'         If m_iCol < 3 Then
'            .CellForeColor = .ForeColorFixed
'            .CellBackColor = .BackColorFixed
'         Else
            .CellForeColor = .ForeColor
            .CellBackColor = .BackColor
'         End If
         m_iRow = 0: m_iCol = 0
      End If
      
      If iRow > 0 Then
         Select Case iCol
            Case 2, 4, 6, 8, 10, 12, 14
               .row = iRow: .col = iCol
               lBackColor = .CellBackColor
               .CellBackColor = .CellForeColor
               .CellForeColor = lBackColor
               m_iCol = .col
               m_iRow = .row
            Case Else
         End Select
      End If
   End With
End Sub
'案件明細
Private Sub GetStatistic(p_iRow As Integer, p_iCol As Integer)
Dim stCon As String
Dim strQuyEmpID As String
Dim ii As Integer
   
   strQuyEmpID = grd1.TextMatrix(p_iRow, 15) '智權人員或承辦人ID
   
   If grd1.TextMatrix(p_iRow, 0) <> "合計" Then
      stCon = " and R109001='" & strQuyEmpID & "'"
   End If
   Select Case p_iCol
      Case 2
         stCon = stCon & " and R109003>0"
      Case 4
         stCon = stCon & " and R109005>0"
      Case 6
         stCon = stCon & " and R109007>0"
      Case 8
         stCon = stCon & " and R109009>0"
      Case 10
         stCon = stCon & " and R109011>0"
      Case 12
         stCon = stCon & " and R109013>0"
      Case 14
         stCon = stCon & " and R109015>0"
   End Select
   
   strExc(0) = "SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(N1.NA03,N1.NA04) As 申請國家,DECODE(PA09,'020',PTM04,PTM03) As 種類,nvl(DECODE(PA09,'000',cpm03,cpm04),cp10) AS 案件性質,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,SUBSTR(' '||sqldatet(CP48),-9) AS 承辦期限,SUBSTR(' '||sqldatet(EP06),-9) AS 齊備日,SUBSTR(' '||sqldatet(EP09),-9) AS 完稿日,SUBSTR(' '||sqldatet(EP28),-9) AS 預會日,SUBSTR(' '||sqldatet(EP07),-9) AS 會稿日" & _
               ",nvl(S3.ST02,ep04) AS 核稿人,SUBSTR(' '||sqldatet(EP08),-9) AS 會稿完成日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,Nvl(EP35,0) AS 承辦天數,EP12 AS 承辦備註,nvl(s2.st02,CP14) AS 承辦人,nvl(s1.st02,CP13) AS 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
               " FROM R090613,ENGINEERPROGRESS,PATENTTRADEMARKMAP,CASEPROGRESS,patent,staff s1,staff s2,staff s3,casepropertymap,Nation N1" & _
               " WHERE ID='" & strUserNum & "'" & stCon & " and R109017=cp09(+) and CP09=EP02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)" & _
               " and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04" & _
               " and cp13=s1.st01(+) and cp14=s2.st01(+) and ep04=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And PA09=N1.NA01(+) "
   strExc(0) = strExc(0) & " union SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(tm05,NVL(tm06,tm07)) AS 案件名稱,nvl(N1.NA03,N1.NA04) As 申請國家,DECODE(tm10,'020',PTM04,PTM03) As 種類,nvl(DECODE(tm10,'000',cpm03,cpm04),cp10) AS 案件性質,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,SUBSTR(' '||sqldatet(CP48),-9) AS 承辦期限,SUBSTR(' '||sqldatet(EP06),-9) AS 齊備日,SUBSTR(' '||sqldatet(EP09),-9) AS 完稿日,SUBSTR(' '||sqldatet(EP28),-9) AS 預會日,SUBSTR(' '||sqldatet(EP07),-9) AS 會稿日" & _
               ",nvl(S3.ST02,ep04) AS 核稿人,SUBSTR(' '||sqldatet(EP08),-9) AS 會稿完成日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,Nvl(EP35,0) AS 承辦天數,EP12 AS 承辦備註,nvl(s2.st02,CP14) AS 承辦人,nvl(s1.st02,CP13) AS 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
               " FROM R090613,ENGINEERPROGRESS,PATENTTRADEMARKMAP,CASEPROGRESS,Trademark,staff s1,staff s2,staff s3,casepropertymap,Nation N1" & _
               " WHERE ID='" & strUserNum & "'" & stCon & " and R109017=cp09(+) and CP09=EP02(+) AND '2'=PTM01(+) AND tm08=PTM02(+)" & _
               " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04" & _
               " and cp13=s1.st01(+) and cp14=s2.st01(+) and ep04=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And tm10=N1.NA01(+) "
   strExc(0) = strExc(0) & " union SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(lc05,NVL(lc06,lc07)) AS 案件名稱,nvl(N1.NA03,N1.NA04) As 申請國家,' ' As 種類,nvl(DECODE(lc15,'000',cpm03,cpm04),cp10) AS 案件性質,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,SUBSTR(' '||sqldatet(CP48),-9) AS 承辦期限,SUBSTR(' '||sqldatet(EP06),-9) AS 齊備日,SUBSTR(' '||sqldatet(EP09),-9) AS 完稿日,SUBSTR(' '||sqldatet(EP28),-9) AS 預會日,SUBSTR(' '||sqldatet(EP07),-9) AS 會稿日" & _
               ",nvl(S3.ST02,ep04) AS 核稿人,SUBSTR(' '||sqldatet(EP08),-9) AS 會稿完成日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,Nvl(EP35,0) AS 承辦天數,EP12 AS 承辦備註,nvl(s2.st02,CP14) AS 承辦人,nvl(s1.st02,CP13) AS 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
               " FROM R090613,ENGINEERPROGRESS,CASEPROGRESS,lawcase,staff s1,staff s2,staff s3,casepropertymap,Nation N1" & _
               " WHERE ID='" & strUserNum & "'" & stCon & " and R109017=cp09(+) and CP09=EP02(+)" & _
               " and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04" & _
               " and cp13=s1.st01(+) and cp14=s2.st01(+) and ep04=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And lc15=N1.NA01(+) "
   strExc(0) = strExc(0) & " union SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,hc06 AS 案件名稱,'台灣' As 申請國家,' ' As 種類,nvl(cpm03,cp10) AS 案件性質,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,SUBSTR(' '||sqldatet(CP48),-9) AS 承辦期限,SUBSTR(' '||sqldatet(EP06),-9) AS 齊備日,SUBSTR(' '||sqldatet(EP09),-9) AS 完稿日,SUBSTR(' '||sqldatet(EP28),-9) AS 預會日,SUBSTR(' '||sqldatet(EP07),-9) AS 會稿日" & _
               ",nvl(S3.ST02,ep04) AS 核稿人,SUBSTR(' '||sqldatet(EP08),-9) AS 會稿完成日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,Nvl(EP35,0) AS 承辦天數,EP12 AS 承辦備註,nvl(s2.st02,CP14) AS 承辦人,nvl(s1.st02,CP13) AS 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
               " FROM R090613,ENGINEERPROGRESS,CASEPROGRESS,hirecase,staff s1,staff s2,staff s3,casepropertymap" & _
               " WHERE ID='" & strUserNum & "'" & stCon & " and R109017=cp09(+) and CP09=EP02(+)" & _
               " and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04" & _
               " and cp13=s1.st01(+) and cp14=s2.st01(+) and ep04=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) "
   strExc(0) = strExc(0) & " union SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(sp05,NVL(sp06,sp07)) AS 案件名稱,nvl(N1.NA03,N1.NA04) As 申請國家,' ' As 種類,nvl(DECODE(sp09,'000',cpm03,cpm04),cp10) AS 案件性質,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,SUBSTR(' '||sqldatet(CP48),-9) AS 承辦期限,SUBSTR(' '||sqldatet(EP06),-9) AS 齊備日,SUBSTR(' '||sqldatet(EP09),-9) AS 完稿日,SUBSTR(' '||sqldatet(EP28),-9) AS 預會日,SUBSTR(' '||sqldatet(EP07),-9) AS 會稿日" & _
               ",nvl(S3.ST02,ep04) AS 核稿人,SUBSTR(' '||sqldatet(EP08),-9) AS 會稿完成日,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,Nvl(EP35,0) AS 承辦天數,EP12 AS 承辦備註,nvl(s2.st02,CP14) AS 承辦人,nvl(s1.st02,CP13) AS 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & _
               " FROM R090613,ENGINEERPROGRESS,CASEPROGRESS,servicepractice,staff s1,staff s2,staff s3,casepropertymap,Nation N1" & _
               " WHERE ID='" & strUserNum & "'" & stCon & " and R109017=cp09(+) and CP09=EP02(+)" & _
               " and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04" & _
               " and cp13=s1.st01(+) and cp14=s2.st01(+) and ep04=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And sp09=N1.NA01(+) "
   strExc(0) = strExc(0) + " ORDER BY 收文日,FSort"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With frm090613_3
         'Added by Morgan 2024/3/15
         If Left(Pub_StrUserSt03, 2) = "P1" And frm090613.Option1(4).Value And p_iCol = 6 Then
            .cmdOK(1).Visible = True
         End If
         'end 2024/3/15
         
         .grdDataList.Visible = False
         Set .grdDataList.Recordset = RsTemp.Clone: DoEvents
         .SetDataListWidth
         For ii = 1 To .grdDataList.Rows - 1
            .grdDataList.TextMatrix(ii, 5) = .grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(.grdDataList.TextMatrix(ii, 18), "1")
            '收款情形
            Dim IntTemp1 As Long
            Dim IntTemp2 As Long
            IntTemp1 = 0
            IntTemp2 = 0
            .grdDataList.row = ii
            .grdDataList.col = 19
            If Not IsNull(.grdDataList.Text) And .grdDataList.Text <> "" Then
               If Mid(.grdDataList.Text, 1, 1) = "X" Then
                  strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0 FROM ACC1K0 WHERE A1K01='" & .grdDataList.Text & "'"
                  CheckOC2
                  adoRecordset1.CursorLocation = adUseClient
                  adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                     If Not IsNull(adoRecordset1.Fields(0)) Then
                        IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
                     End If
                     If Not IsNull(adoRecordset1.Fields(1)) Then
                        IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
                     End If
                     If Not IsNull(adoRecordset1.Fields(4)) Then
                        IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
                     End If
                     If Not IsNull(adoRecordset1.Fields(5)) Then
                        IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
                     End If
                     If IntTemp1 = IntTemp2 Then
                        .grdDataList.Text = "收回"
                     Else
                        If IntTemp2 = 0 Then
                           .grdDataList.Text = "未收"
                        Else
                           If IntTemp1 > IntTemp2 Then
                              .grdDataList.Text = "部分收回"
                           End If
                        End If
                     End If
                  End If
               End If
            End If
            CheckOC2
            DoEvents
         Next ii
         .grdDataList.Visible = True
         '.Show vbModal
         .Show
         Me.Hide
      End With
   Else
      MsgBox "無資料！"
   End If
End Sub
'2015/9/30 END
