VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040204a 
   BorderStyle     =   1  '單線固定
   Caption         =   "審查委員准駁統計"
   ClientHeight    =   5730
   ClientLeft      =   105
   ClientTop       =   990
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton Cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8448
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7224
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   4560
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   1110
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   8043
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   885
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   1110
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1561
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   2
      Left            =   1080
      TabIndex        =   7
      Top             =   900
      Width           =   3720
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   1068
      TabIndex        =   6
      Top             =   660
      Width           =   3720
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   432
      Width           =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "准駁日："
      Height          =   180
      Index           =   2
      Left            =   108
      TabIndex        =   2
      Top             =   900
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   108
      TabIndex        =   1
      Top             =   660
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   0
      Top             =   432
      Width           =   912
   End
End
Attribute VB_Name = "frm040204a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ; Grd1(index)改字型=新細明體-ExtB
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strSql As String, StrTest As String, strTemp As Variant, i As Integer, j As Integer, s As Integer, k As Integer
Dim intK As Integer, IntTotle(0 To 4) As Integer, TxtCheckTOrP As String, strSQL1 As String, strSQL2 As String, BolAddData As Boolean

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Unload Me
     frm040204.Show
Case 1
     Unload Me
     Unload frm040204
Case Else
End Select
End Sub

Private Sub Form_Activate()
strTemp = Split(frm040204.txt1(0), ",")
Select Case strTemp(0)
    Case "CFP", "P", "FCP"
          TxtCheckTOrP = "P"
          'Add by Morgan 2004/4/13
          Grd1(1).Rows = 3
          Grd1(1).FixedRows = 2
          Grd1(1).MergeCells = flexMergeRestrictRows
          Grd1(1).MergeRow(0) = True
          'Add end
          SetGridWidth
          StrMenu
    Case "T", "FCT", "CFT", "TF"
          TxtCheckTOrP = "T"
          'Add by Morgan 2004/4/13
          Grd1(1).Rows = 3
          Grd1(1).FixedRows = 2
          Grd1(1).MergeCells = flexMergeRestrictRows
          Grd1(1).MergeRow(0) = True
          'Add end
          SetGridWidth
          StrMenu1
    Case Else
          s = MsgBox("系統類別錯誤", , "USER 輸入之系統有誤")
          frm040204.txt1(0).SetFocus
          frm040204.txt1(0).SelStart = 0
          frm040204.txt1(0).SelLength = Len(frm040204.txt1(0))
          frm040204.Show
          Unload Me
          Exit Sub
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Sub StrMenu()         '處理專利
Me.Enabled = False
'顯示表單上方
lbl1(0).Caption = frm040204.txt1(0)
lbl1(1).Caption = frm040204.txt1(3) + "-" + frm040204.txt1(4)
lbl1(2).Caption = frm040204.txt1(1) + "-" + frm040204.txt1(2)
strSQL1 = ""
If Len(frm040204.txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & frm040204.txt1(3) & "' "
End If
If Len(frm040204.txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & frm040204.txt1(4) & "' "
End If
If Len(frm040204.txt1(3)) <> 0 Or Len(frm040204.txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm040204.lbl1(2) & frm040204.txt1(3) & "-" & frm040204.txt1(4) 'Add By Sindy 2010/9/28
End If
If Len(Trim(frm040204.txt1(0))) <> 0 Then
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(frm040204.txt1(0), 1) & ") "
   pub_QL05 = pub_QL05 & ";" & frm040204.lbl1(0) & frm040204.txt1(0) 'Add By Sindy 2010/9/28
End If
If Len(Trim(frm040204.txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP25>=" & Val(ChangeTStringToWString(frm040204.txt1(1))) & " "
End If
If Len(Trim(frm040204.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP25<=" & Val(ChangeTStringToWString(frm040204.txt1(2))) & " "
End If
If Len(Trim(frm040204.txt1(1))) <> 0 Or Len(Trim(frm040204.txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm040204.lbl1(1) & frm040204.txt1(1) & "-" & frm040204.txt1(2) 'Add By Sindy 2010/9/28
End If
strSql = strSql & " and cp10 in ('801', '802', '803', '804', '501', '502', '503', '504', '101', '102' , '103', '104', '105','125') "
'運算
strSql = " SELECT CP01,CP35,CP10,CP24,pa08,CP02,CP03,CP04 FROM CASEPROGRESS,patent WHERE CP35 IS NOT NULL and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & strSQL1
'StrSQL = StrSQL + " AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/28
Else
    CheckOC
    Me.Enabled = True
    InsertQueryLog (0) 'Add By Sindy 2010/9/28
    ShowNoData
    frm040204.Show
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End If
'計算統計
adoRecordset.MoveFirst
intK = adoRecordset.RecordCount
j = 0
Grd1(1).Visible = False
Grd1(0).Visible = False
Do While adoRecordset.EOF = False
    If Not IsNull(adoRecordset.Fields(1)) Then
        strTemp = Split(adoRecordset.Fields(1), ",")
    Else
        strTemp = ""
    End If
    For i = 0 To UBound(strTemp)
        Grd1(1).col = 0
        'Modify by Morgan 2004/2/13
'        If grd1(1).Rows = 2 And Len(grd1(1).Text) = 0 Then
'            grd1(1).Row = 1
        If Grd1(1).Rows = 3 And Len(Grd1(1).Text) = 0 Then
            Grd1(1).row = 2
            Grd1(1).col = 0
            Grd1(1).Text = strTemp(i)
            StrMenu3
        Else
            BolAddData = False
            'Modify by Morgan 2004/2/13
            'For k = 1 To grd1(1).Rows - 1
            For k = 2 To Grd1(1).Rows - 1
                Grd1(1).row = k
                Grd1(1).col = 0
                If strTemp(i) = Grd1(1).Text Then
                    StrMenu3
                    BolAddData = True
                    Exit For
                End If
            Next k
            If BolAddData = False Then
               Grd1(1).Rows = Grd1(1).Rows + 1
               Grd1(1).row = Grd1(1).Rows - 1
               Grd1(1).Text = strTemp(i)
               StrMenu3
            End If
            'Grd1(0).Refresh
            'Grd1(1).Refresh
        End If
    Next i
    DoEvents
    adoRecordset.MoveNext
Loop
CheckOC
SetGridWidth
'計算核准率與總件數和排序
StrMenu4
Grd1(1).Visible = True
Grd1(0).Visible = True
Me.Enabled = True

End Sub
Sub StrMenu4()      '計算核准率與總件數和排序           專利
With Grd1(1)
    '.Visible = False
    'Modify by Morgan 2004/4/13
    'For i = 1 To .Rows - 1
    For i = 2 To .Rows - 1
        IntTotle(0) = 0       '總件數准
        IntTotle(1) = 0       '總件數駁
        IntTotle(2) = 0       '單項准
        IntTotle(3) = 0       '單項駁
        .row = i
        .col = 0
        .CellAlignment = flexAlignRightCenter
        .col = 1
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 2
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 3
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) <> 0 Then
            .Text = Format(str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100), "###.00") + " %"
            IntTotle(0) = IntTotle(0) + IntTotle(2)
            IntTotle(1) = IntTotle(1) + IntTotle(3)
        Else
            .Text = "0 %"
        End If
        .col = 4
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 5
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 6
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) <> 0 Then
            .Text = Format(str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100), "###.00") + " %"
            IntTotle(0) = IntTotle(0) + IntTotle(2)
            IntTotle(1) = IntTotle(1) + IntTotle(3)
        Else
             .Text = "0 %"
        End If
        .col = 7
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 8
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 9
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) <> 0 Then
            .Text = Format(str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100), "###.00") + " %"
            IntTotle(0) = IntTotle(0) + IntTotle(2)
            IntTotle(1) = IntTotle(1) + IntTotle(3)
        Else
            .Text = "0 %"
        End If
        .col = 10
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 11
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 12
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) <> 0 Then
            .Text = Format(str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100), "###.00") + " %"
            IntTotle(0) = IntTotle(0) + IntTotle(2)
            IntTotle(1) = IntTotle(1) + IntTotle(3)
        Else
            .Text = "0 %"
        End If
        .col = 13
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 14
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 15
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) <> 0 Then
            .Text = Format(str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100), "###.00") + " %"
            IntTotle(0) = IntTotle(0) + IntTotle(2)
            IntTotle(1) = IntTotle(1) + IntTotle(3)
        Else
            .Text = "0 %"
        End If
        .col = 16
        .CellAlignment = flexAlignRightCenter
        .Text = str(IntTotle(0))
        .col = 17
        .CellAlignment = flexAlignRightCenter
        .Text = str(IntTotle(1))
        .col = 18
        .CellAlignment = flexAlignRightCenter
        'Modified by Morgan 2018/5/7
        'If IntTotle(2) + IntTotle(3) <> 0 Then
        If IntTotle(0) + IntTotle(1) <> 0 Then
            .Text = Format(str((IntTotle(0) / (IntTotle(0) + IntTotle(1))) * 100), "###.00") + " %"
        Else
            .Text = "0 %"
        End If
    Next i
    '排序
    'Modify by Morgan 2004/4/13
    '.Row = 1
    .row = 2
    .RowSel = .Rows - 1
    .col = 0
    .ColSel = 0
    .Sort = 5
    .Visible = True
End With
End Sub
Sub StrMenu3()      '判斷後再加入GRID      專利
If Not IsNull(adoRecordset.Fields(2)) Then
    Select Case Val(adoRecordset.Fields(2))
    Case 101
            If Not IsNull(adoRecordset.Fields(3)) Then
                Select Case adoRecordset.Fields(3)
                    Case 1
                        Grd1(1).col = 1
                        Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case 2
                        Grd1(1).col = 2
                        Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case Else
                End Select
            End If
    Case 102
            If Not IsNull(adoRecordset.Fields(3)) Then
                Select Case adoRecordset.Fields(3)
                    Case 1
                            Grd1(1).col = 4
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case 2
                            Grd1(1).col = 5
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case Else
                End Select
            End If
    Case 103, 105
            If Not IsNull(adoRecordset.Fields(3)) Then
                Select Case adoRecordset.Fields(3)
                    Case 1
                            Grd1(1).col = 7
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case 2
                            Grd1(1).col = 8
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case Else
                End Select
            End If
    Case 501, 502, 503, 504
            If Not IsNull(adoRecordset.Fields(3)) Then
                Select Case adoRecordset.Fields(3)
                    Case 1
                            Grd1(1).col = 10
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case 2
                                 Grd1(1).col = 11
                                 Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case Else
                End Select
            End If
    Case 801, 802, 803, 804
            If Not IsNull(adoRecordset.Fields(3)) Then
                Select Case adoRecordset.Fields(3)
                    Case 1
                            Grd1(1).col = 13
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case 2
                            Grd1(1).col = 14
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case Else
                End Select
            End If
    Case 104
            If Not IsNull(adoRecordset.Fields(4)) Then
                Select Case Val(adoRecordset.Fields(4))
                    Case 1
                            If Not IsNull(adoRecordset.Fields(3)) Then
                                Select Case adoRecordset.Fields(3)
                                    Case 1
                                            Grd1(1).col = 1
                                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                    Case 2
                                            Grd1(1).col = 2
                                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                Case Else
                                End Select
                            End If
                    Case 2
                            If Not IsNull(adoRecordset.Fields(3)) Then
                                 Select Case adoRecordset.Fields(3)
                                    Case 1
                                            Grd1(1).col = 4
                                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                    Case 2
                                            Grd1(1).col = 5
                                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                    Case Else
                                End Select
                            End If
                    Case 3, 4
                            If Not IsNull(adoRecordset.Fields(3)) Then
                                 Select Case adoRecordset.Fields(3)
                                    Case 1
                                            Grd1(1).col = 7
                                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                    Case 2
                                            Grd1(1).col = 8
                                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                    Case Else
                                 End Select
                            End If
                    Case Else
                End Select
            End If
    Case Else
    End Select
End If
End Sub
Sub StrMenu5()            '判斷後再加入GRID          商標
If Not IsNull(adoRecordset.Fields(2)) Then
    Select Case Val(adoRecordset.Fields(2))
    Case 101
            If Not IsNull(adoRecordset.Fields(4)) Then
                Select Case adoRecordset.Fields(4)
                    Case 1
                         If Not IsNull(adoRecordset.Fields(3)) Then
                            Select Case adoRecordset.Fields(3)
                                Case 1
                                     Grd1(1).col = 1
                                     Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                Case 2
                                     Grd1(1).col = 2
                                     Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                Case Else
                            End Select
                         End If
                    Case 2
                         If Not IsNull(adoRecordset.Fields(3)) Then
                            Select Case adoRecordset.Fields(3)
                                Case 1
                                     Grd1(1).col = 4
                                     Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                Case 2
                                     Grd1(1).col = 5
                                     Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                                Case Else
                            End Select
                         End If
                    Case Else
                End Select
            End If
    Case 401, 402, 403, 404, 405
            If Not IsNull(adoRecordset.Fields(3)) Then
                Select Case adoRecordset.Fields(3)
                    Case 1
                            Grd1(1).col = 7
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case 2
                            Grd1(1).col = 8
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case Else
                End Select
            End If
    'modify by sonia 2017/9/5 +623,624
    Case 601, 602, 603, 604, 605, 606, 623, 624
            If Not IsNull(adoRecordset.Fields(3)) Then
                Select Case adoRecordset.Fields(3)
                    Case 1
                            Grd1(1).col = 10
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case 2
                            Grd1(1).col = 11
                            Grd1(1).Text = str(Val(Grd1(1).Text) + 1)
                    Case Else
                End Select
            End If
    Case Else
    End Select
End If
End Sub
Sub StrMenu6()      '計算核准率與總件數和排序           商標
With Grd1(1)
    '.Visible = False
    'Modify by Morgan 2004/4/13
    'For i = 1 To .Rows - 1
    For i = 2 To .Rows - 1
    'Modify end
        IntTotle(0) = 0       '總件數准
        IntTotle(1) = 0       '總件數駁
        IntTotle(2) = 0       '單項准
        IntTotle(3) = 0       '單項駁
        .row = i
        .col = 0
        .CellAlignment = flexAlignRightCenter
        .col = 1
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 2
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 3
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) = 0 Then
          .Text = "0%"
        Else
         .Text = str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100) + "%"
        End If
        IntTotle(0) = IntTotle(0) + IntTotle(2)
        IntTotle(1) = IntTotle(1) + IntTotle(3)
        .col = 4
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 5
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 6
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) = 0 Then
          .Text = "0%"
        Else
         .Text = str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100) + "%"
        End If
        IntTotle(0) = IntTotle(0) + IntTotle(2)
        IntTotle(1) = IntTotle(1) + IntTotle(3)
        .col = 7
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 8
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 9
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) = 0 Then
          .Text = "0%"
        Else
         .Text = str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100) + "%"
        End If
        IntTotle(0) = IntTotle(0) + IntTotle(2)
        IntTotle(1) = IntTotle(1) + IntTotle(3)
        .col = 10
        .CellAlignment = flexAlignRightCenter
        IntTotle(2) = Val(.Text)
        .col = 11
        .CellAlignment = flexAlignRightCenter
        IntTotle(3) = Val(.Text)
        .col = 12
        .CellAlignment = flexAlignRightCenter
        If IntTotle(2) + IntTotle(3) = 0 Then
          .Text = "0%"
        Else
         .Text = str((IntTotle(2) / (IntTotle(2) + IntTotle(3))) * 100) + "%"
        End If
        IntTotle(0) = IntTotle(0) + IntTotle(2)
        IntTotle(1) = IntTotle(1) + IntTotle(3)
        .col = 13
        .CellAlignment = flexAlignRightCenter
        .Text = str(IntTotle(0))
        .col = 14
        .CellAlignment = flexAlignRightCenter
        .Text = str(IntTotle(1))
        .col = 15
        .CellAlignment = flexAlignRightCenter
        If IntTotle(0) + IntTotle(1) = 0 Then
          .Text = "0%"
        Else
         .Text = str((IntTotle(0) / (IntTotle(0) + IntTotle(1))) * 100) + "%"
        End If
        DoEvents
    Next i
    '排序
    'Modify by Morgan 2004/4/13
    '.Row = 1
    .row = 2
    'Modify end
    .RowSel = .Rows - 1
    .col = 0
    .ColSel = 0
    .Sort = 5
    .Visible = True
End With
End Sub

Sub StrMenu1()         '處理商標

Me.Enabled = False
'顯示表單上方
lbl1(0).Caption = frm040204.txt1(0)
lbl1(1).Caption = frm040204.txt1(3) + "-" + frm040204.txt1(4)
lbl1(2).Caption = frm040204.txt1(1) + "-" + frm040204.txt1(2)

lbl1(0).Caption = frm040204.txt1(0)
lbl1(1).Caption = frm040204.txt1(3) + "-" + frm040204.txt1(4)
lbl1(2).Caption = frm040204.txt1(1) + "-" + frm040204.txt1(2)
strSQL1 = ""
If Len(frm040204.txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & frm040204.txt1(3) & "' "
End If
If Len(frm040204.txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & frm040204.txt1(4) & "' "
End If
If Len(frm040204.txt1(3)) <> 0 Or Len(frm040204.txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm040204.lbl1(2) & frm040204.txt1(3) & "-" & frm040204.txt1(4) 'Add By Sindy 2010/9/28
End If
If Len(Trim(frm040204.txt1(0))) <> 0 Then
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(frm040204.txt1(0), 2) & ") "
   pub_QL05 = pub_QL05 & ";" & frm040204.lbl1(0) & frm040204.txt1(0) 'Add By Sindy 2010/9/28
End If
If Len(Trim(frm040204.txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP25>=" & Val(ChangeTStringToWString(frm040204.txt1(1))) & " "
End If
If Len(Trim(frm040204.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP25<=" & Val(ChangeTStringToWString(frm040204.txt1(2))) & " "
End If
If Len(Trim(frm040204.txt1(1))) <> 0 Or Len(Trim(frm040204.txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm040204.lbl1(1) & frm040204.txt1(1) & "-" & frm040204.txt1(2) 'Add By Sindy 2010/9/28
End If
'strSQL1 = strSQL & " and (cp10 = '101' or (cp10 >=601 and cp10<=610) or (cp10>='400' and cp10<='4zz')) "
'modify by sonia 2017/8/30 +623,624
strSQL1 = strSQL1 & " and cp10 in ('601', '602', '603', '604', '605', '606', '623', '624', '401', '402', '403', '404', '405','101') and tm08>=1 and tm08<=2 "
'運算
strSql = " SELECT CP01,CP35,CP10,CP24,tm08,CP02,CP03,CP04 FROM CASEPROGRESS,trademark WHERE CP35 IS NOT NULL and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  " & strSQL1
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/28
Else
    CheckOC
    Me.Enabled = True
    InsertQueryLog (0) 'Add By Sindy 2010/9/28
    ShowNoData
    frm040204.Show
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End If
'計算統計
adoRecordset.MoveFirst
Grd1(1).Clear

'Remove by Morgan 2004/4/13
'Grd1(1).Rows = 2
'Grd1(0).Clear
'Grd1(0).Rows = 2
'Remove end

intK = adoRecordset.RecordCount
j = 0
adoRecordset.MoveFirst
Grd1(1).Visible = False
Grd1(0).Visible = False
Do While adoRecordset.EOF = False
    If Not IsNull(adoRecordset.Fields(1)) Then
        strTemp = Split(adoRecordset.Fields(1), ",")
    Else
        strTemp = ""
    End If
    For i = 0 To UBound(strTemp)
        Grd1(1).col = 0
        'Modify by Morgan 2004/4/13
        'Grd1(1).Row = 1
        'If Grd1(1).Rows = 2 And Len(Grd1(1).Text) = 0 Then
        '    Grd1(1).Row = 1
        If Grd1(1).Rows = 3 And Len(Grd1(1).Text) = 0 Then
            Grd1(1).row = 2
        'Modify end
        
            Grd1(1).col = 0
            Grd1(1).Text = strTemp(i)
            StrMenu5
        Else
            BolAddData = False
            'Modify by Morgan 2004/4/13
            'For k = 1 To Grd1(1).Rows - 1
            For k = 2 To Grd1(1).Rows - 1
            'Modify end
                Grd1(1).row = k
                Grd1(1).col = 0
                DoEvents
                If Trim(strTemp(i)) = Trim(Grd1(1).Text) Then
                    StrMenu5
                    BolAddData = True
                    Exit For
                End If
            Next k
            If BolAddData = False Then
               Grd1(1).Rows = Grd1(1).Rows + 1
               Grd1(1).row = Grd1(1).Rows - 1
               Grd1(1).Text = strTemp(i)
               StrMenu5
            End If
        End If
    Next i
    DoEvents
    adoRecordset.MoveNext
Loop
CheckOC
SetGridWidth
'計算核准率與總件數和排序
StrMenu6
Grd1(1).Visible = True
Grd1(0).Visible = True
Me.Enabled = True

End Sub

Sub SetGridWidth()

Grd1(0).Width = Grd1(1).Width
'Grd1(0).Clear
'Grd1(1).Clear

If UCase(TxtCheckTOrP) = "P" Then        '專利
'Remove by Morgan 2004/2/13
'    With Grd1(0)
'        .Cols = 7
'        .Row = 0
'        .Col = 0
'        .Text = "審查委員"
'        .ColWidth(0) = 1200
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 1
'        .Text = "發    明"
'        .ColWidth(1) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 2
'        .Text = "新    型"
'        .ColWidth(2) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 3
'        .Text = "新 式 樣"
'        .ColWidth(3) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 4
'        .Text = "訴願、再訴、行訴"
'        .ColWidth(4) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 5
'        .Text = "異議、舉發、答辯"
'        .ColWidth(5) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 6
'        .Text = "總 件 數"
'        .ColWidth(6) = 1595
'        .CellAlignment = flexAlignCenterCenter
'    End With

    With Grd1(1)
        'Add by Morgan 2004/2/13
        .Cols = 19
        .row = 0
        .col = 0
        .Text = ""
        .ColWidth(0) = 1200
        .CellAlignment = flexAlignCenterCenter
        .col = 1
        .Text = "發    明"
        .ColWidth(1) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 2
        .Text = "發    明"
        .ColWidth(2) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 3
        .Text = "發    明"
        .ColWidth(3) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 4
        .Text = "新    型"
        .ColWidth(4) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 5
        .Text = "新    型"
        .ColWidth(5) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 6
        .Text = "新    型"
        .ColWidth(6) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 7
        .Text = "新 式 樣"
        .ColWidth(7) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 8
        .Text = "新 式 樣"
        .ColWidth(8) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 9
        .Text = "新 式 樣"
        .ColWidth(9) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 10
        .Text = "訴願、再訴、行訴"
        .ColWidth(10) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 11
        .Text = "訴願、再訴、行訴"
        .ColWidth(11) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 12
        .Text = "訴願、再訴、行訴"
        .ColWidth(12) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 13
        .Text = "異議、舉發、答辯"
        .ColWidth(13) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 14
        .Text = "異議、舉發、答辯"
        .ColWidth(14) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 15
        .Text = "異議、舉發、答辯"
        .ColWidth(15) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 16
        .Text = "總 件 數"
        .ColWidth(16) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 17
        .Text = "總 件 數"
        .ColWidth(17) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 18
        .Text = "總 件 數"
        .ColWidth(18) = 960
        .CellAlignment = flexAlignCenterCenter
        'Add end
        
        'Modify by Morgan 2004/2/3
        '.Cols = 19
        '.Row = 0
        .row = 1
        .col = 0
        '.Text = ""
        .Text = "審查委員"
        'Modify end
        
        .ColWidth(0) = 1200
        .CellAlignment = flexAlignCenterCenter
        .col = 1
        .Text = "准"
        .ColWidth(1) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 2
        .Text = "駁"
        .ColWidth(2) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 3
        .Text = "核准率"
        .ColWidth(3) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 4
        .Text = "准"
        .ColWidth(4) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 5
        .Text = "駁"
        .ColWidth(5) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 6
        .Text = "核准率"
        .ColWidth(6) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 7
        .Text = "准"
        .ColWidth(7) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 8
        .Text = "駁"
        .ColWidth(8) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 9
        .Text = "核准率"
        .ColWidth(9) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 10
        .Text = "准"
        .ColWidth(10) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 11
        .Text = "駁"
        .ColWidth(11) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 12
        .Text = "核准率"
        .ColWidth(12) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 13
        .Text = "准"
        .ColWidth(13) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 14
        .Text = "駁"
        .ColWidth(14) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 15
        .Text = "核准率"
        .ColWidth(15) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 16
        .Text = "准"
        .ColWidth(16) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 17
        .Text = "駁"
        .ColWidth(17) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 18
        .Text = "核准率"
        .ColWidth(18) = 960
        .CellAlignment = flexAlignCenterCenter
    End With
Else         '商標
   'Remove by Morgan 2004/4/13
'    With grd1(0)
'        .Cols = 6
'        .Row = 0
'        .Col = 0
'        .Text = "審查委員"
'        .ColWidth(0) = 1200
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 1
'        .Text = "正 商 標"
'        .ColWidth(1) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 2
'        .Text = "聯合商標"
'        .ColWidth(2) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 3
'        .Text = "訴願、再訴、行訴"
'        .ColWidth(3) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 4
'        .Text = "異議、評定、答辯"
'        .ColWidth(4) = 1595
'        .CellAlignment = flexAlignCenterCenter
'        .Col = 5
'        .Text = "總 件 數"
'        .ColWidth(5) = 1595
'        .CellAlignment = flexAlignCenterCenter
'    End With
    With Grd1(1)
         'Add by Morgan 2004/4/13
         .Cols = 16
        .row = 0
        .col = 0
        .Text = ""
        .ColWidth(0) = 1200
        .CellAlignment = flexAlignCenterCenter
        
        .col = 1
        .Text = "正 商 標"
        .ColWidth(1) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 2
        .Text = "正 商 標"
        .ColWidth(2) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 3
        .Text = "正 商 標"
        .ColWidth(3) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 4
        .Text = "聯合商標"
        .ColWidth(4) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 5
        .Text = "聯合商標"
        .ColWidth(5) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 6
        .Text = "聯合商標"
        .ColWidth(6) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 7
        .Text = "訴願、再訴、行訴"
        .ColWidth(7) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 8
        .Text = "訴願、再訴、行訴"
        .ColWidth(8) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 9
        .Text = "訴願、再訴、行訴"
        .ColWidth(9) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 10
        .Text = "異議、評定、答辯"
        .ColWidth(10) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 11
        .Text = "異議、評定、答辯"
        .ColWidth(11) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 12
        .Text = "異議、評定、答辯"
        .ColWidth(12) = 960
        .CellAlignment = flexAlignCenterCenter
        
        .col = 13
        .Text = "總 件 數"
        .ColWidth(13) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 14
        .Text = "總 件 數"
        .ColWidth(14) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 15
        .Text = "總 件 數"
        .ColWidth(15) = 960
        .CellAlignment = flexAlignCenterCenter
        
        'Add end
        'Modify by Morgan 2004/4/13
        '.Cols = 16
        '.Row = 0
        .row = 1
        .col = 0
        '.Text = ""
        .Text = "審查委員"
        'Modify end
        .ColWidth(0) = 1200
        .CellAlignment = flexAlignCenterCenter
        .col = 1
        .Text = "准"
        .ColWidth(1) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 2
        .Text = "駁"
        .ColWidth(2) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 3
        .Text = "核准率"
        .ColWidth(3) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 4
        .Text = "准"
        .ColWidth(4) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 5
        .Text = "駁"
        .ColWidth(5) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 6
        .Text = "核准率"
        .ColWidth(6) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 7
        .Text = "准"
        .ColWidth(7) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 8
        .Text = "駁"
        .ColWidth(8) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 9
        .Text = "核准率"
        .ColWidth(9) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 10
        .Text = "准"
        .ColWidth(10) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 11
        .Text = "駁"
        .ColWidth(11) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 12
        .Text = "核准率"
        .ColWidth(12) = 960
        .CellAlignment = flexAlignCenterCenter
        .col = 13
        .Text = "准"
        .ColWidth(13) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 14
        .Text = "駁"
        .ColWidth(14) = 320
        .CellAlignment = flexAlignCenterCenter
        .col = 15
        .Text = "核准率"
        .ColWidth(15) = 960
        .CellAlignment = flexAlignCenterCenter
    End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040204a = Nothing
End Sub

Private Sub Grd1_Click(Index As Integer)
'Grd1(Index).Col = Grd1(Index).MouseCol

End Sub
