VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140501 
   BorderStyle     =   1  '單線固定
   Caption         =   "銷案延遲日期輸入作業"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8610
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   0
      MaxLength       =   7
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   2475
      Left            =   75
      TabIndex        =   12
      Top             =   1980
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   4366
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   330
      Index           =   2
      Left            =   7455
      TabIndex        =   7
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&S)"
      Height          =   330
      Index           =   1
      Left            =   6345
      TabIndex        =   6
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   3705
      TabIndex        =   5
      Top             =   390
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   3
      Left            =   3165
      MaxLength       =   2
      TabIndex        =   4
      Top             =   390
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   2
      Left            =   2820
      MaxLength       =   1
      TabIndex        =   3
      Top             =   390
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   1
      Left            =   1830
      MaxLength       =   6
      TabIndex        =   2
      Top             =   390
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   0
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   1
      Top             =   390
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "雙擊該筆資料的延遲期限欄即可輸入！"
      Height          =   180
      Left            =   4920
      TabIndex        =   18
      Top             =   1740
      Width           =   3060
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1170
      TabIndex        =   16
      Top             =   1650
      Width           =   1725
      VariousPropertyBits=   27
      Size            =   "3043;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1170
      TabIndex        =   15
      Top             =   1350
      Width           =   1725
      VariousPropertyBits=   27
      Size            =   "3043;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1170
      TabIndex        =   14
      Top             =   1050
      Width           =   6825
      VariousPropertyBits=   27
      Size            =   "12039;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1170
      TabIndex        =   13
      Top             =   750
      Width           =   6825
      VariousPropertyBits=   27
      Size            =   "12039;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   210
      TabIndex        =   11
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "種類："
      Height          =   180
      Left            =   570
      TabIndex        =   10
      Top             =   1350
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   390
      TabIndex        =   9
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   210
      TabIndex        =   8
      Top             =   750
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1425
      X2              =   3465
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm140501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (Grd1,lbl1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Dim ii As Integer
Dim iRow As Integer '本次點選列數
Dim iCol As Integer '智權人員名稱欄位
'控制輸入方塊用
Dim txtInputMax As Integer
Dim txtInputMin As Integer
Dim txtInputState As String

Private Sub cmdOK_Click(Index As Integer)
Dim strSQL1 As String
Select Case Index
Case 0   'search
         If txt1(0) = "" And txt1(1) = "" Then
            MsgBox "本所案號前2 碼一定要輸！", vbCritical
            Exit Sub
         End If
         '基本資料
         'edit by nickc 2007/04/24  修正，本所案號第 3、4碼沒輸找不到的問題
         strSql = "select pa05,cu04,PTM03,na03 from patent,customer,nation,PATENTTRADEMARKMAP       where substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and pa09=na01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) and pa01='" & txt1(0) & "' and pa02='" & txt1(1) & "' and pa03='" & IIf(Trim(txt1(2)) <> "", txt1(2), "0") & "' and pa04='" & IIf(Trim(txt1(3)) <> "", txt1(3), "00") & "' "
         strSql = strSql & "union select tm05,cu04,PTM03,na03 from trademark,customer,nation,PATENTTRADEMARKMAP  where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and tm10=na01(+) AND '2'=PTM01(+) AND tm08=PTM02(+) and tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) <> "", txt1(2), "0") & "' and tm04='" & IIf(Trim(txt1(3)) <> "", txt1(3), "00") & "' "
         strSql = strSql & "union select sp05,cu04,'',na03 from SERVICEPRACTICE,customer,nation                                        where substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and sp09=na01(+) and sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) <> "", txt1(2), "0") & "' and sp04='" & IIf(Trim(txt1(3)) <> "", txt1(3), "00") & "' "
         strSql = strSql & "union select lc05,cu04,'',na03 from lawcase,customer,nation                                                              where substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) and lc15=na01(+) and lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) <> "", txt1(2), "0") & "' and lc04='" & IIf(Trim(txt1(3)) <> "", txt1(3), "00") & "' "
         strSql = strSql & "union select hc06,cu04,'','台灣' from hirecase,customer                                                                   where substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+) and hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) <> "", txt1(2), "0") & "' and hc04='" & IIf(Trim(txt1(3)) <> "", txt1(3), "00") & "' "
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic
         If AdoRecordSet3.RecordCount <> 0 Then
            lbl1(0).Caption = CheckStr(AdoRecordSet3.Fields(0).Value)
            lbl1(1).Caption = CheckStr(AdoRecordSet3.Fields(1).Value)
            lbl1(2).Caption = CheckStr(AdoRecordSet3.Fields(2).Value)
            lbl1(3).Caption = CheckStr(AdoRecordSet3.Fields(3).Value)
         Else
            lbl1(0).Caption = ""
            lbl1(1).Caption = ""
            lbl1(2).Caption = ""
            lbl1(3).Caption = ""
         End If
         CheckOC3
         'Grdi 資料
         strSQL1 = ""
         strSQL1 = strSQL1 & " and cp01='" & txt1(0) & "' "
         strSQL1 = strSQL1 & " and cp02='" & txt1(1) & "' "
         If txt1(2) <> "" Then
            strSql = strSQL1 & " and cp03='" & txt1(2) & "' "
         End If
         If txt1(3) <> "" Then
            strSQL1 = strSQL1 & " and cp04='" & txt1(3) & "' "
         End If
         'edit by nickc 2005/07/08 已發過的可以輸入，但要提醒
         'strSQL = "select s1.st02," & SQLDate("cp05") & "," & SQLDate("cp06") & "," & SQLDate("cp07") & ",cpm03,s2.st02," & SQLDate("cp108") & ",cp09 from caseprogress,casepropertymap,staff S1,staff S2 where cp01=cpm01(+) and cp10=cpm02(+) and cp14=S1.st01(+) and cp13=s2.st01(+) and cp27 is null and cp57 is null  " & strSQL1
         strSql = "select s1.st02," & SQLDate("cp05") & "," & SQLDate("cp06") & "," & SQLDate("cp07") & ",cpm03,s2.st02," & SQLDate("cp108") & ",cp09,cp27 from caseprogress,casepropertymap,staff S1,staff S2 where cp01=cpm01(+) and cp10=cpm02(+) and cp14=S1.st01(+) and cp13=s2.st01(+)  and cp57 is null  " & strSQL1
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic
         If AdoRecordSet3.RecordCount <> 0 Then
            cmdOK(1).Enabled = True
         Else
            ShowNoData
            cmdOK(1).Enabled = False
         End If
         Set GRD1.Recordset = AdoRecordSet3
         CheckOC
         SetDataListWidth
Case 1   'save
         For ii = 1 To GRD1.Rows - 1
            GRD1.row = ii
            strSql = "update caseprogress set cp108=" & IIf(Trim(GRD1.TextMatrix(ii, 6)) = "", "null", ChangeTStringToWString(ChangeTDateStringToTString(GRD1.TextMatrix(ii, 6)))) & " where cp09='" & GRD1.TextMatrix(ii, 7) & "' "
            cnnConnection.Execute strSql
         Next ii
         MsgBox "存檔成功！", vbOKOnly, "銷案延遲日期輸入作業"
Case 2
      Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140501 = Nothing
End Sub

Private Sub SetDataListWidth()
   
   With GRD1
         .Visible = False
         'edit by nickc 2005/07/08
         '.Cols = 8
         .Cols = 9
         .row = 0
         .col = 0:   .Text = "承辦人"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "收文日"
         .ColWidth(1) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "本所期限"
         .ColWidth(2) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "法定期限"
         .ColWidth(3) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "案件性質"
         .ColWidth(4) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "智權人員"
         .ColWidth(5) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 6: .Text = "延遲期限"
         .ColWidth(6) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 7: .Text = ""
         .ColWidth(7) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 8: .Text = ""
         .ColWidth(8) = 0
         .CellAlignment = flexAlignCenterCenter
         .Visible = True
   End With
End Sub

Private Sub GRD1_DblClick()
   txtInput.Visible = False
   Screen.MousePointer = vbHourglass
    GRD1.Visible = False
    If Me.GRD1.row > 0 Then
'        Grd1.Row = Grd1.MouseRow
'        Grd1.Col = Grd1.MouseCol
        GRD1.Visible = True
        If Trim(GRD1.TextMatrix(GRD1.row, 8)) <> "" Then MsgBox "此筆已經發文過！", vbInformation, "警告！"
        GRD1.Visible = False
        SetBox
    End If
    GRD1.Visible = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Grd1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim iNextRow As Integer, iNextCol As Integer
   If KeyCode = 13 Or (Shift = 0 And KeyCode >= 37 And KeyCode <= 40) Then
      With GRD1
         iNextRow = .row
         iNextCol = .col
         Select Case KeyCode
            Case 13
               SetBox
            Case 38 '上
               iNextRow = .row - 1
            Case 40 '下
               iNextRow = .row + 1
            Case 37 '左
               iNextCol = .col - 1
            Case 39 '右
               iNextCol = .col + 1
         End Select
         If iNextRow > 1 And iNextRow < .Rows And iNextCol > 0 And iNextCol < .Cols - 1 Then
'            .Row = iNextRow:
            .col = iNextCol
         End If
      End With
   End If
End Sub

Private Sub SetBox()
   
   Dim lngLeft As Long, lngTop As Long
   
   With GRD1
      If .row > 0 And .col = 6 Then
         If .TextMatrix(.row, 1) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = ChangeTDateStringToTString(.TextMatrix(.row, .col))
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            iRow = .row: iCol = .col
            txtInput.Visible = True
            txtInput.Enabled = True
            txtInput.SetFocus
            txtInput.SelStart = 0
            txtInput.SelLength = Len(txtInput)
            lngLeft = .Left + 25
            lngTop = .Top + 25 + .RowHeight(iRow)
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
            cmdOK(0).Default = False
         End If
      End If
   End With
End Sub

Private Sub grd1_Scroll()
      GRD1.SetFocus
      txtInput.Visible = False
      cmdOK(0).Default = False
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then
   KeyAscii = UpperCase(KeyAscii)
End If
End Sub

Private Sub txtInput_GotFocus()
txtInput.SelStart = 0
txtInput.SelLength = Len(txtInput)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   Dim Cancel  As Boolean
   If KeyAscii = vbKeyReturn Then
      Cancel = False
      txtInput_Validate Cancel
      If Cancel = False Then
         GRD1.TextMatrix(iRow, iCol) = ChangeTStringToTDateString(txtInput.Text)
         GRD1.SetFocus
         GRD1.Refresh
         txtInput.Visible = False
         cmdOK(0).Default = False
      Else
         KeyAscii = 0
      End If
   ElseIf KeyAscii = vbKeyEscape Then
      GRD1.SetFocus
      txtInput.Visible = False
      cmdOK(0).Default = False
   End If
End Sub

Private Sub txtInput_LostFocus()
   txtInput.Visible = False
   cmdOK(0).Default = False
   txtInput.Tag = ""
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
If Trim(txtInput) = "" Then Exit Sub
'Modify By Sindy 2011/3/4 可輸入非工作日
'If ChkWorkDay(ChangeTStringToWString(txtInput)) = False Then
'   MsgBox "請輸入工作日！", vbCritical, "錯誤！"
'   txtInput_GotFocus
'   Cancel = True
'End If
If ChkDate(txtInput) = False Then
   Call txtInput_GotFocus
   Cancel = True
   Exit Sub
End If
End Sub
