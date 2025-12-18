VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090202_2_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "欲同時電子送件"
   ClientHeight    =   4440
   ClientLeft      =   2790
   ClientTop       =   3720
   ClientWidth     =   7310
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7310
   Begin VB.CommandButton Command5 
      Caption         =   "電子送件取消"
      Height          =   400
      Left            =   1230
      TabIndex        =   5
      Top             =   345
      Width           =   1290
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   270
      Left            =   30
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   300
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6150
      TabIndex        =   2
      Top             =   345
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Left            =   5130
      TabIndex        =   0
      Top             =   345
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   3555
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   7215
      _ExtentX        =   12718
      _ExtentY        =   6279
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "|本所案號|總收文號|案件性質|商品類別|智權人員|收文規費|發文規費"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "請先做智慧局電子送件後才執行本作業!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5025
   End
End
Attribute VB_Name = "frm090202_2_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/7 Form2.0已修改
Option Explicit

Dim i As Integer
Dim iRow As Integer '本次點選列數
Dim iCol As Integer '本次點選欄數


Private Sub cmdok_Click()
Dim tmpArr As Variant, intJ As Integer
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   'Add By Sindy 2019/2/14 輸入發文規費若沒有Tab跳開游標,而是直接用滑鼠按下確定鍵,不會觸發Validate把發文規費寫入Grd欄位裡
   '                       因此在此處重覆執行一次Validate
   If iRow > 0 Then
      'Modify By Sindy 2025/3/31
      'If MSHFlexGrid2.TextMatrix(iRow, 0) = "V" Then
      If MSHFlexGrid2.TextMatrix(iRow, 3) <> "" Then
      '2025/3/31 END
         Call txtInput_Validate(False)
      End If
   End If
   '2019/2/14 END
'   'Add By Sindy 2025/3/31 再全部比對一下發文規費和原發文規費欄位值
'   '                       有異動時第一個欄位才要打勾(因人員都直接輸入金額)
'   If frm090202_2.cmdCP118.Caption = "修改規費" Then
'      For i = 1 To MSHFlexGrid2.Rows - 1
'         If MSHFlexGrid2.TextMatrix(i, 7) <> MSHFlexGrid2.TextMatrix(i, 9) Then
'            MSHFlexGrid2.TextMatrix(i, 0) = "V"
'         Else
'            MSHFlexGrid2.TextMatrix(i, 0) = ""
'         End If
'      Next i
'   End If
'   '2025/3/31 END
   
   For i = 1 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(i, 0) = "V" Then
         frm090202_2.cmdCP118.Tag = "Y"
      End If
   Next i
   If frm090202_2.cmdCP118.Tag = "Y" Then
      For i = 1 To MSHFlexGrid2.Rows - 1
         If MSHFlexGrid2.TextMatrix(i, 0) = "" And MSHFlexGrid2.TextMatrix(i, 2) = frm090202_2.m_EEP01 Then
            MsgBox "有勾電子送件，此案(" & MSHFlexGrid2.TextMatrix(i, 2) & MSHFlexGrid2.TextMatrix(i, 3) & ")必須為電子送件！"
            Exit Sub
         End If
         
         'Add By Sindy 2020/9/30 多案分割時,
         If frm090202_2.m_RetrunRecv <> "" And frm090202_2.m_EEP01 <> frm090202_2.m_RetrunRecv Then
            If MSHFlexGrid2.TextMatrix(i, 8) = "308" Then
               tmpArr = Split(MSHFlexGrid2.TextMatrix(i, 1), "-")
               For intJ = 0 To UBound(tmpArr)
                  If intJ = 0 Then
                     strCP01 = tmpArr(intJ)
                  ElseIf intJ = 1 Then
                     strCP02 = tmpArr(intJ)
                  ElseIf intJ = 2 Then
                     strCP03 = tmpArr(intJ)
                  ElseIf intJ = 3 Then
                     strCP04 = tmpArr(intJ)
                  Else
                     Exit For
                  End If
               Next intJ
               If strCP03 = "" Then strCP03 = "0"
               If strCP04 = "" Then strCP04 = "00"
               '是否為母案
               strExc(0) = "SELECT * FROM divisioncase where dc05='" & strCP01 & "' and dc06='" & strCP02 & "' and dc07='" & strCP03 & "' and dc08='" & strCP04 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If Val(MSHFlexGrid2.TextMatrix(i, 7)) = 0 Then
                     MsgBox MSHFlexGrid2.TextMatrix(i, 1) & " 分割母案必須輸入發文規費！"
                     Exit Sub
                  End If
               Else
                  If Val(MSHFlexGrid2.TextMatrix(i, 7)) > 0 Then
                     MsgBox MSHFlexGrid2.TextMatrix(i, 1) & " 分割子案發文規費應為0！"
                     Exit Sub
                  End If
               End If
            End If
         End If
         '2020/9/30 END
      Next i
      Call frm090202_2.cmdCP118_LostFocus
      Me.Hide
   Else
      Unload Me
   End If
End Sub

Private Sub Command1_Click()
   frm090202_2.cmdCP118.Tag = ""
   Call frm090202_2.cmdCP118_LostFocus
   Unload Me
End Sub

'Add By Sindy 2022/3/17 電子送件取消
Private Sub Command5_Click()
   Dim idx As Integer
   Dim bolSelV As Boolean
   
   With MSHFlexGrid2
   For idx = .Rows - 1 To 1 Step -1
   If .TextMatrix(idx, 0) = "V" Then
      bolSelV = True
      strSql = "update caseprogress set cp85=null,cp84=null" & _
         ",cp118=null where cp09='" & .TextMatrix(idx, 2) & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   Next
   End With
   If bolSelV = False Then
      MsgBox "至少要勾選一筆資料列！"
      Exit Sub
   Else
      Call frm090202_2.cmdCP118_LostFocus
      Me.Hide
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtInput.Visible = False
   'Add By Sindy 2020/12/28
   If frm090202_2.cmdCP118.Caption = "修改規費" Then
      Me.Caption = "修改發文規費"
      Me.Label2.Visible = False
   End If
   '2020/12/28 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090202_2_2 = Nothing
End Sub

Public Function QueryData(Optional ByRef strCaseNo1 As String, Optional ByRef strCaseNo2 As String, Optional ByRef StrCaseNo3 As String, Optional ByRef strCaseNo4 As String, Optional ByRef strTM28 As String) As Boolean
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strCon As String
   
   QueryData = False
   SetDataListWidth
   
   'Modify By Sindy 2020/9/29
   If frm090202_2.m_RetrunRecv <> "" And frm090202_2.m_EEP01 <> frm090202_2.m_RetrunRecv Then
      strExc(0) = "select cp01,cp02,cp03,cp04 from caseprogress" & _
                  " where cp09 in('" & Replace(frm090202_2.m_RetrunRecv, ",", "','") & "')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            strCP01 = RsTemp.Fields("cp01")
            strCP02 = RsTemp.Fields("cp02")
            strCP03 = RsTemp.Fields("cp03")
            strCP04 = RsTemp.Fields("cp04")
            strCon = strCon & " or (cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "')"
            RsTemp.MoveNext
         Loop
      End If
      strCon = Mid(strCon, 4)
      strCon = " and (" & strCon & ")"
   Else
      strExc(0) = "select cp01,cp02,cp03,cp04 from caseprogress" & _
                  " where cp09='" & frm090202_2.m_EEP01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strCP01 = RsTemp.Fields("cp01")
         strCP02 = RsTemp.Fields("cp02")
         strCP03 = RsTemp.Fields("cp03")
         strCP04 = RsTemp.Fields("cp04")
         strCon = " and cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'"
      End If
   End If
    
   'MODIFY BY SONIA 2015/9/9 爭議案也可電子送件,故開放FCT(FCT-032736)
   'Modify By Sindy 2022/3/17 取消 and cp85 is null 限制
   '因為增加"電子送件取消"功能
   'Modify By Sindy 2025/3/31 +原發文規費
   strExc(0) = "select '' V,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號,cp09 總收文號" & _
      ",cpm03 案件性質,TM09 商品類別,st02 智權人員,cp17 收文規費,decode(cp84,null,cp17,cp84) 發文規費,cp10,decode(cp84,null,cp17,cp84) 原發文規費" & _
      " from caseprogress,casepropertymap,trademark,staff" & _
      " where 1=1" & strCon & " and cp158=0 and cp159=0 and cp01 in ('T','FCT')" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm10='000'" & _
      " and st01(+)=cp13 order by 1,2,3"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      QueryData = True
      Set MSHFlexGrid2.Recordset = RsTemp
   End If
End Function

Private Sub MSHFlexGrid2_Click()
   With MSHFlexGrid2
   .row = .MouseRow
   .col = .MouseCol
   If .row = 0 Then
      '.Sort = 7
   Else
      SetBox
   End If
   End With
End Sub

Private Sub MSHFlexGrid2_SelChange()
   Dim ii As Integer
   If MSHFlexGrid2.MouseCol = 7 Then Exit Sub
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid2
   .row = .MouseRow
   .Visible = False
   .col = 0
   If .row <> 0 And .TextMatrix(.row, 1) <> "" Then
      If .Text = "V" Then
           .Text = ""
           For ii = 0 To .Cols - 1
                .col = ii
                .CellBackColor = QBColor(15)
          Next ii
      Else
           .Text = "V"
           For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = &HFFC0C0
           Next ii
      End If
   End If
   .Visible = True
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub SetDataListWidth()
    With MSHFlexGrid2
   .Cols = 11 '10 'Modify By Sindy 2025/3/31 +原發文規費
   .row = 0
   .col = 0: .Text = "V"
   .ColWidth(0) = 250
   .CellAlignment = flexAlignLeftCenter
   .col = 1: .Text = "本所案號"
   .ColWidth(1) = 1200
   .CellAlignment = flexAlignLeftCenter
   .col = 2: .Text = "總收文號"
   .ColWidth(2) = 900
   .CellAlignment = flexAlignLeftCenter
   .col = 3: .Text = "案件性質"
   .ColWidth(3) = 1000
   .CellAlignment = flexAlignLeftCenter
   .col = 4: .Text = "商品類別"
   .ColWidth(4) = 900
   .CellAlignment = flexAlignLeftCenter
   .col = 5: .Text = "智權人員"
   .ColWidth(5) = 800
   .CellAlignment = flexAlignLeftCenter
   .col = 6: .Text = "收文規費"
   .ColWidth(6) = 840
   .CellAlignment = flexAlignRightCenter
   .col = 7: .Text = "發文規費"
   .ColWidth(7) = 840
   .CellAlignment = flexAlignRightCenter
   For intI = 8 To .Cols - 1
      .ColWidth(intI) = 0
   Next
   End With
End Sub

Private Sub GoNext()
   With MSHFlexGrid2
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      SetBox
   End With
End Sub

Private Sub SetBox()
   
   Dim lngLeft As Long, lngTop As Long, ii As Integer
   
   With MSHFlexGrid2
      If .row > 0 And .col = 7 Then
         'If .TextMatrix(.row, 7) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            iRow = .row: iCol = .col
            txtInput.Visible = True
            txtInput.SetFocus
            TextInverse txtInput
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
         'End If
      End If
   End With
End Sub

Private Sub txtInput_GotFocus()
   CloseIme
   TextInverse txtInput
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = Asc(".") Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = vbKeyReturn Then
         'Modified by Morgan 2012/8/13 取消規費>0的限制(分割案只有第一件有規費)
         'If Val(txtInput) > 0 Then
         '   MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput.Text
         '   GoNext
         'Else
         '   MsgBox "發文規費必須大於 0 ！"
         'End If
         MSHFlexGrid2.TextMatrix(iRow, iCol) = Val(txtInput.Text)
         If Val(txtInput.Text) > 0 Then MSHFlexGrid2.TextMatrix(iRow, 0) = "V" 'Add By Sindy 2018/8/14
         GoNext
         'end 2012/8/13
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   End If
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
   'Modified by Morgan 2012/8/13 取消規費>0的限制(分割案只有第一件有規費)
   'If Val(txtInput) > 0 Then
   '   MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput.Text
   '   txtInput.Visible = False
   'Else
   '   MsgBox "發文規費必須大於 0 ！"
   '   Cancel = True
   'End If
   MSHFlexGrid2.TextMatrix(iRow, iCol) = Val(txtInput.Text)
   If Val(txtInput.Text) > 0 Then MSHFlexGrid2.TextMatrix(iRow, 0) = "V" 'Add By Sindy 2018/8/14
   'end 2012/8/13
End Sub
