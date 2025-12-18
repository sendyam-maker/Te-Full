VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm071001 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   5820
   ClientLeft      =   -3210
   ClientTop       =   1155
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&D)"
      Height          =   400
      Left            =   4185
      TabIndex        =   14
      Top             =   70
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3312
      Left            =   72
      TabIndex        =   22
      Top             =   2400
      Width           =   9132
      _ExtentX        =   16113
      _ExtentY        =   5847
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
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
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "收文類別"
      Height          =   852
      Left            =   72
      TabIndex        =   20
      Top             =   1464
      Width           =   3015
      Begin VB.OptionButton Option7 
         Caption         =   "機關來文"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Caption         =   "接洽及內部收文單"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1752
      Left            =   3336
      TabIndex        =   21
      Top             =   576
      Width           =   4575
      Begin VB.TextBox txtcp04 
         Height          =   288
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   11
         Top             =   703
         Width           =   615
      End
      Begin VB.TextBox txtcp03 
         Height          =   288
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   10
         Top             =   703
         Width           =   375
      End
      Begin VB.TextBox txtcp02 
         Height          =   288
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   9
         Top             =   703
         Width           =   1095
      End
      Begin VB.TextBox txtGDate2 
         Height          =   288
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   6
         Top             =   223
         Width           =   1092
      End
      Begin VB.TextBox txtcp01 
         Height          =   288
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   8
         Top             =   703
         Width           =   495
      End
      Begin VB.OptionButton Option5 
         Caption         =   "以前未分案"
         Height          =   312
         Left            =   264
         TabIndex        =   12
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtGDate1 
         Height          =   288
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   5
         Top             =   223
         Width           =   1092
      End
      Begin VB.OptionButton Option4 
         Caption         =   "本所案號："
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "收文日期："
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   2760
         Y1              =   361
         Y2              =   361
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "系統別"
      Height          =   852
      Left            =   72
      TabIndex        =   19
      Top             =   552
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "FCL"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "法務+顧問"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   3045
      TabIndex        =   13
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton ComUCase 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5310
      TabIndex        =   15
      Top             =   70
      Width           =   945
   End
   Begin VB.CommandButton ComAllData 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6288
      TabIndex        =   16
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8232
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7416
      TabIndex        =   17
      Top             =   70
      Width           =   800
   End
End
Attribute VB_Name = "frm071001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim intCmdKind As Integer
Dim m_row As Integer
Dim m_color As String
Dim m_oldcolor As String
Dim m_Temprow As Integer
Dim com1 As Boolean, com2 As Boolean, com3 As Boolean, com4 As Boolean
Dim m_ReQuery As Integer '1.所有資料 2.未分案
Dim m_CP09() As String
'Added by Lydia 2020/07/14 欄位.Col
Dim colCP06 As Integer '本所期限
Dim colCP27 As Integer '發文日
Dim colClose As Integer  '是否閉卷
Dim colCP57 As Integer '取消收文日

' 設定該筆收文資料已做完存檔的工作
Public Sub SetDataComplete(ByVal strCP09 As String)
   Dim nIndex As Integer
   Dim i As Integer
   Dim j As Integer
      
   If Option4 = True Then  '本所案號
      If m_ReQuery = 1 Then
         ComAllData_Click
      Else
         ComUCase_Click
      End If
      
      With MSHFlexGrid1
         For i = 0 To UBound(m_CP09)
             For j = 1 To .Rows - 1
                 If .TextMatrix(j, 2) = m_CP09(i) Then
                    If MSHFlexGrid1.TextMatrix(j, 2) <> strCP09 Then
                       .TextMatrix(j, 0) = "v"
                       Exit For
                    End If
                 End If
            Next j
        Next i
      End With
   Else
       For nIndex = 1 To MSHFlexGrid1.Rows - 1
         If MSHFlexGrid1.TextMatrix(nIndex, 2) = strCP09 Then
            MSHFlexGrid1.TextMatrix(nIndex, 0) = Empty
            UpdateCurrRecord nIndex
            Exit For
         End If
      Next nIndex
   End If
End Sub
Private Sub cmdClear_Click()
 Dim i As Integer
   With MSHFlexGrid1
      .col = 0
      For i = 1 To .Rows - 1
        .row = i
        .Text = " "
      Next
   End With
   ComSure.Enabled = True
End Sub

Private Sub cmdSearch_Click()
 Dim i As Integer
   With MSHFlexGrid1
      .col = 0
      For i = 1 To .Rows - 1
        .row = i
        .Text = "v"
      Next
   End With
   ComSure.Enabled = True
End Sub

Private Sub ComAllData_Click()
 Dim yn As Boolean
   m_ReQuery = 1
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   intCmdKind = 2
   Screen.MousePointer = vbHourglass
   If CheckChoese(2) Then
      PutDataInGrid
      GridHead 'Added by Morgan 2020/6/9
      ComUCase.Enabled = True
   End If
   'GridHead 'Removed by Morgan 2020/6/9
   Screen.MousePointer = vbDefault
   m_row = Empty
   m_color = Empty
   m_oldcolor = Empty
   m_Temprow = Empty
End Sub

Private Sub ComBack_Click()
   blnIsFormBack = False
   Unload Me
   Set frm071001 = Nothing
End Sub

Private Sub ComSure_Click()
 Dim i As Integer, n As Integer
   With MSHFlexGrid1
      'If .Text = "" Then Exit Sub
      n = 0
      For i = 1 To .Rows - 1
         .Visible = False
         .row = i
         .col = 0
         If .Text = "v" Then
            Exit For
         Else
            If i = .Rows - 1 Then
               MsgBox "請點選欲分案資料"
               .Visible = True
               Exit Sub
             End If
         End If
      Next
      .Visible = True
   End With
   GetChoose
   Me.Hide
   Select Case ChoeseForm
       Case 1 '其他
         'Added by Lydia 2023/03/14 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm071002") = False Then
            Set frm071002 = Nothing
         End If
         'end 2023/03/14
         frm071002.Caption = Me.Caption
         frm071002.Show
         'Add by Morgan 2004/4/20
         '若為主管機關來函時，轉本所案號不可輸入
         If Option7.Value = True Then
            frm071002.txtcp01.Enabled = False
            frm071002.txtcp02.Enabled = False
            frm071002.txtcp03.Enabled = False
            frm071002.txtcp04.Enabled = False
         End If
       Case 2 '系統類別為顧問且案件性質為"顧問聘任"
         frm071003.Caption = Me.Caption
         frm071003.Show
   End Select
   
End Sub

Private Sub ComUCase_Click()
   m_ReQuery = 2
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   intCmdKind = 1
   Screen.MousePointer = vbHourglass
   If CheckChoese(1) Then
      PutDataInGrid
      GridHead
      ComUCase.Enabled = True
   End If
   Screen.MousePointer = vbDefault
   'Add By Cheng 2002/04/24
   '若只搜尋到一筆資料時, 則直接進入下一畫面
   'Modify by Morgan 2003/12/23
   'If Me.MSHFlexGrid1.Rows = 2 And Len("" & Me.MSHFlexGrid1.TextMatrix(1, 1)) > 0 Then
   If Me.MSHFlexGrid1.Rows = 2 And Len("" & Me.MSHFlexGrid1.TextMatrix(1, 1)) > 0 And Me.Visible = True Then
   'Modify end 2003/12/23
   
      cmdSearch_Click
      ComSure_Click
   End If
End Sub

Private Sub Form_Activate()
   'ComUCase.SetFocus
'   txtGDate1.SetFocus
'   If blnIsFormBack Then
'      If CheckChoese(intCmdKind) = True Then
'         PutDataInGrid
'         GridHead
'      End If
'   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'txtGDate1.SetFocus
   If blnIsFormBack Then
      If CheckChoese(intCmdKind) = True Then
         PutDataInGrid
         GridHead
      End If
   'Added by Lydia 2020/07/14
   Else
       GridHead  '預設欄位：取得欄位變數
   'end 2020/07/14
   End If
   txtGDate1.Text = GetTaiwanTodayDate
   txtGDate2 = GetTaiwanTodayDate
   ComUable
End Sub

Private Sub GridHead()
Dim intField As Integer 'Memo by Lydia 2020/07/14 改用變數來設定欄位.Col

   With MSHFlexGrid1
      blnOKtoShow = False
      .Visible = False
      .row = 0
      .col = intField
      .ColWidth(intField) = 200: .Text = "v"
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(intField) = 800: .Text = "收文日"
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(intField) = 930: .Text = "收文號"
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(intField) = 1100: .Text = "案件性質"
      intField = intField + 1
      .col = intField
      .ColWidth(intField) = 1200: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      'add by nickc 2005/09/13
      Dim iDep As String
      iDep = PUB_GetST06(strUserNum)
      intField = intField + 1
      .col = intField
      .col = intField: .Text = "分所號"
      '電腦中心，跟分所才秀
      If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
          .ColWidth(intField) = 0
      Else
          .ColWidth(intField) = 620
      End If
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(intField) = 1000: .Text = "當事人"
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(intField) = 700: .Text = "智權人員"
      'Added by Lydia 2020/07/14 法律所案源收文：案源之介紹人
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(intField) = 700: .Text = "介紹人"
      'end 2020/07/14
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      'edit by nickc 2005/09/30
      '.ColWidth(8) = 900: .Text = "承辦人"
      'Modified by Lydia 2015/10/05
      '.ColWidth(8) = 800: .Text = "承辦律師"
      .ColWidth(intField) = 800: .Text = "承辦人"
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      'edit by nickc 2005/09/30
      '.ColWidth(9) = 900: .Text = "法務人員"
      'Modified by Lydia 2015/10/05
      '.ColWidth(9) = 800: .Text = "承辦法務"
      .ColWidth(intField) = 800: .Text = "協辦人員 "
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField
      .CellAlignment = flexAlignCenterCenter
      .ColWidth(intField) = 1000: .Text = "進度備註"
      intField = intField + 1
      .col = intField
      .ColWidth(intField) = 0: .Text = "本所期限"
      colCP06 = intField 'Added by Lydia 2020/07/14
      intField = intField + 1
      .col = intField
      .ColWidth(intField) = 0: .Text = "是否閉卷"
      colClose = intField 'Added by Lydia 2020/07/14
      intField = intField + 1
      .col = intField
      .ColWidth(intField) = 0: .Text = "取消收文日期"
      colCP57 = intField 'Added by Lydia 2020/07/14
      'add by sonia 2019/11/19
      intField = intField + 1
      .col = intField
      .ColWidth(intField) = 0: .Text = "發文日"
      'end 2019/11/19
      colCP27 = intField 'Added by Lydia 2020/07/14
      intLastRow = 0
      blnOKtoShow = True
      '判斷是否有資料
      .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071001 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   Dim nCol As Integer
   Dim i As Integer
   
   m_row = MSHFlexGrid1.row
   MSHFlexGrid1.col = 1
   m_color = MSHFlexGrid1.CellBackColor
   
   intCols = MSHFlexGrid1.Cols - 1
   If Not CheckGridChoese(MSHFlexGrid1, intLastRow, intCols) Then Exit Sub
   ComSure.Enabled = True
   ComSure.SetFocus
   If m_Temprow <> 0 And m_Temprow <> MSHFlexGrid1.row Then
        i = m_Temprow
        With MSHFlexGrid1
            'Modified by Lydia 2020/07/14 改成變數
'            If IsEmptyText(.TextMatrix(i, 11)) = False Then
'                'modify by sonia 2019/11/19 加入未發文條件
'                If Val(DBDATE(.TextMatrix(i, 11))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(i, 14))) = 0 Then
'                   .row = i
'                   For nCol = 1 To .Cols - 1
'                       .row = i
'                       .col = nCol
'                       .CellBackColor = &H8080FF
'                   Next nCol
'                End If
'            End If
'            If IsEmptyText(.TextMatrix(i, 12)) = False Then
'                If .TextMatrix(i, 12) = "Y" Then
'                   .row = i
'                   For nCol = 1 To .Cols - 1
'                       .row = i
'                       .col = nCol
'                       .CellBackColor = &HFFFF&
'                   Next nCol
'                End If
'            End If
'            If IsEmptyText(.TextMatrix(i, 13)) = False Then
'                   .row = i
'                   .col = 2
'                   .CellBackColor = &HE0E0E0
'            End If
            If IsEmptyText(.TextMatrix(i, colCP06)) = False Then
                '本所期限：到期(紅色)
                If Val(DBDATE(.TextMatrix(i, colCP06))) <= Val(strSrvDate(1)) And Val(DBDATE(.TextMatrix(i, colCP27))) = 0 Then
                   .row = i
                   For nCol = 1 To .Cols - 1
                       .row = i
                       .col = nCol
                       .CellBackColor = &H8080FF
                   Next nCol
                End If
            End If
            If IsEmptyText(.TextMatrix(i, colClose)) = False Then
                '閉卷：黃色
                If .TextMatrix(i, colClose) = "Y" Then
                   .row = i
                   For nCol = 1 To .Cols - 1
                       .row = i
                       .col = nCol
                       .CellBackColor = &HFFFF&
                   Next nCol
                End If
            End If
            If IsEmptyText(.TextMatrix(i, colCP57)) = False Then
              '取消收文：灰色
                   .row = i
                   .col = 2
                   .CellBackColor = &HE0E0E0
            End If
            'end 2020/07/14
        End With
   
   End If
   m_Temprow = m_row
   m_oldcolor = m_color
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
 Dim i As Integer
   intCols = MSHFlexGrid1.Cols - 1
   ShowBar MSHFlexGrid1, intLastRow, intCols
   With MSHFlexGrid1
      .col = 1
       If .Text = "" Then Exit Sub
      .col = 0
      If .Text = "v" Then
         .Text = ""
      Else
         .Text = "v"
      End If
      ComSure.Enabled = True
   End With
End Sub

Private Sub Option1_Click()
   'txtcp01 = "": txtcp02 = "": txtcp03 = "": txtcp04 = ""
End Sub

Private Sub Option3_Click()
   CheckOption
   txtGDate1.SetFocus
End Sub

Private Sub Option4_Click()
   CheckOption
   If Option2 Then
      txtcp01 = "FCL"
   End If
   txtcp01.SetFocus
End Sub

Private Sub Option5_Click()
   CheckOption
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
   CloseIme
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If txtcp01 <> "" Then
      txtcp01 = UCase(txtcp01)
      If Option1 Then
         If txtcp01 = "L" Or txtcp01 = "LA" Then
            com1 = True
         Else
            DataErrorMessage 1, "系統類別"
            Cancel = True
         End If
      Else
         If txtcp01 = "FCL" Then
            com1 = True
         Else
            DataErrorMessage 1, "系統類別"
            Cancel = True
         End If
      End If
   End If
   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub ComUable()
   cmdSearch.Enabled = False
   cmdClear.Enabled = False
   ComSure.Enabled = False
End Sub

Private Sub PutDataInGrid()
 Dim i As Integer, strPropertyName As String, strTempName As String
   With MSHFlexGrid1
      If Not (RsTemp.EOF And RsTemp.BOF) Then
         RsTemp.MoveFirst
         .Visible = False
         Set .Recordset = RsTemp
         CheckColor
         .Visible = True
         cmdSearch.Enabled = True
         cmdClear.Enabled = True
      Else
         cmdSearch.Enabled = False
         cmdClear.Enabled = False
      End If
   End With
End Sub
Private Sub CheckColor()
Dim i As Integer
Dim nCol As Integer
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         'Modified by Lydia 2020/07/14 改成變數
'         If IsEmptyText(.TextMatrix(i, 11)) = False Then
'             'modify by sonia 2019/11/19 加入未發文條件
'             If Val(DBDATE(.TextMatrix(i, 11))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(i, 14))) = 0 Then
'                .row = i
'                For nCol = 1 To .Cols - 1
'                    .row = i
'                    .col = nCol
'                    .CellBackColor = &H8080FF
'                Next nCol
'             End If
'         End If
'         If IsEmptyText(.TextMatrix(i, 12)) = False Then
'             If .TextMatrix(i, 12) = "Y" Then
'                .row = i
'                For nCol = 1 To .Cols - 1
'                    .row = i
'                    .col = nCol
'                    .CellBackColor = &HFFFF&
'                Next nCol
'             End If
'         End If
'         If IsEmptyText(.TextMatrix(i, 13)) = False Then
'                .row = i
'                For nCol = 1 To .Cols - 1
'                   .col = nCol
'                   .CellBackColor = &HE0E0E0
'                Next
'         End If
         If IsEmptyText(.TextMatrix(i, colCP06)) = False Then
             '本所期限：到期(紅色)
             If Val(DBDATE(.TextMatrix(i, colCP06))) <= Val(strSrvDate(1)) And Val(DBDATE(.TextMatrix(i, colCP27))) = 0 Then
                .row = i
                For nCol = 1 To .Cols - 1
                    .row = i
                    .col = nCol
                    .CellBackColor = &H8080FF
                Next nCol
             End If
         End If
         If IsEmptyText(.TextMatrix(i, colClose)) = False Then
             '閉卷：黃色
             If .TextMatrix(i, colClose) = "Y" Then
                .row = i
                For nCol = 1 To .Cols - 1
                    .row = i
                    .col = nCol
                    .CellBackColor = &HFFFF&
                Next nCol
             End If
         End If
         If IsEmptyText(.TextMatrix(i, colCP57)) = False Then
             '取消收文：灰色
                .row = i
                For nCol = 1 To .Cols - 1
                   .col = nCol
                   .CellBackColor = &HE0E0E0
                Next
         End If
         'end 2020/07/14
      Next i
   End With
End Sub

Private Function CheckChoese(ByRef i As Integer) As Boolean
 Dim Str2 As String, str3 As String
 Dim hstr2 As String, ustr3 As String
 Dim LcTmp As String
   If Option3.Value Then
      If IsNull(txtGDate2) Then MsgBox "請輸入日期": Exit Function
      If IsNull(txtGDate1) Then
         Str2 = " and cp05<" + ChangeTStringToWString(txtGDate2)
      Else
         Str2 = " and (cp05 between " & ChangeTStringToWString(txtGDate1) & " AND " & ChangeTStringToWString(txtGDate2) & ")"
      End If
      hstr2 = Str2
   ElseIf Option4.Value Then
      If txtcp03 = "" Then txtcp03 = "0"
      If txtcp04 = "" Then txtcp04 = "00"
      LcTmp = txtcp01 + txtcp02 + txtcp03 + txtcp04
'      If Option1.Value Then
'         hstr2 = " and " & strcp & "=" + CNULL(LcTmp) + " and " & StrHc & "=" & CNULL(LcTmp)
'      End If
      Str2 = " and cp01=" & CNULL(txtcp01) + " and cp02=" & CNULL(txtcp02) + " and cp03=" & CNULL(txtcp03) + " and cp04=" & CNULL(txtcp04)
   ElseIf Option5.Value Then
      Str2 = " and cp14 is null "
      ustr3 = ""
   End If
   
   If Option6.Value Then
      'MOdify By Cheng 2002/03/25
'      str3 = " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B')"
      str3 = " and CP09<'C' "
   ElseIf Option7.Value Then
      'Modify By Cheng 2002/03/25
'      str3 = " and substr(cp09,1,1)='C'"
      str3 = " and CP09>='C' "
   End If
   Select Case i
      Case 1
         ustr3 = " cp14 is null  and "
      Case 2
        ustr3 = ""
   End Select
   If Option1.Value Then
   'edit by nickc 2005/09/30 加欄位
'    strExc(1) = "SELECT ' ',SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," + _
'        "CP09,DECODE(LC15,020,cpm04,cpm03),CP01||'-'||CP02||'-'||CP03||'-'||CP04," & _
'        "decode(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),decode(CP13,S1.ST01,S1.ST02)," & _
'        "decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02),CP06,LC08,CP57 " & _
'        "FROM LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP," & _
'        "CUSTOMER WHERE " & ustr3 & " CP01='L' " & _
'        Str2 + str3 + " AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and (SUBSTR(LC11,1,8)=CU01(+) and " & _
'        "SUBSTR(LC11,9,1)=CU02(+)) AND CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) and CP29 = S3.ST01(+) AND " + _
'        "cp01=cpm01(+) and cp10=cpm02(+) union all select ' ',SUBSTR(CP05,1,4)-1911||'/'||" + _
'        "SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2),CP09,cpm03," + _
'        "CP01||'-'||CP02||'-'||CP03||'-'||CP04, decode(HC05, CU01||CU02," + _
'        "NVL(CU04, NVL(CU05,CU06))),decode(CP13,S1.ST01,S1.ST02)," + _
'        "decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02),CP06,HC09,CP57 " + _
'        "from HIRECASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER " + _
'        "where " + ustr3 + " cp01='LA' " + Str2 + str3 + " AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+)" & _
'        " and (substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+)) AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and " + _
'        "cp29 = s3.st01(+) AND cp01=cpm01(+) and cp10=cpm02(+)"
    'modify by sonia 2019/11/19 +CP27
    'Modified by Lydia 2020/07/14 +別名; 增加案源之介紹人
    'strExc(1) = "SELECT ' ',SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," + _
        "CP09,DECODE(LC15,020,cpm04,cpm03),CP01||'-'||CP02||'-'||CP03||'-'||CP04,lc16," & _
        "decode(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),decode(CP13,S1.ST01,S1.ST02)," & _
        "decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02),cp64,CP06,LC08,CP57,CP27 " & _
        "FROM LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP," & _
        "CUSTOMER WHERE " & ustr3 & " CP01='L' " & _
        Str2 + str3 + " AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and (SUBSTR(LC11,1,8)=CU01(+) and " & _
        "SUBSTR(LC11,9,1)=CU02(+)) AND CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) and CP29 = S3.ST01(+) AND " + _
        "cp01=cpm01(+) and cp10=cpm02(+) union all select ' ',SUBSTR(CP05,1,4)-1911||'/'||" + _
        "SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2),CP09,cpm03," + _
        "CP01||'-'||CP02||'-'||CP03||'-'||CP04, hc07,decode(HC05, CU01||CU02," + _
        "NVL(CU04, NVL(CU05,CU06))),decode(CP13,S1.ST01,S1.ST02)," + _
        "decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02),cp64,CP06,HC09,CP57,CP27 " + _
        "from HIRECASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER " + _
        "where " + ustr3 + " cp01='LA' " + Str2 + str3 + " AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+)" & _
        " and (substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+)) AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and " + _
        "cp29 = s3.st01(+) AND cp01=cpm01(+) and cp10=cpm02(+)"
    strExc(1) = "SELECT ' ' V,SQLDATET(CP05) 收文日,CP09 收文號,DECODE(LC15,020,CPM04,CPM03) 案件性質,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號, " & _
                     "LC16 分所號,DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) 當事人,DECODE(CP13,S1.ST01,S1.ST02) 智權人員, " & _
                     "DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,DECODE(CP14,S2.ST01,S2.ST02) 承辦人,DECODE(CP29,S3.ST01,S3.ST02) 協辦人員, " & _
                     "CP64 進度備註,CP06 本所期限,LC08 是否閉卷,CP57 取消收文日期,CP27 發文日 " & _
                     "FROM LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,LawOfficeSource " & _
                     "WHERE " & ustr3 & " CP01='L' " & Str2 & str3 + " AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (SUBSTR(LC11,1,8)=CU01(+) AND " & _
                     "SUBSTR(LC11,9,1)=CU02(+)) AND CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) AND CP29 = S3.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP162=LOS15(+) "
    strExc(1) = strExc(1) & " Union all SELECT ' ' V,SQLDATET(CP05) 收文日,CP09 收文號,CPM03 案件性質,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號, " & _
                     "HC07 分所號,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) 當事人,DECODE(CP13,S1.ST01,S1.ST02) 智權人員, " & _
                     "DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,DECODE(CP14,S2.ST01,S2.ST02) 承辦人,DECODE(CP29,S3.ST01,S3.ST02) 協辦人員, " & _
                     "CP64 進度備註,CP06 本所期限,HC09 是否閉卷,CP57 取消收文日期,CP27 發文日 " & _
                     "FROM HIRECASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,LawOfficeSource " & _
                     "WHERE " & ustr3 & " CP01='LA' " & Str2 & str3 + " AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND " & _
                     "SUBSTR(HC05,9,1)=CU02(+)) AND CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) AND CP29 = S3.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP162=LOS15(+) "
   Else
'edit by nickc 2005/09/30 加欄位
'     strExc(1) = "SELECT ' ',SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," + _
        "cp09,cpm03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,decode(LC11, CU01 || CU02," + _
        "NVL(CU04, NVL(CU05, CU06))),decode(CP13,S1.ST01,S1.ST02)," + _
        "decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02),CP06,LC08,CP57 " + _
        " from LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER" + _
        " WHERE " + ustr3 + " CP01='FCL' " + Str2 + "" + str3 + " AND (substr(lc11,1,8)=cu01(+) and " + _
        " substr(lc11,9,1)=cu02(+)) AND cp13 = s1.st01(+)  AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04" & _
        " and cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) and cp10=cpm02(+)"
    'modify by sonia 2019/11/19 +CP27
     'Modified by Lydia 2020/07/14 +別名; 增加案源之介紹人
     'strExc(1) = "SELECT ' ',SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," + _
        "cp09,cpm03,CP01||'-'||CP02||'-'||CP03||'-'||CP04,lc16,decode(LC11, CU01 || CU02," + _
        "NVL(CU04, NVL(CU05, CU06))),decode(CP13,S1.ST01,S1.ST02)," + _
        "decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02),cp64,CP06,LC08,CP57,CP27 " + _
        " from LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER" + _
        " WHERE " + ustr3 + " CP01='FCL' " + Str2 + "" + str3 + " AND (substr(lc11,1,8)=cu01(+) and " + _
        " substr(lc11,9,1)=cu02(+)) AND cp13 = s1.st01(+)  AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04" & _
        " and cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) and cp10=cpm02(+)"
    strExc(1) = "SELECT ' ' V,SQLDATET(CP05) 收文日,CP09 收文號,DECODE(LC15,020,CPM04,CPM03) 案件性質,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號, " & _
                     "LC16 分所號,DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) 當事人,DECODE(CP13,S1.ST01,S1.ST02) 智權人員, " & _
                     "DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,DECODE(CP14,S2.ST01,S2.ST02) 承辦人,DECODE(CP29,S3.ST01,S3.ST02) 協辦人員, " & _
                     "CP64 進度備註,CP06 本所期限,LC08 是否閉卷,CP57 取消收文日期,CP27 發文日 " & _
                     "FROM LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,LawOfficeSource " & _
                     "WHERE " & ustr3 & " CP01='FCL' " & Str2 & str3 + " AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (SUBSTR(LC11,1,8)=CU01(+) AND " & _
                     "SUBSTR(LC11,9,1)=CU02(+)) AND CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) AND CP29 = S3.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP162=LOS15(+) "
   End If
   intI = 0
   'edit by nickc 2007/02/07 不用 dll 了
   'Set rsTemp = objLawDll.ReadRstMsg(intI, strExc(1))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then CheckChoese = True Else CheckChoese = False
   m_Temprow = Empty 'Add By Sindy 2016/2/4
End Function

Private Sub CheckOption()
   If Option5.Value Then
      ComAllData.Enabled = False
   Else
      ComAllData.Enabled = True
   End If
End Sub

Private Sub ClearGrid()
 Dim i As Integer, j As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
      .row = i
          For j = 0 To .Cols - 1 '11
           .col = j
           .Text = ""
          Next
      Next
     .Rows = 1
   End With
End Sub

Private Function ChoeseForm() As Integer
 Dim CP09 As String, yn As Boolean, CP10 As String
 Dim i As Integer 'Added by Lydia 2018/01/08

   With MSHFlexGrid1
      'Added by Lydia 2018/01/08 逐筆讀取
      For i = 1 To .Rows - 1
         .row = i
      'end 2018/01/08
         .col = 0
         If .Text = "v" Then
            .col = 3
            CP10 = .Text
            .col = 4
            CP09 = .Text
            Exit For 'Added by Lydia 2018/01/08
         End If
      Next 'Added by Lydia 2018/01/08
   End With
   If Left(CP09, 2) = "LA" And CP10 = "顧問聘任" Then
      ChoeseForm = 2
   Else
      ChoeseForm = 1
   End If
End Function

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
Dim strCP03 As String, strCP04 As String
If Option4.Value Then
   If txtcp03 = "" Then strCP03 = "0"
   If txtcp04 = "" Then strCP04 = "00"
   'edit by nickc 2007/02/07 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtcp01, txtcp02, strCP03, strCP04) Then
   If ClsPDCheckCaseCodeIsExist(txtcp01, txtcp02, strCP03, strCP04) Then
      com2 = True
      ComUCase.Enabled = True
      ComAllData.Enabled = True
   End If
End If
End Sub

Private Sub txtGDate1_GotFocus()
   TextInverse txtGDate1
End Sub

Private Sub txtGDate1_Validate(Cancel As Boolean)
   If Not CheckIsTaiwanDate(txtGDate1) Then
      Cancel = True
   End If
   If Cancel Then TextInverse txtGDate1
End Sub

Private Sub txtGDate2_GotFocus()
   TextInverse txtGDate2
End Sub

Private Sub txtGDate2_Validate(Cancel As Boolean)
   If Option3.Value Then
      If Not CheckIsTaiwanDate(txtGDate2) Then
         Cancel = True
      Else
        ComUCase.Enabled = True
        cmdSearch.Enabled = True
        cmdClear.Enabled = True
      End If
   End If
   If Cancel Then TextInverse txtGDate2
End Sub
Private Sub UpdateCurrRecord(ByVal nIndex As String)
   Dim rsSrcTmp As New ADODB.Recordset
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nCol As Integer
   
   If nIndex > 0 And nIndex <= MSHFlexGrid1.Rows - 1 Then
      'modify by sonia 2019/11/19 +CP27
      'Modified by Lydia 2020/07/14 +增加案源之介紹人LOS
      strSql = "SELECT ' '," & SQLDate("CP05") & "," + _
        "CP09,DECODE(LC15,020,cpm04,cpm03),CP01||'-'||CP02||'-'||CP03||'-'||CP04,lc16," & _
        "decode(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),S1.ST02," & _
        "DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,S2.ST02,S3.ST02,cp64,CP06,LC08,CP57,CP27 " & _
        "FROM LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP," & _
        "CUSTOMER,LawOfficeSource WHERE CP09 ='" & MSHFlexGrid1.TextMatrix(nIndex, 2) & "' AND CP01='L' " & _
        "AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and (SUBSTR(LC11,1,8)=CU01(+) and " & _
        "SUBSTR(LC11,9,1)=CU02(+)) AND CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) and CP29 = S3.ST01(+) AND " + _
        "cp01=cpm01(+) and cp10=cpm02(+) AND CP162=LOS15(+) union all select ' '," & SQLDate("CP05") & ",CP09,cpm03," + _
        "CP01||'-'||CP02||'-'||CP03||'-'||CP04, hc07,decode(HC05, CU01||CU02," + _
        "NVL(CU04, NVL(CU05,CU06))),S1.ST02,DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,S2.ST02,S3.ST02,cp64,CP06,HC09,CP57,CP27 " + _
        "from HIRECASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER,LawOfficeSource " + _
        "where CP09='" + MSHFlexGrid1.TextMatrix(nIndex, 2) + "' AND cp01='LA' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+)" & _
        " and (substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+)) AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and " + _
        "cp29 = s3.st01(+) AND cp01=cpm01(+) and cp10=cpm02(+) AND CP162=LOS15(+) "
         rsSrcTmp.CursorLocation = adUseClient
         rsSrcTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If Not rsSrcTmp.EOF Then
            With MSHFlexGrid1
                 'Modified by Lydia 2020/07/15 改成For迴圈
'                 If Not IsNull(rsSrcTmp.Fields(1)) Then
'                    .TextMatrix(nIndex, 1) = rsSrcTmp.Fields(1)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(2)) Then
'                    .TextMatrix(nIndex, 2) = rsSrcTmp.Fields(2)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(3)) Then
'                    .TextMatrix(nIndex, 3) = rsSrcTmp.Fields(3)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(4)) Then
'                    .TextMatrix(nIndex, 4) = rsSrcTmp.Fields(4)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(5)) Then
'                   .TextMatrix(nIndex, 5) = rsSrcTmp.Fields(5)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(6)) Then
'                    .TextMatrix(nIndex, 6) = rsSrcTmp.Fields(6)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(7)) Then
'                    .TextMatrix(nIndex, 7) = rsSrcTmp.Fields(7)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(8)) Then
'                    .TextMatrix(nIndex, 8) = rsSrcTmp.Fields(8)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(9)) Then
'                    .TextMatrix(nIndex, 9) = rsSrcTmp.Fields(9)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(10)) Then
'                    .TextMatrix(nIndex, 10) = rsSrcTmp.Fields(10)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(11)) Then
'                    .TextMatrix(nIndex, 11) = rsSrcTmp.Fields(11)
'                 End If
'                 'add by nickc 2005/09/30
'                 If Not IsNull(rsSrcTmp.Fields(12)) Then
'                    .TextMatrix(nIndex, 12) = rsSrcTmp.Fields(12)
'                 End If
'                 If Not IsNull(rsSrcTmp.Fields(13)) Then
'                    .TextMatrix(nIndex, 13) = rsSrcTmp.Fields(13)
'                 End If
                 For nCol = 1 To 14
                    .TextMatrix(nIndex, nCol) = "" & rsSrcTmp.Fields(nCol)
                 Next nCol
                 'end 2020/07/14
                 
                 'Modified by Lydia 2020/07/14 改成變數
'                 If IsEmptyText(.TextMatrix(nIndex, 11)) = False Then
'                    'modify by sonia 2019/11/19 加入未發文條件
'                    If Val(DBDATE(.TextMatrix(nIndex, 11))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(nIndex, 14))) = 0 Then
'                        .row = nIndex
'                        For nCol = 1 To .Cols - 1
'                            .row = nIndex
'                            .col = nCol
'                            .CellBackColor = &H8080FF
'                        Next nCol
'                     End If
'                 End If
'                 If IsEmptyText(.TextMatrix(nIndex, 12)) = False Then
'                   If .TextMatrix(nIndex, 12) = "Y" Then
'                           .row = nIndex
'                        For nCol = 1 To .Cols - 1
'                            .row = nIndex
'                            .col = nCol
'                            .CellBackColor = &HFFFF&
'                        Next nCol
'                   End If
'                 End If
'                 If IsEmptyText(.TextMatrix(nIndex, 13)) = False Then
'                     .row = nIndex
'                     For nCol = 1 To .Cols - 1
'                        .col = nCol
'                        .CellBackColor = &HE0E0E0
'                     Next
'                 End If
                 If IsEmptyText(.TextMatrix(nIndex, colCP06)) = False Then
                    '本所期限：到期(紅色)
                    If Val(DBDATE(.TextMatrix(nIndex, colCP06))) <= Val(strSrvDate(1)) And Val(DBDATE(.TextMatrix(nIndex, colCP27))) = 0 Then
                        .row = nIndex
                        For nCol = 1 To .Cols - 1
                            .row = nIndex
                            .col = nCol
                            .CellBackColor = &H8080FF
                        Next nCol
                     End If
                 End If
                 If IsEmptyText(.TextMatrix(nIndex, colClose)) = False Then
                   '閉卷：黃色
                   If .TextMatrix(nIndex, colClose) = "Y" Then
                           .row = nIndex
                        For nCol = 1 To .Cols - 1
                            .row = nIndex
                            .col = nCol
                            .CellBackColor = &HFFFF&
                        Next nCol
                   End If
                 End If
                 If IsEmptyText(.TextMatrix(nIndex, colCP57)) = False Then
                     '取消收文：灰色
                     .row = nIndex
                     For nCol = 1 To .Cols - 1
                        .col = nCol
                        .CellBackColor = &HE0E0E0
                     Next
                 End If
                 'end 2020/07/14
            End With
          End If
   End If
   End Sub
Private Sub GetChoose()
  Dim n As Integer
  Dim i As Integer
  
   With MSHFlexGrid1
           n = 0
           For i = 1 To .Rows - 1
              .row = i
              If .TextMatrix(i, 0) = "v" Then
                  ReDim Preserve m_CP09(n)
                  m_CP09(n) = .TextMatrix(i, 2)
                  n = n + 1
              End If
           Next i
   End With
End Sub
