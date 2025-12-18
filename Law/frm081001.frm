VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm081001 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   5820
   ClientLeft      =   90
   ClientTop       =   615
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9315
   Begin VB.CommandButton ComSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7548
      TabIndex        =   19
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8376
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComAllData 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6420
      TabIndex        =   17
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton ComUCase 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5445
      TabIndex        =   16
      Top             =   70
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   3195
      TabIndex        =   15
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "全部清除(&D)"
      Height          =   400
      Left            =   4320
      TabIndex        =   14
      Top             =   70
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "收文類別"
      Height          =   1104
      Left            =   4920
      TabIndex        =   12
      Top             =   528
      Width           =   3255
      Begin VB.OptionButton Option7 
         Caption         =   "機關來文"
         Height          =   255
         Left            =   336
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Caption         =   "接洽及內部收文單"
         Height          =   375
         Left            =   336
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1128
      Left            =   108
      TabIndex        =   11
      Top             =   504
      Width           =   4575
      Begin VB.TextBox txtcp04 
         Height          =   270
         Left            =   3672
         MaxLength       =   2
         TabIndex        =   7
         Top             =   552
         Width           =   615
      End
      Begin VB.TextBox txtcp03 
         Height          =   270
         Left            =   3216
         MaxLength       =   1
         TabIndex        =   6
         Top             =   552
         Width           =   375
      End
      Begin VB.TextBox txtcp02 
         Height          =   270
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   5
         Top             =   552
         Width           =   1095
      End
      Begin VB.TextBox txtGDate2 
         Height          =   270
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1092
      End
      Begin VB.TextBox txtcp01 
         Height          =   270
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   4
         Top             =   552
         Width           =   550
      End
      Begin VB.OptionButton Option5 
         Caption         =   "以前未分案"
         Height          =   336
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtGDate1 
         Height          =   270
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
      Begin VB.OptionButton Option4 
         Caption         =   "本所案號："
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   552
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "收文日期："
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   2760
         Y1              =   360
         Y2              =   360
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4080
      Left            =   48
      TabIndex        =   13
      Top             =   1680
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   7197
      _Version        =   393216
      Cols            =   15
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm081001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim intCmdKind As Integer
Dim com1 As Boolean, com2 As Boolean, com3 As Boolean, com4 As Boolean
Dim m_row As Integer
Dim m_color As String
Dim m_oldcolor As String
Dim m_Temprow As Integer
Dim m_ReQuery As Integer '1.所有資料 2.未分案
Dim m_CP09() As String
'Added by Lydia 2020/07/14 欄位.Col
Dim colCP06 As Integer '本所期限
Dim colCP27 As Integer '發文日
Dim colClose As Integer  '是否閉卷
Dim colCP57 As Integer '取消收文日

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
   m_Temprow = 0 'Add By Sindy 2009/06/29
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   intCmdKind = 2
   Screen.MousePointer = vbHourglass
   If CheckChoese(2) Then
      PutDataInGrid
      ComUCase.Enabled = True
   End If
   GridHead
   Screen.MousePointer = vbDefault
End Sub

Private Sub ComBack_Click()
   blnIsFormBack = False
   Unload Me
End Sub

Private Sub ComSure_Click()
 Dim i As Integer, n As Integer
   With MSHFlexGrid1
     ' If .Text = "" Then Exit Sub
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
   'Added by Lydia 2023/03/14 配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
   If PUB_CheckFormExist("frm081002") = False Then
      Set frm081002 = Nothing
   End If
   'end 2023/03/14
   frm081002.Caption = Me.Caption
   frm081002.Show
   'Add by Morgan 2004/4/20
   '若為主管機關來函時，轉本所案號不可輸入
   If Option7.Value = True Then
      frm081002.txtcp01.Enabled = False
      frm081002.txtcp02.Enabled = False
      frm081002.txtcp03.Enabled = False
      frm081002.txtcp04.Enabled = False
   End If
   'Add end
   Me.Hide
End Sub

Private Sub ComUCase_Click()
   m_ReQuery = 2
   m_Temprow = 0 'Add By Sindy 2009/06/29
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   intCmdKind = 1
   Screen.MousePointer = vbHourglass
   If CheckChoese(1) Then
      PutDataInGrid
      ComUCase.Enabled = True
   End If
   GridHead
   Screen.MousePointer = vbDefault
   'Add By Cheng 2002/04/24
   'Modify by Morgan 2003/12/23
   'If Me.MSHFlexGrid1.Rows = 2 And Len(Me.MSHFlexGrid1.TextMatrix(1, 1)) > 0 Then
   If Me.MSHFlexGrid1.Rows = 2 And Len(Me.MSHFlexGrid1.TextMatrix(1, 1)) > 0 And Me.Visible = True Then
   'Modify end 2003/12/23
   
      cmdSearch_Click
      ComSure_Click
   End If
End Sub

Private Sub Form_Activate()
   'ComUCase.SetFocus
   txtGDate1.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtGDate1.Text = ChangeWStringToTString(GetTodayDate)
   txtGDate2 = txtGDate1.Text
   ComUable
   If blnIsFormBack Then
      If CheckChoese(intCmdKind) Then
         PutDataInGrid
         GridHead
      End If
   'Added by Lydia 2020/07/14
   Else
       GridHead  '預設欄位：取得欄位變數
   'end 2020/07/14
   End If

End Sub

Private Sub GridHead()
Dim intField As Integer 'Memo by Lydia 2020/07/14 改用變數來設定欄位.Col

   With MSHFlexGrid1
      blnOKtoShow = False
      .Visible = False
      .row = 0
      .col = intField: .ColWidth(intField) = 200: .Text = "v"
      'Mark by Lydia 2020/07/14
      '.MergeCells = flexMergeRestrictRows
      '.MergeRow(intField) = True
      'end 2020/07/14
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField: .ColWidth(intField) = 900: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField: .ColWidth(intField) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField: .ColWidth(intField) = 1200: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField: .ColWidth(intField) = 1500: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField: .ColWidth(intField) = 1000: .Text = "當事人"
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField: .ColWidth(intField) = 1000: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      'Added by Lydia 2020/07/14 法律所案源收文：案源之介紹人
      intField = intField + 1
      .col = intField: .ColWidth(intField) = 1000: .Text = "介紹人"
      .CellAlignment = flexAlignCenterCenter
      'end 2020/07/14
      intField = intField + 1
      'Modified by Lydia 2015/10/05
      '.col = 7: .ColWidth(7) = 900: .Text = "承辦律師"
      .col = intField: .ColWidth(intField) = 900: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      'Modified by Lydia 2015/10/05
      '.col = 8: .ColWidth(8) = 900: .Text = "法務人員"
      .col = intField: .ColWidth(intField) = 900: .Text = "協辦人員"
      .CellAlignment = flexAlignCenterCenter
      intField = intField + 1
      .col = intField:       .ColWidth(intField) = 0: .Text = "本所期限"
      colCP06 = intField 'Added by Lydia 2020/07/14
      intField = intField + 1
      .col = intField:       .ColWidth(intField) = 0: .Text = "是否閉卷"
      colClose = intField 'Added by Lydia 2020/07/14
      intField = intField + 1
      .col = intField:       .ColWidth(intField) = 0: .Text = "取消收文日期"
      colCP57 = intField 'Added by Lydia 2020/07/14
      intField = intField + 1
      .col = intField:       .ColWidth(intField) = 0: .Text = "發文日"
      colCP27 = intField 'Added by Lydia 2020/07/14
      intLastRow = 0
      blnOKtoShow = True
      '判斷是否有資料
      .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081001 = Nothing
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
'            If IsEmptyText(.TextMatrix(i, 9)) = False Then
'                'modify by sonia 2019/11/19 加入未發文條件
'                If Val(DBDATE(.TextMatrix(i, 9))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(i, 12))) = 0 Then
'                   .row = i
'                   For nCol = 1 To .Cols - 1
'                       .row = i
'                       .col = nCol
'                       .CellBackColor = &H8080FF
'                   Next nCol
'                End If
'            End If
'            If IsEmptyText(.TextMatrix(i, 10)) = False Then
'                If .TextMatrix(i, 10) = "Y" Then
'                   .row = i
'                   For nCol = 1 To .Cols - 1
'                       .row = i
'                       .col = nCol
'                       .CellBackColor = &HFFFF&
'                   Next nCol
'                End If
'            End If
'            If IsEmptyText(.TextMatrix(i, 11)) = False Then
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

Private Sub Option3_Click()
   CheckOption
   txtGDate1.SetFocus
End Sub

Private Sub Option4_Click()
   CheckOption
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
      'modify by sonia 2015/9/15 開放外專外商部分人員可操作外法分案,已在checkuse檢查
      'If ChkSysName(txtcp01) = True Then
      If (txtcp01 = "FCL" Or txtcp01 = "CFL" Or txtcp01 = "LIN") And Left(Pub_StrUserSt03, 1) = "F" Then
      ElseIf ChkSysName(txtcp01) = True Then
      'end 2015/9/15
         com1 = True
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub ComUable()
   CmdSearch.Enabled = False
   CmdClear.Enabled = False
   ComSure.Enabled = False
End Sub

Private Sub PutDataInGrid()
 Dim i As Integer, strPropertyName As String, strTempName As String
   With MSHFlexGrid1
      If Not (RsTemp.EOF And RsTemp.BOF) Then
         RsTemp.MoveFirst
         .Visible = False
         Set .Recordset = RsTemp
         .Visible = True
         CheckColor
         CmdClear.Enabled = True
         CmdSearch.Enabled = True
      Else
         CmdClear.Enabled = False
         CmdSearch.Enabled = False
      End If
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
         Str2 = " and cp05 between " + ChangeTStringToWString(txtGDate1) + " and " + ChangeTStringToWString(txtGDate2)
      End If
      hstr2 = Str2
   ElseIf Option4.Value Then
      If txtcp03.Text = "" Then txtcp03 = "0"
      If txtcp04.Text = "" Then txtcp04.Text = "00"
      LcTmp = txtcp01 + txtcp02 + txtcp03 + txtcp04
      Str2 = " and " & ChgCaseprogress(LcTmp) + " and " & ChgLawcase(LcTmp)
   ElseIf Option5.Value Then
      Str2 = " and cp14 is null "
      ustr3 = ""
   End If
   If Option6.Value Then
      'Modify By Cheng 2002/03/25
'      str3 = " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B')"
      str3 = " and cp09<'C' "
   ElseIf Option7.Value Then
      'Modify By Cheng 2002/03/25
'      str3 = " and substr(cp09,1,1)='C'"
      str3 = " and cp09>='C' "
   End If
   Select Case i
      Case 1
         ustr3 = " CP14 is null  and "
      Case 2
         ustr3 = ""
   End Select
   'Modify By Sindy 2009/07/24 增加LIN系統類別
   'modify by sonia 2019/11/19 +CP27
   'Modified by Lydia 2020/07/14 +別名; 增加案源之介紹人
   'strExc(0) = "SELECT ' ',SUBSTR(CP05, 1, 4)-1911||'/'||SUBSTR(CP05, 5, 2)||'/'||SUBSTR(CP05,7,2)," + _
      "CP09,CPM03,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & _
      "decode(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
      "decode(CP13,S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02)," & _
      "decode(CP29,S3.ST01,S3.ST02),CP06,LC08,CP57,CP27 from LAWCASE,CASEPROGRESS,STAFF S1," & _
      "STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER WHERE " & ustr3 & _
      " CP01 IN ('FCL','CFL','LIN') " + Str2 + "" + str3 + " AND (substr(lc11,1,8)=cu01(+) and " + _
      "SUBSTR(LC11,9,1)=CU02(+))  AND CP13 = S1.ST01(+)  AND " & _
      "CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and CP14=S2.ST01(+) " & _
      "and cp29 = S3.ST01(+) AND CP01=CPM01 and CP10=CPM02"
    strExc(0) = "SELECT ' ' V,SQLDATET(CP05) 收文日,CP09 收文號,CPM03 案件性質,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號, " & _
                     "DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) 當事人,DECODE(CP13,S1.ST01,S1.ST02) 智權人員, " & _
                     "DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,DECODE(CP14,S2.ST01,S2.ST02) 承辦人,DECODE(CP29,S3.ST01,S3.ST02) 協辦人員, " & _
                     "CP06 本所期限,LC08 是否閉卷,CP57 取消收文日期,CP27 發文日 " & _
                     "FROM LAWCASE,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,LawOfficeSource " & _
                     "WHERE " & ustr3 & " CP01 IN ('FCL','CFL','LIN') " & Str2 & str3 + " AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (SUBSTR(LC11,1,8)=CU01(+) AND " & _
                     "SUBSTR(LC11,9,1)=CU02(+)) AND CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) AND CP29 = S3.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP162=LOS15(+) "
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
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
          For j = 0 To 11
           .col = j
           .Text = ""
          Next
      Next
     .Rows = 1
   End With
End Sub

Private Function ChoeseForm() As Integer
 Dim CP09 As String, yn As Boolean, CP10 As String
   With MSHFlexGrid1
      .col = 0
      If .Text = "v" Then
         .col = 3
         CP10 = .Text
         .col = 4
         CP09 = .Text
      End If
   End With
   If Left(CP09, 2) = "LA" And CP10 = "顧問聘任" Then
      ChoeseForm = 2
      Exit Function
   End If
   ChoeseForm = 1
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
      If txtcp03 = "" Then
         strCP03 = "0"
      End If
      If txtcp04 = "" Then
         strCP04 = "00"
      End If
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
        CmdSearch.Enabled = True
        CmdClear.Enabled = True
      End If
   End If
   If Cancel Then TextInverse txtGDate2
End Sub

Private Sub CheckColor()
   Dim i As Integer
   Dim nCol As Integer
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         'Modified by Lydia 2020/07/14 改成變數
'         If IsEmptyText(.TextMatrix(i, 9)) = False Then
'             'modify by sonia 2019/11/19 加入未發文條件
'             If Val(DBDATE(.TextMatrix(i, 9))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(i, 12))) = 0 Then
'                .row = i
'                For nCol = 1 To .Cols - 1
'                    .row = i
'                    .col = nCol
'                    .CellBackColor = &H8080FF
'                Next nCol
'             End If
'         End If
'         If IsEmptyText(.TextMatrix(i, 10)) = False Then
'             If .TextMatrix(i, 10) = "Y" Then
'                .row = i
'                For nCol = 1 To .Cols - 1
'                    .row = i
'                    .col = nCol
'                    .CellBackColor = &HFFFF&
'                Next nCol
'             End If
'         End If
'         If IsEmptyText(.TextMatrix(i, 11)) = False Then
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
' 設定該筆收文資料已做完存檔的工作
Public Sub SetDataComplete(ByVal strCP09 As String)
   Dim nIndex As Integer
   Dim i As Integer
   Dim j As Integer
      
   If Option4.Value = True Then   '本所案號
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
Private Sub UpdateCurrRecord(ByVal nIndex As String)
   Dim rsSrcTmp As New ADODB.Recordset
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nCol As Integer
   
   If nIndex > 0 And nIndex <= MSHFlexGrid1.Rows - 1 Then
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/11/19 +CP27
      'Modified by Lydia 2020/07/14 +增加案源之介紹人LOS
      strSql = "SELECT ' ',SUBSTR(CP05, 1, 4)-1911||'/'||SUBSTR(CP05, 5, 2)||'/'||SUBSTR(CP05,7,2)," + _
      "CP09,CPM03,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & _
      "decode(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
      "decode(CP13,S1.ST01,S1.ST02),DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,decode(CP14,S2.ST01,S2.ST02)," & _
      "decode(CP29,S3.ST01,S3.ST02),CP06,LC08,CP57,CP27 from LAWCASE,CASEPROGRESS,STAFF S1," & _
      "STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER,LawOfficeSource WHERE " & _
      " CP01 IN ('FCL','CFL','LIN') AND CP09='" & MSHFlexGrid1.TextMatrix(nIndex, 2) & "' AND (substr(lc11,1,8)=cu01(+) and " + _
      "SUBSTR(LC11,9,1)=CU02(+))  AND CP13 = S1.ST01(+)  AND " & _
      "CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and CP14=S2.ST01(+) " & _
      "and cp29 = S3.ST01(+) AND CP01=CPM01 and CP10=CPM02 AND CP162=LOS15(+) "
          
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
                 For nCol = 1 To 13
                    .TextMatrix(nIndex, nCol) = "" & rsSrcTmp.Fields(nCol)
                 Next nCol
                 'end 2020/07/14
                'Modified by Lydia 2020/07/14 改成變數
'                If IsEmptyText(.TextMatrix(nIndex, 9)) = False Then
'                    'modify by sonia 2019/11/19 加入未發文條件
'                    If Val(DBDATE(.TextMatrix(nIndex, 9))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(nIndex, 12))) = 0 Then
'                        .row = nIndex
'                        For nCol = 1 To .Cols - 1
'                            .row = nIndex
'                            .col = nCol
'                            .CellBackColor = &H8080FF
'                        Next nCol
'                     End If
'                End If
'                If IsEmptyText(.TextMatrix(nIndex, 10)) = False Then
'                   If .TextMatrix(nIndex, 10) = "Y" Then
'                           .row = nIndex
'                        For nCol = 1 To .Cols - 1
'                            .row = nIndex
'                            .col = nCol
'                            .CellBackColor = &HFFFF&
'                        Next nCol
'                   End If
'                End If
'                If IsEmptyText(.TextMatrix(nIndex, 11)) = False Then
'                     .row = nIndex
'                     For nCol = 1 To .Cols - 1
'                        .col = nCol
'                        .CellBackColor = &HE0E0E0
'                     Next
'                End If
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
