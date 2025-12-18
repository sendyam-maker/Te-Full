VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm081031 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   5820
   ClientLeft      =   96
   ClientTop       =   612
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdReceive 
      Caption         =   "取消數量(&C)"
      Height          =   400
      Left            =   1980
      TabIndex        =   21
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton ComSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7548
      TabIndex        =   20
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8376
      TabIndex        =   19
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComAllData 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6420
      TabIndex        =   18
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton ComUCase 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5445
      TabIndex        =   17
      Top             =   70
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   3195
      TabIndex        =   16
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "全部清除(&D)"
      Height          =   400
      Left            =   4320
      TabIndex        =   15
      Top             =   70
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "收文類別"
      Height          =   1104
      Left            =   4920
      TabIndex        =   13
      Top             =   528
      Width           =   3255
      Begin VB.OptionButton Option7 
         Caption         =   "機關來文"
         Height          =   255
         Left            =   336
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Caption         =   "接洽及內部收文單"
         Height          =   375
         Left            =   336
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1248
      Left            =   108
      TabIndex        =   12
      Top             =   444
      Width           =   4575
      Begin VB.OptionButton Option1 
         Caption         =   "電子收文未分案"
         Height          =   240
         Left            =   1980
         TabIndex        =   9
         Top             =   900
         Width           =   1815
      End
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
         Text            =   "ACS"
         Top             =   552
         Width           =   550
      End
      Begin VB.OptionButton Option5 
         Caption         =   "以前未分案"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   900
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
      TabIndex        =   14
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
Attribute VB_Name = "frm081031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/26 Form2.0已修改; MSHFlexGrid1改字型=新細明體-ExtB
'Create by sonia 2019/7/24
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
'Add by Amy 2021/07/09
Dim arrField, intWidth
Dim lngX As Long, lngY As Long
Dim intOrderQty As Integer  '接洽單案件性質數量

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

'Add by Amy 2021/07/09 '取消數量
Private Sub cmdReceive_Click()
    With MSHFlexGrid1
     '有數量才需要清空白
        If .TextMatrix(.row, GetValue("v")) = "v" And Val(.TextMatrix(.row, GetValue("數量"))) > 0 Then
            strSql = "Update Caseprogress Set CP156=null Where CP09='" & .TextMatrix(.row, GetValue("收文號")) & "'"
            cnnConnection.Execute strSql
            .TextMatrix(.row, GetValue("數量")) = ""
            .TextMatrix(.row, GetValue("v")) = ""
        End If
    End With
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
   m_Temprow = 0
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
   frm081031_1.intCP09Col = GetValue("收文號") 'Add by Amy 2021/07/09
   frm081031_1.Caption = Me.Caption
   frm081031_1.Show
  
   '若為主管機關來函時，轉本所案號不可輸入
   If Option7.Value = True Then
      frm081031_1.txtcp01.Enabled = False
      frm081031_1.txtcp02.Enabled = False
      frm081031_1.txtcp03.Enabled = False
      frm081031_1.txtcp04.Enabled = False
   End If
   Me.Hide
End Sub

Private Sub ComUCase_Click()
   m_ReQuery = 2
   m_Temprow = 0
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
   'Modify by Amy 2021/07/09 原:Len(Me.MSHFlexGrid1.TextMatrix(1, 1))
   If Me.MSHFlexGrid1.Rows = 2 And Len(Me.MSHFlexGrid1.TextMatrix(1, GetValue("收文日"))) > 0 And Me.Visible = True Then
      cmdSearch_Click
      ComSure_Click
   End If
End Sub

Private Sub Form_Activate()
   txtGDate1.SetFocus
End Sub

Private Sub Form_Load()
   'Add by Amy 2021/07/09 Grid欄名改arrField設定
   'Modify by Amy 2022/12/07 +CP122
   arrField = Array("v", "數量", "收文日", "收文號", "案件性質", _
                             "本所案號", "當事人", "智權人員", "承辦人", "本所期限", _
                            "是否閉卷", "取消收文日期", "發文日", "法定期限", "CP122")
   intWidth = Array(200, 500, 800, 1000, 1200, _
                               1500, 1000, 900, 800, 800, _
                               0, 0, 0, 900, 0)
   'end 2021/07/09
   
   MoveFormToCenter Me
   'Modify By Sindy 2023/3/8
   'txtGDate1.Text = ChangeWStringToTString(GetTodayDate)
   'txtGDate2 = txtGDate1.Text
   txtGDate1.Text = TransDate(CompWorkDay(-2, strSrvDate(1), 1), 1)
   txtGDate2.Text = strSrvDate(2)
   '2023/3/8 END
   ComUable
  If blnIsFormBack Then
      If CheckChoese(intCmdKind) Then
         PutDataInGrid
         GridHead
      End If
   End If

End Sub

'Mark by Amy 2021/07/09 改動態
Private Sub GridHead_Old()
' Dim i As Integer
'   With MSHFlexGrid1
'      blnOKtoShow = False
'      .Visible = False
'      .row = 0
'      .col = 0: .ColWidth(0) = 200: .Text = "v"
'      .MergeCells = flexMergeRestrictRows
'      .MergeRow(0) = True
'      .CellAlignment = flexAlignCenterCenter
'      .col = 1: .ColWidth(1) = 800: .Text = "收文日"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 3: .ColWidth(3) = 1200: .Text = "案件性質"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 4: .ColWidth(4) = 1500: .Text = "本所案號"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 5: .ColWidth(5) = 1000: .Text = "當事人"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 6: .ColWidth(6) = 800: .Text = "智權人員"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 7: .ColWidth(7) = 800: .Text = "承辦人"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 8: .ColWidth(8) = 800: .Text = "本所期限"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 9
'      .ColWidth(9) = 0: .Text = "是否閉卷"
'      .col = 10
'      .ColWidth(10) = 0: .Text = "取消收文日期"
'      .col = 11
'      .ColWidth(11) = 0: .Text = "發文日"
'      .CellAlignment = flexAlignCenterCenter
'      .col = 12: .ColWidth(12) = 900: .Text = "法定期限"
'      intLastRow = 0
'      blnOKtoShow = True
'      '判斷是否有資料
'      .Visible = True
'   End With
End Sub

'Add by Amy 2021/07/09 改設定1次且改動態
Private Sub GridHead()
    Dim i As Integer
 
    With MSHFlexGrid1
        .Visible = False
        .row = 0
        For i = LBound(arrField) To UBound(arrField)
            .col = i
            .ColWidth(i) = intWidth(i)
            .Text = arrField(i)
            If i = GetValue("v") Then
                .MergeCells = flexMergeRestrictRows
                .MergeRow(0) = True
            End If
            .CellAlignment = flexAlignCenterCenter
        Next i
        intLastRow = 0
        .Visible = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081031 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   Dim nCol As Integer
   Dim i As Integer
   Dim intRow As Integer, strHeight As String 'Add by Amy 2021/07/09
      
   m_row = MSHFlexGrid1.row
   MSHFlexGrid1.col = 1
   m_color = MSHFlexGrid1.CellBackColor
   intCols = MSHFlexGrid1.Cols - 1
   
   'Add by Amy 2021/07/09 若未輸Qty彈輸入視窗(-為不用輸)
   intRow = MSHFlexGrid1.row
   If intRow <> 0 And MSHFlexGrid1.TextMatrix(intRow, GetValue("v")) <> "v" And MSHFlexGrid1.TextMatrix(intRow, GetValue("數量")) = "" Then
        '彈出表單位置控制
        frm040101_3.Label3.Caption = MSHFlexGrid1.TextMatrix(intRow, GetValue("本所案號")) & " (" & MSHFlexGrid1.TextMatrix(intRow, GetValue("收文號")) & ")"
        strHeight = mdiMain.Top + Me.Top + MSHFlexGrid1.Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
        If Val(strHeight) + frm040101_3.Height > Val(Val(mdiMain.Top + mdiMain.Height)) Then
            strHeight = Val(strHeight) - frm040101_3.Height - Val(MSHFlexGrid1.RowHeight(1))
        End If
        frm040101_3.Move mdiMain.Left + Me.Left + MSHFlexGrid1.Left + lngX, Val(strHeight)
        frm040101_3.Show vbModal
        intOrderQty = Val(strPublicTemp)
        strPublicTemp = ""
        If intOrderQty = 0 Then
            Exit Sub
        Else
            strSql = "Update Caseprogress Set CP156=" & Val(intOrderQty) & " Where CP09='" & MSHFlexGrid1.TextMatrix(intRow, GetValue("收文號")) & "'"
            cnnConnection.Execute strSql
            MSHFlexGrid1.TextMatrix(intRow, GetValue("數量")) = intOrderQty
        End If
   End If
   'end 2021/07/09
   
   'Modify by Amy 2021/07/09 +GetValue("收文日") 因加數量欄位變動,顏色設定會有問題
   If Not CheckGridChoese(MSHFlexGrid1, intLastRow, intCols, GetValue("收文日")) Then Exit Sub
   ComSure.Enabled = True
   ComSure.SetFocus
   If m_Temprow <> 0 And m_Temprow <> MSHFlexGrid1.row Then
        i = m_Temprow
        With MSHFlexGrid1
        'Modify by Amy 2021/07/09 原:.TextMatrix(i, 8)
        'Modify by Amy 2022/12/07 +CP122
        If IsEmptyText(.TextMatrix(i, GetValue("本所期限"))) = False Or IsEmptyText(.TextMatrix(i, GetValue("CP122"))) = False Then
            'modify by sonia 2019/11/19 加入未發文條件
            'Modify by Amy 2021/07/09 原:.TextMatrix(i, 8) /.TextMatrix(i, 11)
            If (IsEmptyText(.TextMatrix(i, GetValue("本所期限"))) = False And Val(DBDATE(.TextMatrix(i, GetValue("本所期限")))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(i, GetValue("發文日")))) = 0) _
             Or .TextMatrix(i, GetValue("CP122")) = "Y" Then
               .row = i
               For nCol = 1 To .Cols - 1
                   .row = i
                   .col = nCol
                   .CellBackColor = &H8080FF '紅色
               Next nCol
            End If
        End If
        'Modify by Amy 2021/07/09 原:.TextMatrix(i, 9)
        If IsEmptyText(.TextMatrix(i, GetValue("是否閉卷"))) = False Then
            If .TextMatrix(i, GetValue("是否閉卷")) = "Y" Then
        'end 2021/07/09
               .row = i
               For nCol = 1 To .Cols - 1
                   .row = i
                   .col = nCol
                   .CellBackColor = &HFFFF&
               Next nCol
            End If
        End If
        'Modify by Amy 2021/07/09 原:.TextMatrix(i, 10)
        If IsEmptyText(.TextMatrix(i, GetValue("取消收文日期"))) = False Then
               .row = i
               .col = 2
               .CellBackColor = &HE0E0E0
        End If
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
   txtcp02.SetFocus
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
      If ChkSysName(txtcp01) = True Then
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
         .Visible = True
         CheckColor
         cmdClear.Enabled = True
         CmdSearch.Enabled = True
      Else
         cmdClear.Enabled = False
         CmdSearch.Enabled = False
      End If
   End With
End Sub

Private Function CheckChoese(ByRef i As Integer) As Boolean
Dim Str2 As String, str3 As String
Dim hstr2 As String, ustr3 As String
Dim LcTmp As String
   
   If Option3.Value Then '收文日期
      If IsNull(txtGDate2) Then MsgBox "請輸入日期": Exit Function
      If IsNull(txtGDate1) Then
         Str2 = " and cp05<" + ChangeTStringToWString(txtGDate2)
      Else
         Str2 = " and cp05 between " + ChangeTStringToWString(txtGDate1) + " and " + ChangeTStringToWString(txtGDate2)
      End If
      hstr2 = Str2
      
   ElseIf Option4.Value Then '本所案號
      If txtcp03.Text = "" Then txtcp03 = "0"
      If txtcp04.Text = "" Then txtcp04.Text = "00"
      LcTmp = txtcp01 + txtcp02 + txtcp03 + txtcp04
      Str2 = " and " & ChgCaseprogress(LcTmp) + " and " & ChgLawcase(LcTmp)
      
   ElseIf Option5.Value Then '以前未分案
      Str2 = " and cp14 is null "
      ustr3 = ""
   
   'Add By Sindy 2023/6/8 電子收文未分案
   ElseIf Option1.Value Then
      Str2 = " And F0309='" & Flow_已收文 & "' And F0301 IS NOT NULL "
      ustr3 = " CP14 is null and "
      '2023/6/8 END
   End If
   
   If Option6.Value Then '接洽及內部收文單
      str3 = " and cp09<'C' "
      
   ElseIf Option7.Value Then '機關來文
      str3 = " and cp09>='C' "
   End If
   
   Select Case i
      Case 1 '未分案
         ustr3 = " CP14 is null and "
      Case 2
         'Add By Sindy 2023/6/8 排除電子收文未分案
         If Option1.Value = False Then
         '2023/6/8 END
            ustr3 = ""
         End If
   End Select
   
   'Memo by Amy 2021/07/09 此語法有改要確認 UpdateCurrRecord 也要改
   'modify by sonia 2019/11/19 +CP27
   'Modify By Sindy 2020/10/12 +CP07
   'Modify by Amy 2021/07/09 +數量及欄名
   'Modify by Amy 2022/12/07 +CP122
   'Modify By Sindy 2023/6/8 +,Flow003 +And CP140=F0301(+)
   strExc(0) = "SELECT ' ' as v,Decode(CP140||substr(cp09,1,1),'A',''||CP156,'-') as cntCP156,SUBSTR(CP05, 1, 4)-1911||'/'||SUBSTR(CP05, 5, 2)||'/'||SUBSTR(CP05,7,2) as CP05," + _
      "CP09,CPM03,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CaseNo," & _
      "decode(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) as CUS," & _
      "decode(CP13,S1.ST01,S1.ST02) as CP13,decode(CP14,S2.ST01,S2.ST02) as CP14," & _
      "sqldatet(CP06) as CP06,LC08,CP57,CP27,sqldatet(CP07) as CP07,CP122 from LAWCASE,CASEPROGRESS,STAFF S1," & _
      "STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER,Flow003 WHERE " & ustr3 & _
      " CP01='ACS' " + Str2 + "" + str3 + " AND (substr(lc11,1,8)=cu01(+) and " + _
      "SUBSTR(LC11,9,1)=CU02(+))  AND CP13 = S1.ST01(+)  AND " & _
      "CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and CP14=S2.ST01(+) " & _
      "and cp29 = S3.ST01(+) AND CP01=CPM01 AND CP10=CPM02 And CP140=F0301(+)"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then CheckChoese = True Else CheckChoese = False
   m_Temprow = Empty
   
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
        cmdClear.Enabled = True
      End If
   End If
   If Cancel Then TextInverse txtGDate2
End Sub

Private Sub CheckColor()
   Dim i As Integer
   Dim nCol As Integer
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        'Add by Amy 2021/07/09 數量設置中
        .row = i: .col = 1
        .CellAlignment = flexAlignCenterCenter
        'end 2021/07/09
        'Modify by Amy 2021/07/09 原:.TextMatrix(i, 8)
        'Modify by Amy 2022/12/07 +CP122
        If IsEmptyText(.TextMatrix(i, GetValue("本所期限"))) = False Or IsEmptyText(.TextMatrix(i, GetValue("CP122"))) = False Then
            'modify by sonia 2019/11/19 加入未發文條件
            'Modify by Amy 2021/07/09 原:.TextMatrix(i, 8) /.TextMatrix(i, 11)
            If (IsEmptyText(.TextMatrix(i, GetValue("本所期限"))) = False And Val(DBDATE(.TextMatrix(i, GetValue("本所期限")))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(i, GetValue("發文日")))) = 0) _
              Or .TextMatrix(i, GetValue("CP122")) = "Y" Then
        'end 2022/12/07
               .row = i
               For nCol = 1 To .Cols - 1
                   .row = i
                   .col = nCol
                   .CellBackColor = &H8080FF '紅色
               Next nCol
            End If
        End If
        'Modify by Amy 2021/07/09 原:.TextMatrix(i, 9)
        If IsEmptyText(.TextMatrix(i, GetValue("是否閉卷"))) = False Then
            If .TextMatrix(i, GetValue("是否閉卷")) = "Y" Then
        'end 2021/07/09
               .row = i
               For nCol = 1 To .Cols - 1
                   .row = i
                   .col = nCol
                   .CellBackColor = &HFFFF&
               Next nCol
            End If
        End If
        'Modify by Amy 2021/07/09 原:.TextMatrix(i, 10)
        If IsEmptyText(.TextMatrix(i, GetValue("取消收文日期"))) = False Then
               .row = i
               For nCol = 1 To .Cols - 1
                  .col = nCol
                  .CellBackColor = &HE0E0E0
               Next
        End If
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
                 'Modify by Amy 2021/07/09 原:.TextMatrix(j, x)
                 If .TextMatrix(j, GetValue("收文號")) = m_CP09(i) Then
                    If .TextMatrix(j, GetValue("收文號")) <> strCP09 Then
                       .TextMatrix(j, GetValue("v")) = "v"
                       Exit For
                    End If
                 End If
            Next j
        Next i
      End With
   Else
       For nIndex = 1 To MSHFlexGrid1.Rows - 1
         'Modify by Amy 2021/07/09 原:MSHFlexGrid1.TextMatrix(nIndex, x)
         If MSHFlexGrid1.TextMatrix(nIndex, GetValue("收文號")) = strCP09 Then
            MSHFlexGrid1.TextMatrix(nIndex, GetValue("v")) = Empty
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
   
   'Memo by Amy 2021/07/09 此語法有改要確認 CheckChoese 也要改
   'modify by sonia 2019/11/19 +CP27
   'Modify By Sindy 2020/10/12 +CP07
   'Modify by Amy 2021/07/09 +數量及欄名,與CheckChoese不一致,多decode(CP29,S3.ST01,S3.ST02) as CP29,畫面更新會錯,秀玲:cp29只有法務用,所以拿掉 原:MSHFlexGrid1.TextMatrix(nIndex, 2)
   'Modify by Amy 2022/12/07 +CP122
   If nIndex > 0 And nIndex <= MSHFlexGrid1.Rows - 1 Then
      strSql = "SELECT ' ' as v,Decode(CP140||substr(cp09,1,1),'A',''||CP156,'-') as cntCP156,SUBSTR(CP05, 1, 4)-1911||'/'||SUBSTR(CP05, 5, 2)||'/'||SUBSTR(CP05,7,2) as CP05," + _
      "CP09,CPM03,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CaseNo," & _
      "decode(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) as CUS," & _
      "decode(CP13,S1.ST01,S1.ST02) as CP13,decode(CP14,S2.ST01,S2.ST02) as CP14," & _
      "sqldatet(CP06) as CP06,LC08,CP57,CP27,sqldatet(CP07) as CP07,CP122 from LAWCASE,CASEPROGRESS,STAFF S1," & _
      "STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER WHERE " & _
      " CP01='ACS' AND CP09='" & MSHFlexGrid1.TextMatrix(nIndex, GetValue("收文號")) & "' AND (substr(lc11,1,8)=cu01(+) and " + _
      "SUBSTR(LC11,9,1)=CU02(+))  AND CP13 = S1.ST01(+)  AND " & _
      "CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and CP14=S2.ST01(+) " & _
      "and cp29 = S3.ST01(+) AND CP01=CPM01 and CP10=CPM02"
          
      rsSrcTmp.CursorLocation = adUseClient
      rsSrcTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not rsSrcTmp.EOF Then
         With MSHFlexGrid1
              'Modify by Amy 2021/07/09 +數量 且改欄名 原:.TextMatrix(i,x)/rsSrcTmp.Fields(X)
              If Not IsNull(rsSrcTmp.Fields("cntCP156")) Then
                .TextMatrix(nIndex, GetValue("數量")) = rsSrcTmp.Fields("cntCP156")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP05")) Then
                 .TextMatrix(nIndex, GetValue("收文日")) = rsSrcTmp.Fields("CP05")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP09")) Then
                 .TextMatrix(nIndex, GetValue("收文號")) = rsSrcTmp.Fields("CP09")
              End If
              If Not IsNull(rsSrcTmp.Fields("CPM03")) Then
                 .TextMatrix(nIndex, GetValue("案件性質")) = rsSrcTmp.Fields("CPM03")
              End If
              If Not IsNull(rsSrcTmp.Fields("CaseNo")) Then
                 .TextMatrix(nIndex, GetValue("本所案號")) = rsSrcTmp.Fields("CaseNo")
              End If
              If Not IsNull(rsSrcTmp.Fields("CUS")) Then
                .TextMatrix(nIndex, GetValue("當事人")) = rsSrcTmp.Fields("CUS")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP13")) Then
                 .TextMatrix(nIndex, GetValue("智權人員")) = rsSrcTmp.Fields("CP13")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP14")) Then
                 .TextMatrix(nIndex, GetValue("承辦人")) = rsSrcTmp.Fields("CP14")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP06")) Then
                 .TextMatrix(nIndex, GetValue("本所期限")) = rsSrcTmp.Fields("CP06")
              End If
              If Not IsNull(rsSrcTmp.Fields("LC08")) Then
                 .TextMatrix(nIndex, GetValue("是否閉卷")) = rsSrcTmp.Fields("LC08")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP57")) Then
                 .TextMatrix(nIndex, GetValue("取消收文日期")) = rsSrcTmp.Fields("CP57")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP27")) Then
                 .TextMatrix(nIndex, GetValue("發文日")) = rsSrcTmp.Fields("CP27")
              End If
              If Not IsNull(rsSrcTmp.Fields("CP07")) Then
                 .TextMatrix(nIndex, GetValue("法定期限")) = rsSrcTmp.Fields("CP07")
              End If
              
              'Modify by Amy 2022/12/07 +CP122
              If IsEmptyText(.TextMatrix(nIndex, GetValue("本所期限"))) = False Or IsEmptyText(.TextMatrix(nIndex, GetValue("CP122"))) = False Then
                 'modify by sonia 2019/11/19 加入未發文條件
                 If (IsEmptyText(.TextMatrix(nIndex, GetValue("本所期限"))) = False And Val(DBDATE(.TextMatrix(nIndex, GetValue("本所期限")))) <= Val(DBDATE(Date)) And Val(DBDATE(.TextMatrix(nIndex, GetValue("發文日")))) = 0) _
                  Or .TextMatrix(nIndex, GetValue("CP122")) = "Y" Then
                     .row = nIndex
                     For nCol = 1 To .Cols - 1
                         .row = nIndex
                         .col = nCol
                         .CellBackColor = &H8080FF '紅色
                     Next nCol
                  End If
            End If
             If IsEmptyText(.TextMatrix(nIndex, GetValue("是否閉卷"))) = False Then
                If .TextMatrix(nIndex, GetValue("是否閉卷")) = "Y" Then
                        .row = nIndex
                     For nCol = 1 To .Cols - 1
                         .row = nIndex
                         .col = nCol
                         .CellBackColor = &HFFFF&
                     Next nCol
                End If
              End If
              If IsEmptyText(.TextMatrix(nIndex, GetValue("取消收文日期"))) = False Then
                  .row = nIndex
                  For nCol = 1 To .Cols - 1
                     .col = nCol
                     .CellBackColor = &HE0E0E0
                  Next
               End If
         'end 2021/07/09
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
              'Modify by Amy 2021/07/09 原:.TextMatrix(i, 0)
              If .TextMatrix(i, GetValue("v")) = "v" Then
                  ReDim Preserve m_CP09(n)
                  'Modify by Amy 2021/07/09 原:.TextMatrix(i, 2)
                  m_CP09(n) = .TextMatrix(i, GetValue("收文號"))
                  n = n + 1
              End If
           Next i
   End With
End Sub

Private Function GetValue(pRowN As String) As Integer
    Dim jj As Integer
 
    For jj = LBound(arrField) To UBound(arrField)
       If UCase(arrField(jj)) = UCase(pRowN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

