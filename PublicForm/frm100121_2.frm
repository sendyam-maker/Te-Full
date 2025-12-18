VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100121_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工姓名查詢員工資料"
   ClientHeight    =   5710
   ClientLeft      =   160
   ClientTop       =   1000
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5710
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "可補休"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   7
      Left            =   2880
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "旅遊補助(&T)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   6
      Left            =   3690
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "相片(&P)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   5
      Left            =   6380
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "全部取消(&D)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   4
      Left            =   1350
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   60
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "全部選取(&A)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   150
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   60
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "人事明細資料(&D)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   4840
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   3
      Left            =   8400
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7200
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   60
      Width           =   1140
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5280
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   9285
      _ExtentX        =   16387
      _ExtentY        =   9313
      _Version        =   393216
      Cols            =   22
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   22
   End
End
Attribute VB_Name = "frm100121_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/22 改成Form2.0(grdDataList改Fonts,frm100121__1.txt1(0)改為txtName)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSql As String, i As Integer, j As Integer, strTemp As Variant, strTemp1 As Variant, s As Integer
Dim StrTag As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim m_bQuery As Boolean 'Add By Sindy 2012/6/14
'Add by Amy 2016/07/15
Dim strFieldN()
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序


'2010/9/21 MODIFY BY SONIA 加性別,離職日
Private Sub SetDataListWidth()
'Add by Amy 2016/07/15
ReDim strFieldN(1 To GrdDataList.Cols)

GrdDataList.row = 0
'Add By Sindy 2012/6/14
GrdDataList.col = 0: GrdDataList.Text = "V"
'Mark by Amy 2017/07/21 +相片鈕
'If m_bQuery = True Then
   GrdDataList.ColWidth(0) = 200
'Else
'   grdDataList.ColWidth(0) = 0
'End If
'end 2017/07/21
GrdDataList.CellAlignment = flexAlignCenterCenter
'2012/6/14 End
GrdDataList.col = 1: GrdDataList.Text = "部門": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(1) = 1000
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 2: GrdDataList.Text = "編號": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(2) = 600
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 3: GrdDataList.Text = "姓　名": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(3) = 800
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 4: GrdDataList.Text = "英文別名": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(4) = 800
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 5: GrdDataList.Text = "性別": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(5) = 450
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 6: GrdDataList.Text = "到職日": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(6) = 810
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 7: GrdDataList.Text = "所別": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(7) = 500
GrdDataList.CellAlignment = flexAlignCenterCenter
'Add By Sindy 2014/4/17
GrdDataList.col = 8: GrdDataList.Text = "分機": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(8) = 500
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 9: GrdDataList.Text = "樓層": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(9) = 700
GrdDataList.CellAlignment = flexAlignCenterCenter
'2014/4/17 END
GrdDataList.col = 10: GrdDataList.Text = "離職日": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(10) = 810
GrdDataList.CellAlignment = flexAlignCenterCenter
'2009/6/18 add by sonia
GrdDataList.col = 11: GrdDataList.Text = "職稱": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(11) = 500
GrdDataList.CellAlignment = flexAlignCenterCenter
'2009/6/18 end
'Add By Sindy 2014/4/17
GrdDataList.col = 12: GrdDataList.Text = "職位": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(12) = 700
GrdDataList.CellAlignment = flexAlignCenterCenter
'2014/4/17 END
' 98/02/06 Add By Sindy
GrdDataList.col = 13: GrdDataList.Text = "專業代號-國內": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(13) = 1000
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 14: GrdDataList.Text = "專業代號-國外": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(14) = 1000
GrdDataList.CellAlignment = flexAlignCenterCenter

GrdDataList.col = 15: GrdDataList.Text = "眷屬姓名": strFieldN(GrdDataList.col) = GrdDataList.Text
'Add By Sindy 2018/1/26 有”櫃檯每日信件輸入”權限的人才可以使用”含眷屬姓名做查詢”查詢功能
If frm100121_1.Check1.Tag = "T" Then
   GrdDataList.ColWidth(15) = 3000
Else
   GrdDataList.ColWidth(15) = 0
End If
'2018/1/26 END
GrdDataList.CellAlignment = flexAlignCenterCenter

' 98/02/06 END
'Add By Sindy 2009/10/30
GrdDataList.col = 16: GrdDataList.Text = "期限及郵件之各級主管": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(16) = 6000
GrdDataList.CellAlignment = flexAlignCenterCenter
'2009/10/30 End
'Add by Amy 2016/0715 +st03 for 排序
GrdDataList.col = 17: GrdDataList.Text = "ST03": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(17) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 18: GrdDataList.Text = "ST06": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(18) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 19: GrdDataList.Text = "ST20": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(19) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 20: GrdDataList.Text = "ST21": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(20) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
'end 2016/07/15
'Add by Amy 2019/01/10 +ST71
GrdDataList.col = 21: GrdDataList.Text = "執業地區": strFieldN(GrdDataList.col) = GrdDataList.Text
GrdDataList.ColWidth(21) = 2000
GrdDataList.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
'Add By Sindy 2012/6/14
Dim intRow As Integer
On Error GoTo ErrHnd
'2012/6/14 End

Select Case cmdState
'Add By Sindy 2012/6/14
Case 0 '人事明細資料
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Me.Hide
      For intRow = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = intRow
         If Trim(GrdDataList.Text) = "V" Then
            frm160001.Hide
            frm160001.Enabled = False
            Call frm160001.QueryRecord(GrdDataList.TextMatrix(intRow, 2))
            Set frm160001.UpForm = Me
            frm160001.TBar1.Visible = False
            frm160001.cmdok(5).Visible = True '結束
            frm160001.cmdok(6).Visible = True '查詢下一筆
            frm160001.Enabled = True
            frm160001.Show
            '資料列恢復原狀
            GrdDataList.TextMatrix(intRow, 0) = ""
            For i = 0 To GrdDataList.Cols - 1
               GrdDataList.col = i
               GrdDataList.CellBackColor = QBColor(15)
            Next i
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      Next intRow
      Me.Show
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Exit Sub
Case 1 '全部選取
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Me.Hide
      For intRow = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = intRow
         GrdDataList.TextMatrix(intRow, 0) = "V"
         For i = 0 To GrdDataList.Cols - 1
            GrdDataList.col = i
            GrdDataList.CellBackColor = &HFFC0C0
         Next i
      Next intRow
      Me.Show
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Exit Sub
Case 4 '全部取消
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Me.Hide
      For intRow = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = intRow
         GrdDataList.TextMatrix(intRow, 0) = ""
         For i = 0 To GrdDataList.Cols - 1
            GrdDataList.col = i
            GrdDataList.CellBackColor = QBColor(15)
         Next i
      Next intRow
      Me.Show
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Exit Sub
'2012/6/14 End
'Add by Amy 2017/07/21 員工相片
Case 5
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Me.Hide
    For intRow = 1 To GrdDataList.Rows - 1
        GrdDataList.col = 0
        GrdDataList.row = intRow
        If Trim(GrdDataList.Text) = "V" Then
            frm160001_2.Hide
            frm160001_2.Enabled = False
            '資料列恢復原狀
            GrdDataList.TextMatrix(intRow, 0) = ""
            For i = 0 To GrdDataList.Cols - 1
               GrdDataList.col = i
               GrdDataList.CellBackColor = QBColor(15)
            Next i
            If ChkStaffST04(GrdDataList.TextMatrix(intRow, 2), False) = True Then
                MsgBox GrdDataList.TextMatrix(intRow, 2) & " " & GrdDataList.TextMatrix(intRow, 3) & "已離職！"
            '抓取員工相片
            ElseIf frm160001_2.ReadPhoto(GrdDataList.TextMatrix(intRow, 2)) = True Then
                Set frm160001_2.UpForm = Me
                frm160001_2.Label2(0) = GrdDataList.TextMatrix(intRow, 2)
                frm160001_2.Label2(1) = GrdDataList.TextMatrix(intRow, 3)
                frm160001_2.Label2(2) = GrdDataList.TextMatrix(intRow, 1)
                frm160001_2.Enabled = True
                frm160001_2.Show
                Exit Sub
            Else
                MsgBox GrdDataList.TextMatrix(intRow, 2) & " " & GrdDataList.TextMatrix(intRow, 3) & "無相片！"
                Me.Enabled = True
            End If
        End If
    Next intRow
    Me.Show
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
'Add By Sindy 2019/9/9
Case 6 '旅遊補助
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Me.Hide
      For intRow = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = intRow
         If Trim(GrdDataList.Text) = "V" And Val(DBDATE(GrdDataList.TextMatrix(intRow, 6))) > 0 Then
            'Add By Sindy 2023/7/25
            If cmdok(cmdState).Tag = "" And strUserNum <> GrdDataList.TextMatrix(intRow, 2) Then
               Me.Show
               Me.Enabled = True
               Screen.MousePointer = vbDefault
               MsgBox "無權限查詢他人資料！", vbExclamation
               Exit Sub
            End If
            '2023/7/25 END
            frm160020.Hide
            frm160020.Enabled = False
            frm160020.txt1(0) = GrdDataList.TextMatrix(intRow, 2)
            frm160020.txt1(1) = GrdDataList.TextMatrix(intRow, 2)
            'Add By Sindy 2023/7/25
            If cmdok(cmdState).Tag = "Q" Then
               frm160020.txt1(0).Enabled = True
               frm160020.txt1(1).Enabled = True
               frm160020.cmdok(6).Visible = True '查詢下一筆
            Else
               frm160020.txt1(0).Enabled = False '員工編號欄位鎖住
               frm160020.txt1(1).Enabled = False
               frm160020.cmdok(6).Visible = False '查詢下一筆
            End If
            '2023/7/25 END
            Call frm160020.GetData
            Set frm160020.UpForm = Me
            'Add By Sindy 2023/7/25
            Call frm160020.GetFeeMoney(GrdDataList.TextMatrix(intRow, 2))
            frm160020.Frame1.Visible = True
            '2023/7/25 END
            frm160020.TBar1.Visible = False
            frm160020.Frame2.Visible = False
            frm160020.SSTab1.TabVisible(0) = False
            frm160020.cmdok(5).Visible = True '結束
            'Add By Sindy 2023/11/27
            frm160020.txt1(0).Enabled = False
            frm160020.txt1(1).Enabled = False
            '2023/11/27 END
            frm160020.Enabled = True
            frm160020.Show
            '資料列恢復原狀
            GrdDataList.TextMatrix(intRow, 0) = ""
            For i = 0 To GrdDataList.Cols - 1
               GrdDataList.col = i
               GrdDataList.CellBackColor = QBColor(15)
            Next i
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      Next intRow
      Me.Show
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Exit Sub
'Add By Sindy 2024/12/9
Case 7 '可補休
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Me.Hide
      For intRow = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = intRow
         If Trim(GrdDataList.Text) = "V" And Val(DBDATE(GrdDataList.TextMatrix(intRow, 6))) > 0 Then
            If cmdok(cmdState).Tag = "" And strUserNum <> GrdDataList.TextMatrix(intRow, 2) Then
               Me.Show
               Me.Enabled = True
               Screen.MousePointer = vbDefault
               MsgBox "無權限查詢他人資料！", vbExclamation
               Exit Sub
            End If
            frm160017.Hide
            frm160017.Enabled = False
            frm160017.txt1(0) = GrdDataList.TextMatrix(intRow, 2)
            frm160017.txt1(1) = GrdDataList.TextMatrix(intRow, 2)
            frm160017.txt1(4) = Left(strSrvDate(1), 4) - 1911 & "0101"
            frm160017.txt1(5) = "" 'Left(strSrvDate(1), 4) - 1911 & "1231"
            If cmdok(cmdState).Tag = "Q" Then
               frm160017.txt1(0).Enabled = True
               frm160017.txt1(1).Enabled = True
               frm160017.cmdok(6).Visible = True '查詢下一筆
            Else
               frm160017.txt1(0).Enabled = False '員工編號欄位鎖住
               frm160017.txt1(1).Enabled = False
               frm160017.cmdok(6).Visible = False '查詢下一筆
            End If
            Call frm160017.GetData
            Set frm160017.UpForm = Me
            frm160017.TBar1.Visible = False
            frm160017.SSTab1.TabVisible(0) = False
            frm160017.cmdok(5).Visible = True '結束
            frm160017.txt1(0).Enabled = False
            frm160017.txt1(1).Enabled = False
            frm160017.txtB1008_14(0) = GetCurrFor14RestDay(GrdDataList.TextMatrix(intRow, 2))
            frm160017.txtB1008_14(0).Visible = True
            frm160017.Enabled = True
            frm160017.Show
            '資料列恢復原狀
            GrdDataList.TextMatrix(intRow, 0) = ""
            For i = 0 To GrdDataList.Cols - 1
               GrdDataList.col = i
               GrdDataList.CellBackColor = QBColor(15)
            Next i
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      Next intRow
      Me.Show
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Exit Sub
Case 2
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
Case 3
      fnCloseAllFrm100
      Exit Sub
Case Else
End Select

ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   
   'Add By Sindy 2012/6/14
   m_bQuery = IsUserHasRightOfFunction("frm100121_1", strFind, False)
   '人事明細資料:
   '權限開放:電腦中心,人事處及決策會主管(董事長,副董事長,桂所長,閻副所長,江總,何副總,王副總,杜副總,蘇主祕,林特助)
   If m_bQuery = True And Pub_StrUserSt03 <> "M31" Then '不可為財務處,因財務處ST05也為00
      cmdok(0).Visible = True
      cmdok(1).Visible = True
      cmdok(4).Visible = True
      'cmdok(6).Visible = True 'Add By Sindy 2019/9/9
      cmdok(6).Tag = "Q" 'Add By Sindy 2023/7/25
      cmdok(7).Tag = "Q" 'Add By Sindy 2025/1/7
   Else
      m_bQuery = False
      cmdok(0).Visible = False
      cmdok(1).Visible = False
      cmdok(4).Visible = False
      'cmdok(6).Visible = False 'Add By Sindy 2019/9/9
      cmdok(6).Tag = "" 'Add By Sindy 2023/7/25
      cmdok(7).Tag = "" 'Add By Sindy 2025/1/7
   End If
   '2012/6/14 End
   'Add by Sindy 2019/9/9
   If Pub_StrUserSt03 = "M31" Then
      'cmdOK(6).Visible = True 'Add By Sindy 2019/9/9
      cmdok(6).Tag = "Q" 'Add By Sindy 2023/7/25
      cmdok(7).Tag = "Q" 'Add By Sindy 2025/1/7
   End If
   '2019/9/9 END
   
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
   m_blnColOrderAsc = True 'Add by Amy 2016/07/15
End Sub

Sub StrMenu()
Dim m_i As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSR04 As String

Me.Enabled = False

'2010/9/21 MODIFY BY SONIA
'strSQL1 = ""
strSQL1 = " AND S1.ST01>'1' AND S1.ST01<'PATENT' "
'2010/9/21 END

'modify by sonia 2022/1/22 改Form2.0(txt1(0)改為txtName)
'If Len(frm100121_1.txt1(0).Text) > 0 Then
If Len(frm100121_1.txtName.Text) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100121_1.Label1 & frm100121_1.txtName 'Add By Sindy 2010/11/16
   If frm100121_1.Check1.Value = 0 Then
      strSQL1 = strSQL1 + " And S1.ST02 LIKE '%" & frm100121_1.txtName.Text & "%' "
   ' 98/02/06 Modify By Sindy 含眷屬姓名做查詢
   Else
      strSQL1 = strSQL1 + " And (S1.ST02 LIKE '%" & frm100121_1.txtName.Text & "%' Or SR04 LIKE '%" & frm100121_1.txtName.Text & "%') "
      strSQL1 = strSQL1 + " And S1.ST01=SR01(+) "
      pub_QL05 = pub_QL05 & ";" & frm100121_1.Check1.Caption 'Add By Sindy 2010/11/16
   End If
   ' 98/02/06 End
End If
If Len(Trim(frm100121_1.cboDepName.Text)) > 0 Then
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      strSQL1 = strSQL1 + " And nvl(A0921,A0901) = '" & Trim(Left(frm100121_1.cboDepName.Text, 5)) & "' "
   Else
   '2023/12/22 END
      strSQL1 = strSQL1 + " And A0901 = '" & Trim(Left(frm100121_1.cboDepName.Text, 5)) & "' "
   End If
   pub_QL05 = pub_QL05 & ";" & frm100121_1.Label2 & frm100121_1.cboDepName.Text 'Add By Sindy 2010/11/16
End If
If Len(frm100121_1.txt1(1).Text) > 0 Then
   strSQL1 = strSQL1 + " And S1.ST13 >= '" & Val(Trim(frm100121_1.txt1(1).Text)) + 19110000 & "' "
End If
If Len(frm100121_1.txt1(2).Text) > 0 Then
   strSQL1 = strSQL1 + " And S1.ST13 <= '" & Val(Trim(frm100121_1.txt1(2).Text)) + 19110000 & "' "
End If
If Len(frm100121_1.txt1(1).Text) > 0 Or Len(frm100121_1.txt1(2).Text) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100121_1.Label4 & frm100121_1.txt1(1) & "-" & frm100121_1.txt1(2) 'Add By Sindy 2010/11/16
End If
If frm100121_1.txt1(3).Text <> "Y" Then
   strSQL1 = strSQL1 + " And S1.ST04 <> '2' "
Else
   pub_QL05 = pub_QL05 & ";" & Left(frm100121_1.Label9, 8) & frm100121_1.txt1(3) 'Add By Sindy 2010/11/16
End If
'Add By Cheng 2003/09/03
'員工編號
'Begin
If frm100121_1.txt1(4).Text <> "" Then
   strSQL1 = strSQL1 + " And S1.ST01 = '" & frm100121_1.txt1(4).Text & "' "
   pub_QL05 = pub_QL05 & ";" & frm100121_1.Label5 & frm100121_1.txt1(4) 'Add By Sindy 2010/11/16
End If
'End
'Add By Sindy 2012/6/14
'所別
If frm100121_1.txt1(5).Text <> "" Then
   strSQL1 = strSQL1 + " And S1.ST06 = '" & frm100121_1.txt1(5).Text & "' "
   pub_QL05 = pub_QL05 & ";" & frm100121_1.Label7 & frm100121_1.txt1(5) 'Add By Sindy 2010/11/16
End If
'2012/6/14 End

'Modify By Sindy 2012/6/20
'在職所內員工
If frm100121_1.Check2.Value = 1 Then
   strSQL1 = strSQL1 + " And S1.ST04 = '1' AND exists (select * from SalaryData where S1.ST01=SD01 and (sd02 not in('P','F') or sd02 is null)) "
   pub_QL05 = pub_QL05 & ";" & frm100121_1.Check2.Caption
End If
'2012/6/20 End

'2010/9/21 MODIFY BY SONIA 加性別,離職日
'Modify By Sindy 2012/6/14 +'' as V,
'Modify by Amy 2016/07/15 +ST03/ST06/ST20/ST21,並改以部門編號排序
'Modify by Amy 2019/01/10 +ST71
If frm100121_1.Check1.Value = 0 Then
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "SELECT '' as V,nvl(A0922,'(舊)'||A0902) AS 部門,S1.ST01 AS 編號,S1.ST02 AS 姓名,ED04 AS 英文別名,DECODE(S1.ST22,'M','男','F','女',DECODE(SUBSTR(S1.ST26,2,1),'1','男','2','女',S1.ST22)) AS 性別,SUBSTR(' '||sqldatet(S1.ST13),-9) AS 到職日,DECODE(S1.ST06,'1','北','2','中','3','南','4','高','其他') AS 所別,ED01 AS 分機,ED05 AS 樓層,SUBSTR(' '||sqldatet(S1.ST51),-9) AS 離職日,A1.AC03 AS 職稱,A2.AC03 AS 職位,S1.ST07 AS 國內專業代號,S1.ST17 AS 國外專業代號,'　' as 眷屬姓名, " & _
               " S2.ST02|| " & _
               " DECODE(S3.ST02,null,'',DECODE(S2.ST02,null,'','、')||S3.ST02)|| " & _
               " DECODE(S4.ST02,null,'',DECODE(S2.ST02||S3.ST02,null,'','、')||S4.ST02)|| " & _
               " DECODE(S5.ST02,null,'',DECODE(S2.ST02||S3.ST02||S4.ST02,null,'','、')||S5.ST02) as 期限及郵件之各級主管 " & _
               " ,nvl(S1.ST93,S1.ST03) ST03,S1.ST06,S1.ST20,S1.ST21,S1.ST71 as 執業地區 " & _
               " FROM STAFF S1,ACC090,ACC090NEW,ALLCODE A1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,ALLCODE A2,ExtensionData " & _
               " WHERE S1.ST93=A0921(+) AND S1.ST03=A0901(+) AND '01'=A1.AC01(+) AND S1.ST20=A1.AC02(+) AND '02'=A2.AC01(+) AND S1.ST21=A2.AC02(+) AND S1.ST01=ED02(+) " & _
               " AND S1.ST52=S2.ST01(+) " & _
               " AND S1.ST53=S3.ST01(+) " & _
               " AND S1.ST54=S4.ST01(+) " & _
               " AND S1.ST55=S5.ST01(+) " & strSQL1 & _
               " ORDER BY nvl(S1.ST93,S1.ST03),編號"
   Else
   '2023/12/22 END
      strSql = "SELECT '' as V,NVL(A0902,' ') AS 部門,S1.ST01 AS 編號,S1.ST02 AS 姓名,ED04 AS 英文別名,DECODE(S1.ST22,'M','男','F','女',DECODE(SUBSTR(S1.ST26,2,1),'1','男','2','女',S1.ST22)) AS 性別,SUBSTR(' '||sqldatet(S1.ST13),-9) AS 到職日,DECODE(S1.ST06,'1','北','2','中','3','南','4','高','其他') AS 所別,ED01 AS 分機,ED05 AS 樓層,SUBSTR(' '||sqldatet(S1.ST51),-9) AS 離職日,A1.AC03 AS 職稱,A2.AC03 AS 職位,S1.ST07 AS 國內專業代號,S1.ST17 AS 國外專業代號,'　' as 眷屬姓名, " & _
               " S2.ST02|| " & _
               " DECODE(S3.ST02,null,'',DECODE(S2.ST02,null,'','、')||S3.ST02)|| " & _
               " DECODE(S4.ST02,null,'',DECODE(S2.ST02||S3.ST02,null,'','、')||S4.ST02)|| " & _
               " DECODE(S5.ST02,null,'',DECODE(S2.ST02||S3.ST02||S4.ST02,null,'','、')||S5.ST02) as 期限及郵件之各級主管 " & _
               " ,S1.ST03,S1.ST06,S1.ST20,S1.ST21,S1.ST71 as 執業地區 " & _
               " FROM STAFF S1,ACC090,ALLCODE A1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,ALLCODE A2,ExtensionData " & _
               " WHERE S1.ST03=A0901(+) AND '01'=A1.AC01(+) AND S1.ST20=A1.AC02(+) AND '02'=A2.AC01(+) AND S1.ST21=A2.AC02(+) AND S1.ST01=ED02(+) " & _
               " AND S1.ST52=S2.ST01(+) " & _
               " AND S1.ST53=S3.ST01(+) " & _
               " AND S1.ST54=S4.ST01(+) " & _
               " AND S1.ST55=S5.ST01(+) " & strSQL1 & _
               " ORDER BY S1.ST03,編號"
   End If
'Modify By Sindy 98/02/06
'含眷屬姓名做查詢
Else
   'Modify by Amy 2016/07/15 +ST03/ST06/ST20/ST21,並改以部門編號排序
   'Modify by Amy 2016/11/22 +Group by ,S1.ST03,S1.ST06,S1.ST20,S1.ST21 否則會錯
   'Modify by Amy 2019/01/10 +ST71
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "SELECT '' as V,nvl(A0922,'(舊)'||A0902) AS 部門,S1.ST01 AS 編號,S1.ST02 AS 姓名,ED04 AS 英文別名,DECODE(S1.ST22,'M','男','F','女',DECODE(SUBSTR(S1.ST26,2,1),'1','男','2','女',S1.ST22)) AS 性別,SUBSTR(' '||sqldatet(S1.ST13),-9) AS 到職日,DECODE(S1.ST06,'1','北','2','中','3','南','4','高','其他') AS 所別,ED01 AS 分機,ED05 AS 樓層,SUBSTR(' '||sqldatet(S1.ST51),-9) AS 離職日,A1.AC03 AS 職稱,A2.AC03 AS 職位,S1.ST07 AS 國內專業代號,S1.ST17 AS 國外專業代號,'　' as 眷屬姓名, " & _
                  " S2.ST02|| " & _
                  " DECODE(S3.ST02,null,'',DECODE(S2.ST02,null,'','、')||S3.ST02)|| " & _
                  " DECODE(S4.ST02,null,'',DECODE(S2.ST02||S3.ST02,null,'','、')||S4.ST02)|| " & _
                  " DECODE(S5.ST02,null,'',DECODE(S2.ST02||S3.ST02||S4.ST02,null,'','、')||S5.ST02) as 期限及郵件之各級主管 " & _
                  " ,nvl(S1.ST93,S1.ST03) ST03,S1.ST06,S1.ST20,S1.ST21,S1.ST71 as 執業地區 " & _
                  " FROM STAFF S1,ACC090,ACC090NEW,Staff_Relation,ALLCODE A1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,ALLCODE A2,ExtensionData " & _
                  " WHERE S1.ST93=A0921(+) AND S1.ST03=A0901(+) AND '01'=A1.AC01(+) AND S1.ST20=A1.AC02(+) AND '02'=A2.AC01(+) AND S1.ST21=A2.AC02(+) AND S1.ST01=ED02(+) " & _
                  " AND S1.ST52=S2.ST01(+) " & _
                  " AND S1.ST53=S3.ST01(+) " & _
                  " AND S1.ST54=S4.ST01(+) " & _
                  " AND S1.ST55=S5.ST01(+) " & strSQL1 & _
                  " Group By nvl(A0922,'(舊)'||A0902),S1.ST01,S1.ST02,ED04,S1.ST22,S1.ST26,S1.ST13,S1.ST06,ED01,ED05,S1.ST51,S1.ST04,A1.AC03,A2.AC03,S1.ST07,S1.ST17,S2.ST02,S3.ST02,S4.ST02,S5.ST02,S1.ST93,S1.ST03,S1.ST06,S1.ST20,S1.ST21,S1.ST17,S1.ST71" & _
                  " ORDER BY nvl(S1.ST93,S1.ST03),編號"
   Else
      strSql = "SELECT '' as V,NVL(A0902,' ') AS 部門,S1.ST01 AS 編號,S1.ST02 AS 姓名,ED04 AS 英文別名,DECODE(S1.ST22,'M','男','F','女',DECODE(SUBSTR(S1.ST26,2,1),'1','男','2','女',S1.ST22)) AS 性別,SUBSTR(' '||sqldatet(S1.ST13),-9) AS 到職日,DECODE(S1.ST06,'1','北','2','中','3','南','4','高','其他') AS 所別,ED01 AS 分機,ED05 AS 樓層,SUBSTR(' '||sqldatet(S1.ST51),-9) AS 離職日,A1.AC03 AS 職稱,A2.AC03 AS 職位,S1.ST07 AS 國內專業代號,S1.ST17 AS 國外專業代號,'　' as 眷屬姓名, " & _
                  " S2.ST02|| " & _
                  " DECODE(S3.ST02,null,'',DECODE(S2.ST02,null,'','、')||S3.ST02)|| " & _
                  " DECODE(S4.ST02,null,'',DECODE(S2.ST02||S3.ST02,null,'','、')||S4.ST02)|| " & _
                  " DECODE(S5.ST02,null,'',DECODE(S2.ST02||S3.ST02||S4.ST02,null,'','、')||S5.ST02) as 期限及郵件之各級主管 " & _
                  " ,S1.ST03,S1.ST06,S1.ST20,S1.ST21,S1.ST71 as 執業地區 " & _
                  " FROM STAFF S1,ACC090,Staff_Relation,ALLCODE A1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,ALLCODE A2,ExtensionData " & _
                  " WHERE S1.ST03=A0901(+) AND '01'=A1.AC01(+) AND S1.ST20=A1.AC02(+) AND '02'=A2.AC01(+) AND S1.ST21=A2.AC02(+) AND S1.ST01=ED02(+) " & _
                  " AND S1.ST52=S2.ST01(+) " & _
                  " AND S1.ST53=S3.ST01(+) " & _
                  " AND S1.ST54=S4.ST01(+) " & _
                  " AND S1.ST55=S5.ST01(+) " & strSQL1 & _
                  " Group By A0902,S1.ST01,S1.ST02,ED04,S1.ST22,S1.ST26,S1.ST13,S1.ST06,ED01,ED05,S1.ST51,S1.ST04,A1.AC03,A2.AC03,S1.ST07,S1.ST17,S2.ST02,S3.ST02,S4.ST02,S5.ST02,S1.ST03,S1.ST06,S1.ST20,S1.ST21,S1.ST17,S1.ST71" & _
                  " ORDER BY S1.ST03,編號"
   End If
End If
'98/02/06 End
CheckOC
Dim StrTest1 As String, StrTest2 As String
adoRecordset.CursorLocation = adUseClient

adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/16
Else
   InsertQueryLog (0) 'Add By Sindy 2010/11/16
   ShowNoData
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   'Modify By Cheng 2003/07/30
'    frm100121_1.Visible = True
'    Unload Me
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

Set GrdDataList.Recordset = adoRecordset
' 98/02/06 Add By Sindy
For m_i = 1 To GrdDataList.Rows - 1
   strSql = "SELECT SR04 FROM Staff_Relation " & _
            "WHERE SR01 = '" & GrdDataList.TextMatrix(m_i, 2) & "' " & _
                  "order by SR03 ASC,SR02 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   strSR04 = ""
   If rsTmp.RecordCount > 0 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            If strSR04 = "" Then
               strSR04 = CheckStr(.Fields("SR04"))
            Else
               strSR04 = strSR04 & "、" & CheckStr(.Fields("SR04"))
            End If
            .MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   GrdDataList.TextMatrix(m_i, 15) = strSR04
Next
' 98/02/06 End
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100121_2 = Nothing
End Sub

'Add by Amy 2016/07/15
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GrdDataList.MouseCol < 0 Or GrdDataList.MouseRow < 0 Then Exit Sub
    
    GrdDataList.col = GrdDataList.MouseCol
    GrdDataList.row = GrdDataList.MouseRow
    If Me.GrdDataList.row < 1 And Me.GrdDataList.Text <> "V" Then
        '數字
        If GrdDataList.col = GetValue("部門") Then GrdDataList.col = GetValue("部門", True)
        If GrdDataList.col = GetValue("所別") Then GrdDataList.col = GetValue("所別", True)
        If GrdDataList.col = GetValue("職稱") Then GrdDataList.col = GetValue("職稱", True)
        If GrdDataList.col = GetValue("職位") Then GrdDataList.col = GetValue("職位", True)
        
        If GrdDataList.col = GetValue("ST06") Or GrdDataList.col = GetValue("ST20") Or GrdDataList.col = GetValue("ST21") Then
            If m_blnColOrderAsc = True Then
                Me.GrdDataList.Sort = 3 '數值昇冪
                m_blnColOrderAsc = False
            Else
                Me.GrdDataList.Sort = 4 '數值降冪
                m_blnColOrderAsc = True
            End If
        '文字
        Else
            If m_blnColOrderAsc = True Then
                Me.GrdDataList.Sort = 5 '字串昇冪
                m_blnColOrderAsc = False
            Else
                Me.GrdDataList.Sort = 6 '字串降冪
                m_blnColOrderAsc = True
            End If
        End If
    End If

End Sub

'Add By Sindy 2012/6/14
Private Sub grdDataList_SelChange()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GrdDataList.row = GrdDataList.MouseRow
GrdDataList.col = 0
If GrdDataList.row <> 0 Then
   If GrdDataList.Text = "V" Then
      GrdDataList.Text = ""
      For i = 0 To GrdDataList.Cols - 1
         GrdDataList.col = i
         GrdDataList.CellBackColor = QBColor(15)
      Next i
   Else
      GrdDataList.Text = "V"
      For i = 0 To GrdDataList.Cols - 1
         GrdDataList.col = i
         GrdDataList.CellBackColor = &HFFC0C0
      Next i
   End If
End If
End Sub

'Add by Amy 2016/17/15
Private Function GetValue(pFieldN As String, Optional ByVal bolChange As Boolean = False) As Integer
   Dim jj As Integer, ii As Integer
   Dim strFind As String
 
    For jj = 1 To UBound(strFieldN)
        If UCase(strFieldN(jj)) = UCase(pFieldN) Then
            If bolChange = True Then
                Select Case UCase(pFieldN)
                    Case "部門"
                        strFind = "ST03"
                    Case "所別"
                        strFind = "ST06"
                    Case "職稱"
                        strFind = "ST20"
                    Case "職位"
                        strFind = "ST21"
                End Select
                For ii = 1 To UBound(strFieldN)
                    If UCase(strFieldN(ii)) = UCase(strFind) Then
                        GetValue = ii
                        Exit For
                    End If
                Next ii
            Else
                GetValue = jj
            End If
            Exit For
        End If
    Next jj
End Function
