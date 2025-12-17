VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010028 
   BorderStyle     =   1  '單線固定
   Caption         =   "電話分機資料維護"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.TextBox txtED02_2 
      Height          =   315
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1520
      Width           =   2775
   End
   Begin VB.TextBox txtED05 
      Height          =   315
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1830
      Width           =   3825
   End
   Begin VB.TextBox txtED03 
      Height          =   270
      Left            =   5880
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1560
      Width           =   285
   End
   Begin VB.TextBox txtED04 
      Height          =   270
      Left            =   5880
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1260
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2370
      Picture         =   "frm010028.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   960
      Width           =   350
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   2010
      Picture         =   "frm010028.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   660
      Width           =   350
   End
   Begin VB.OptionButton Option1 
      Caption         =   "場地名稱："
      Height          =   225
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   1530
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "員工姓名："
      Height          =   225
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   990
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtED01 
      Height          =   285
      Left            =   1260
      MaxLength       =   5
      TabIndex        =   0
      Top             =   660
      Width           =   705
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":0204
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":0520
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":083C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":0A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":0D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":1050
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":136C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":1688
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":1CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010028.frx":1FDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Bindings        =   "frm010028.frx":22F8
      Height          =   3255
      Left            =   45
      TabIndex        =   11
      Top             =   2460
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5736
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "所別|分機號碼|部門|員工編號|姓名或場地|英文別名|樓層|備註"
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
      _Band(0).Cols   =   18
   End
   Begin MSForms.TextBox txtED06 
      Height          =   315
      Left            =   1260
      TabIndex        =   10
      Top             =   2130
      Width           =   7515
      VariousPropertyBits=   679493659
      MaxLength       =   100
      Size            =   "13256;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtED02_1 
      Height          =   285
      Left            =   1260
      TabIndex        =   4
      Top             =   960
      Width           =   1095
      VariousPropertyBits=   679493659
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "模糊比對"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   2730
      TabIndex        =   27
      Top             =   1020
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "(1.北 2.中 3.南 4.高)"
      Height          =   225
      Left            =   6210
      TabIndex        =   26
      Top             =   1590
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "(1.北 2.中 3.南 4.高)"
      Height          =   225
      Left            =   6210
      TabIndex        =   25
      Top             =   990
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文別名："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   4950
      TabIndex        =   24
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   225
      Index           =   7
      Left            =   5310
      TabIndex        =   23
      Top             =   990
      Width           =   540
   End
   Begin VB.Label LblED03 
      Height          =   225
      Left            =   5880
      TabIndex        =   22
      Top             =   990
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   225
      Index           =   0
      Left            =   3600
      TabIndex        =   21
      Top             =   990
      Width           =   900
   End
   Begin VB.Label LblED02_1 
      Height          =   225
      Left            =   4530
      TabIndex        =   20
      Top             =   990
      Width           =   585
   End
   Begin VB.Label Label23 
      Alignment       =   1  '靠右對齊
      Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
      Height          =   225
      Left            =   2670
      TabIndex        =   19
      Top             =   660
      Width           =   6255
   End
   Begin VB.Label LblDept 
      Height          =   195
      Left            =   1260
      TabIndex        =   18
      Top             =   1290
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   225
      Index           =   3
      Left            =   5310
      TabIndex        =   17
      Top             =   1590
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部　　門："
      Height          =   180
      Index           =   6
      Left            =   300
      TabIndex        =   16
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "樓　　層："
      Height          =   180
      Index           =   5
      Left            =   300
      TabIndex        =   15
      Top             =   1890
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Index           =   4
      Left            =   300
      TabIndex        =   14
      Top             =   2160
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分機號碼："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   13
      Top             =   690
      Width           =   930
   End
End
Attribute VB_Name = "frm010028"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/02 Form2.0已修改 txtED02_1/txtED06/grd1
'Create by Sindy 2014/4/16
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的Key
Dim m_FirstKEY(3) As String
' 最後一筆資料的Key
Dim m_LastKEY(3) As String
' 目前正在顯示的Key
Dim m_CurrKEY(3) As String
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_PKey1 As String 'Add By Sindy 2014/5/6
Dim m_PKey2 As String 'Add By Sindy 2014/5/6
Dim m_ConSql As String, m_ConWhereSql As String 'Add By Sindy 2018/12/24


Private Sub Command1_Click()
'   If txtED01.Text = "" Then
'      MsgBox "分機號碼不可以空白！", vbExclamation
'      txtED01.SetFocus
'      Exit Sub
'   End If
   If QueryRecord(1) = False Then
      MsgBox "無此資料！", vbExclamation, "查詢資料"
      UpdateCtrlData
      ReadAllData
   End If
   SetKeyReadOnly True
   m_EditMode = 0
   SetCtrlReadOnly True
   UpdateToolbarState
End Sub

Private Sub Command2_Click()
   If Option1(0).Value = True Then
'      If txtED02_1.Text = "" Then
'         MsgBox "員工姓名不可以空白！", vbExclamation
'         txtED02_1.SetFocus
'         Exit Sub
'      End If
      If QueryRecord(2) = False Then
         MsgBox "無此資料！", vbExclamation, "查詢資料"
         UpdateCtrlData
         ReadAllData
      End If
      SetKeyReadOnly True
      m_EditMode = 0
      SetCtrlReadOnly True
      UpdateToolbarState
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   'Add By Sindy 2018/12/24
   m_ConSql = "": m_ConWhereSql = ""
   If PUB_GetST06(strUserNum) = "2" Then '中所
      m_ConSql = " and ED03='2' "
      m_ConWhereSql = " where ED03='2' "
   ElseIf PUB_GetST06(strUserNum) = "3" Then '南所
      m_ConSql = " and ED03='3' "
      m_ConWhereSql = " where ED03='3' "
   ElseIf PUB_GetST06(strUserNum) = "4" Then '高所
      m_ConSql = " and ED03='4' "
      m_ConWhereSql = " where ED03='4' "
   End If
   '2018/12/24 END
   
   SetDataListWidth
   m_blnColOrderAsc = True
   
   ClearField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   ReadAllData
   'OnAction vbKeyF4
   OnAction vbKeyF10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010028 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grd1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   
   '因排序欲要所別+點選欄位做排序,所以另外對應相關欄位
   If nCol = 0 Then nCol = 9
   If nCol = 1 Then nCol = 10
   If nCol = 2 Then nCol = 11
   If nCol = 3 Then nCol = 12
   If nCol = 4 Then nCol = 16 '離職欄
   If nCol = 5 Then nCol = 13
   If nCol = 6 Then nCol = 17 '英文別名
   If nCol = 7 Then nCol = 14
   If nCol = 8 Then nCol = 15 '備註
   
   grd1.col = nCol
   grd1.row = nRow
   If Me.grd1.row < 1 Then
      If Me.grd1.Text = "分機號碼" Then
         If m_blnColOrderAsc = True Then
            Me.grd1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grd1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub grd1_SelChange()
grd1.Visible = False
If grd1.MouseRow <> 0 Then
'   '上一筆資料列清除反白
'   If dblPrevRow > 0 Then
'      grd1.col = 2
'      grd1.row = dblPrevRow
'      For i = 0 To 1
'         grd1.col = i
'         grd1.CellBackColor = &H8000000F
'      Next i
'      For i = 2 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = QBColor(15)
'      Next i
'   End If
'   '目前資料列反白
'   grd1.col = 0
'   grd1.row = grd1.MouseRow
'   dblPrevRow = grd1.row
'   For i = 0 To grd1.Cols - 1
'      grd1.col = i
'      grd1.CellBackColor = &HFFC0C0
'   Next i
   '查詢目前資料列
   m_CurrKEY(0) = grd1.TextMatrix(grd1.row, 1)
   m_CurrKEY(1) = IIf(grd1.TextMatrix(grd1.row, 3) <> "", grd1.TextMatrix(grd1.row, 3), grd1.TextMatrix(grd1.row, 5))
   m_CurrKEY(2) = grd1.TextMatrix(grd1.row, 9)
   UpdateCtrlData
End If
grd1.Visible = True
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If txtED01.Text = "" Then
      MsgBox "分機號碼不可以空白！", vbExclamation
      txtED01.SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2014/5/5
   If txtED02_1.Text <> "" And txtED02_2.Text <> "" Then
      MsgBox "員工姓名和場地名稱請二選一輸入！", vbExclamation
      Exit Function
   End If
   '2014/5/5 END
   
   If Option1(0).Value = True Then
      If txtED02_1.Text = "" Then
         MsgBox "員工姓名不可以空白！", vbExclamation
         txtED02_1.SetFocus
         Exit Function
      End If
   Else
      If txtED02_2.Text = "" Then
         MsgBox "場地名稱不可以空白！", vbExclamation
         txtED02_2.SetFocus
         Exit Function
      End If
      If txtED03.Text = "" Then
         MsgBox "所別不可以空白！", vbExclamation
         txtED03.SetFocus
         Exit Function
      End If
   End If
   
   Cancel = False
   txtED02_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   Cancel = False
   txtED02_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If

   Cancel = False
   txtED03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If

   Cancel = False
   txtED04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   Cancel = False
   txtED05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   Cancel = False
   txtED06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   'Add by Amy 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
        Exit Function
   End If

   TxtValidate = True
End Function

' 更新資料
Private Function SaveData(strEditMode As Integer) As Boolean
Dim strKEY01 As String, strKEY02 As String, strKEY03 As String
Dim bDifference As Boolean
   
On Error GoTo ErrHand
   
   SaveData = False
   bDifference = False
   
   strKEY01 = txtED01
   If Option1(0).Value = True Then
      strKEY02 = LblED02_1
      strKEY03 = LblED03
   Else
      strKEY02 = txtED02_2
      strKEY03 = txtED03
   End If
   
   If strEditMode = 1 Then
      If IsRecordExist(strKEY01, strKEY02) = True Then
         MsgBox "此分機資料已存在！", vbExclamation
         Exit Function
      End If
   ElseIf strEditMode = 2 Then
      If m_PKey1 <> strKEY01 Or m_PKey2 <> strKEY02 Then
         If IsRecordExist(strKEY01, strKEY02) = True Then
            MsgBox "修改資料，但此分機資料已存在！", vbExclamation
            Exit Function
         End If
      End If
   End If
   
   '新增
   If strEditMode = 1 Then
      bDifference = True
      If Option1(0).Value = True Then
         strSql = "INSERT INTO ExtensionData(ED01,ED02,ED03,ED04,ED05,ED06)" & _
                  " VALUES(" & CNULL(strKEY01) & "," & CNULL(strKEY02) & "," & CNULL(LblED03) & _
                  "," & CNULL(txtED04) & "," & CNULL(txtED05) & "," & CNULL(txtED06) & ")"
      Else
         strSql = "INSERT INTO ExtensionData(ED01,ED02,ED03,ED04,ED05,ED06)" & _
                  " VALUES(" & CNULL(strKEY01) & "," & CNULL(strKEY02) & "," & CNULL(txtED03) & _
                  ",null," & CNULL(txtED05) & "," & CNULL(txtED06) & ")"
      End If
   '修改
   ElseIf strEditMode = 2 Then
      If Option1(0).Value = True Then
         If m_PKey1 <> strKEY01 Or _
            m_PKey2 <> strKEY02 Or _
            LblED03.Tag <> LblED03.Caption Or _
            txtED04.Tag <> txtED04.Text Or _
            txtED05.Tag <> txtED05.Text Or _
            txtED06.Tag <> txtED06.Text Then
            bDifference = True
         End If
         strSql = "UPDATE ExtensionData" & _
                  " SET ED01=" & CNULL(strKEY01) & _
                      ",ED02=" & CNULL(strKEY02) & _
                      ",ED03=" & CNULL(LblED03) & _
                      ",ED04=" & CNULL(txtED04) & _
                      ",ED05=" & CNULL(txtED05) & _
                      ",ED06=" & CNULL(txtED06) & _
                  " WHERE ED01='" & m_PKey1 & "' and ED02='" & m_PKey2 & "'"
      Else
         If m_PKey1 <> strKEY01 Or _
            m_PKey2 <> strKEY02 Or _
            txtED03.Tag <> txtED03.Text Or _
            txtED05.Tag <> txtED05.Text Or _
            txtED06.Tag <> txtED06.Text Then
            bDifference = True
         End If
         strSql = "UPDATE ExtensionData" & _
                  " SET ED01=" & CNULL(strKEY01) & _
                      ",ED02=" & CNULL(strKEY02) & _
                      ",ED03=" & CNULL(txtED03) & _
                      ",ED04=null" & _
                      ",ED05=" & CNULL(txtED05) & _
                      ",ED06=" & CNULL(txtED06) & _
                  " WHERE ED01='" & m_PKey1 & "' and ED02='" & m_PKey2 & "'"
      End If
   End If
   If bDifference = True Then
      cnnConnection.BeginTrans
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      cnnConnection.CommitTrans
   End If
   
   If (strKEY03 & strKEY01 & strKEY02 < m_FirstKEY(2) & m_FirstKEY(0) & m_FirstKEY(1)) Or _
      (strKEY03 & strKEY01 & strKEY02 > m_LastKEY(2) & m_LastKEY(0) & m_LastKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strKEY01, strKEY02, strKEY03
   
   SaveData = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strKEY01 As String, strKEY02 As String
   
On Error GoTo ErrHand
   
   DelRecord = False
   
   strKEY01 = txtED01
   If Option1(0).Value = True Then
      strKEY02 = LblED02_1
   Else
      strKEY02 = txtED02_2
   End If
   
   cnnConnection.BeginTrans
   
   strSql = "DELETE FROM ExtensionData WHERE ED01 = " & CNULL(strKEY01) & " and ED02 = " & CNULL(strKEY02)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
'intQKind=0.查詢
'         1.分機號碼望遠鏡
'         2.員工姓名望遠鏡
Private Function QueryRecord(intQKind As Integer) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String, strCon2 As String
   
   QueryRecord = False
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   strCon = ""
   strCon2 = ""
   Select Case intQKind
      Case 0
         If txtED01.Text <> "" Then
            strCon = strCon & " and ED01='" & txtED01 & "' "
            strCon2 = strCon2 & " and ED01='" & txtED01 & "' "
         End If
         If txtED02_1 <> "" And txtED02_2 <> "" Then
            strCon = strCon & " and (ST01='" & txtED02_1 & "' or ST02='" & txtED02_1 & "')"
            strCon2 = strCon2 & " and ED02='" & txtED02_2 & "'"
         End If
         If txtED02_1 <> "" Then
            strCon = strCon & " and (ST01='" & txtED02_1 & "' or ST02='" & txtED02_1 & "')"
            strCon2 = strCon2 & " and ED02='" & txtED02_1 & "'"
         End If
         If txtED02_2 <> "" Then
            strCon = strCon & " and (ST01='" & txtED02_2 & "' or ST02='" & txtED02_2 & "')"
            strCon2 = strCon2 & " and ED02='" & txtED02_2 & "'"
         End If
      Case 1
         If txtED01.Text <> "" Then
            strCon = strCon & " and ED01='" & txtED01 & "' "
            strCon2 = strCon2 & " and ED01='" & txtED01 & "' "
         End If
      Case 2
         If txtED02_1 <> "" Then
            'Modify By Sindy 2016/2/22
            'strCon = strCon & " and (ST01='" & txtED02_1 & "' or ST02='" & txtED02_1 & "')"
            strCon = strCon & " and (ST01='" & txtED02_1 & "' or instr(ST02,'" & txtED02_1 & "')>0)"
            '2016/2/22 END
            strCon2 = strCon2 & " and ED02='" & txtED02_1 & "'"
         End If
   End Select
   
   grd1.Rows = 2
   grd1.Clear
   dblPrevRow = 0
   'Added by Lydia 2023/12/26
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,NVL(A0922,A0902) AS a0902,ED02,decode(st04,'1','','Y'),st02,ED04,ED05,ED06,ED03,ED03||ED01,ED03||ST03,ED03||ST01,ED03||ST02,ED03||ED05,ED03||ED06,ED03||ST04,ED03||ED04" & _
               " From ExtensionData,staff,acc090,ACC090NEW" & _
               " where ED02=st01 and st03=a0901(+) AND ST93=A0921(+) and st04='1'" & strCon & m_ConSql & _
               " Union" & _
               " select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,'','','',ED02,ED04,ED05,ED06,ED03,ED03||ED01,'','',ED03||ED02,ED03||ED05,ED03||ED06,'',ED03||ED04" & _
               " From ExtensionData,staff" & _
               " where ED02=st01(+) and st01 is null" & strCon2 & m_ConSql
   Else
   'end 2023/12/26
      strSql = "select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,a0902,ED02,decode(st04,'1','','Y'),st02,ED04,ED05,ED06,ED03,ED03||ED01,ED03||ST03,ED03||ST01,ED03||ST02,ED03||ED05,ED03||ED06,ED03||ST04,ED03||ED04" & _
               " From ExtensionData,staff,acc090" & _
               " where ED02=st01 and st03=a0901(+) and st04='1'" & strCon & m_ConSql & _
               " Union" & _
               " select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,'','','',ED02,ED04,ED05,ED06,ED03,ED03||ED01,'','',ED03||ED02,ED03||ED05,ED03||ED06,'',ED03||ED04" & _
               " From ExtensionData,staff" & _
               " where ED02=st01(+) and st01 is null" & strCon2 & m_ConSql
   End If
   Select Case intQKind
      Case 0
         strSql = strSql & " order by ED03 asc,ED01 asc,ED02 asc"
      Case 1
         strSql = strSql & " order by ED03 asc,ED01 asc,ED02 asc"
      Case 2
         strSql = strSql & " order by ED03 asc,ED02 asc,ED01 asc"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
      QueryRecord = True
      rsTmp.MoveFirst
      m_CurrKEY(0) = rsTmp.Fields("ED01")
      If "" & rsTmp.Fields("ED02") = "" Then
         m_CurrKEY(1) = rsTmp.Fields("st02")
      Else
         m_CurrKEY(1) = rsTmp.Fields("ED02")
      End If
      m_CurrKEY(2) = rsTmp.Fields("ED03")
      UpdateCtrlData
      dblPrevRow = 1
   End If
   rsTmp.Close
   SetDataListWidth
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   UpdateToolbarState
   
EXITSUB:
   Set rsTmp = Nothing
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
Dim Cancel As Boolean
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = True Then
             RefreshRange
             ReadAllData
             SetKeyReadOnly True
         Else
             Exit Function
         End If
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = False Then Exit Function
         ReadAllData
         SetKeyReadOnly True
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
            ReadAllData
            SetKeyReadOnly True
         Else
            Exit Function
         End If
      Case 4: '查詢
         If txtED01.Text = "" Then
            MsgBox "分機號碼不可以空白！", vbExclamation
            txtED01.SetFocus
            Exit Function
         End If
         If txtED02_1 = "" And txtED02_2 = "" Then
            MsgBox "員工姓名及場地名稱至少要輸入一項！", vbExclamation
            If txtED02_1.Enabled = True Then txtED02_1.SetFocus
            If txtED02_2.Enabled = True Then txtED02_2.SetFocus
            Exit Function
         Else
            If QueryRecord(0) = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
               ReadAllData
            End If
            SetKeyReadOnly True
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
   
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 0, 1, 4: If Me.txtED01.Visible = True Then txtED01.SetFocus
      Case 2
         If Option1(0).Value = True Then
            txtED04.SetFocus
         Else
            txtED03.SetFocus
         End If
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM ExtensionData WHERE ED01=" & CNULL(strKEY01) & " and ED02=" & CNULL(strKEY02)
   
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      m_CurrKEY(2) = strKEY03
   Else
      strSql = "select ED01,ED02,ED03 from ExtensionData where ED01='" & m_CurrKEY(0) & "' and ED02='" & m_CurrKEY(1) & "' and ED03='" & m_CurrKEY(2) & "'" & m_ConSql
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "select ED01,ED02,ED03 from ExtensionData " & m_ConWhereSql & " order by ED03 asc,ED01 asc,ED02 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   m_CurrKEY(2) = m_FirstKEY(2)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "select ED01,ED02,ED03 from ExtensionData where ED03||ED01||ED02<'" & m_CurrKEY(2) & m_CurrKEY(0) & m_CurrKEY(1) & "'" & m_ConSql & _
            " order by ED03 desc,ED01 desc,ED02 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "select ED01,ED02,ED03 from ExtensionData " & m_ConWhereSql & " order by ED03 asc,ED01 asc,ED02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "select ED01,ED02,ED03 from ExtensionData where ED03||ED01||ED02>'" & m_CurrKEY(2) & m_CurrKEY(0) & m_CurrKEY(1) & "'" & m_ConSql & _
            " order by ED03 asc,ED01 asc,ED02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "select ED01,ED02,ED03 from ExtensionData" & m_ConWhereSql & " order by ED03 asc,ED01 asc,ED02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   m_CurrKEY(2) = m_LastKEY(2)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         SetKeyReadOnly False
         UpdateToolbarState
         SetInputEntry
         Option1(0).Value = True
         Call Option1_Click(0)
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         Command1.Enabled = True
         Command2.Enabled = True
         
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
         Option1(0).Value = True
         Call Option1_Click(0)
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  SetKeyReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               ReadAllData
               UpdateCtrlData
               SetCtrlReadOnly True
               SetKeyReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   Command1.Enabled = Not bEnable
   Command2.Enabled = Not bEnable
   'Add By Sindy 2014/5/6 修改時開放可以改PKey資料
   If m_EditMode = 2 Then
      Option1(0).Enabled = bEnable
      Option1(1).Enabled = bEnable
      txtED01.Locked = Not bEnable
      If Not bEnable Then txtED01.BackColor = &H8000000F Else txtED01.BackColor = &H80000005
      txtED02_1.Locked = Not bEnable
      If Not bEnable Then txtED02_1.BackColor = &H8000000F Else txtED02_1.BackColor = &H80000005
      txtED02_2.Locked = Not bEnable
      If Not bEnable Then txtED02_2.BackColor = &H8000000F Else txtED02_2.BackColor = &H80000005
   Else
   '2014/5/6 END
      Option1(0).Enabled = Not bEnable
      Option1(1).Enabled = Not bEnable
      txtED01.Locked = bEnable
      If bEnable Then txtED01.BackColor = &H8000000F Else txtED01.BackColor = &H80000005
      txtED02_1.Locked = bEnable
      If bEnable Then txtED02_1.BackColor = &H8000000F Else txtED02_1.BackColor = &H80000005
      txtED02_2.Locked = bEnable
      If bEnable Then txtED02_2.BackColor = &H8000000F Else txtED02_2.BackColor = &H80000005
   End If
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   txtED03.Locked = bEnable
   If bEnable Then txtED03.BackColor = &H8000000F Else txtED03.BackColor = &H80000005
   txtED04.Locked = bEnable
   If bEnable Then txtED04.BackColor = &H8000000F Else txtED04.BackColor = &H80000005
   txtED05.Locked = bEnable
   If bEnable Then txtED05.BackColor = &H8000000F Else txtED05.BackColor = &H80000005
   txtED06.Locked = bEnable
   If bEnable Then txtED06.BackColor = &H8000000F Else txtED06.BackColor = &H80000005
End Sub

Private Sub ClearField()
   txtED01 = Empty
   txtED02_1 = Empty
   txtED02_2 = Empty
   txtED03 = Empty
   txtED04 = Empty
   txtED05 = Empty
   txtED06 = Empty
   LblED02_1 = Empty
   LblED03 = Empty
   LblDept = Empty
   Label23 = Empty
   '記錄從DB查詢出來的欄位值
   LblED03.Tag = Empty
   txtED03.Tag = Empty
   txtED04.Tag = Empty
   txtED05.Tag = Empty
   txtED06.Tag = Empty
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         txtED02_1.Enabled = True
         txtED04.Enabled = True
         txtED02_2.Enabled = False
         txtED03.Enabled = False
      Case 1
         txtED02_1.Enabled = False
         txtED04.Enabled = False
         txtED02_2.Enabled = True
         txtED03.Enabled = True
   End Select
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ReadAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   m_EditMode = 0
   grd1.Rows = 2
   grd1.Clear
   'Added by Lydia 2023/12/26
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,NVL(A0922,A0902) AS a0902,ED02,decode(st04,'1','','Y'),st02,ED04,ED05,ED06,ED03,ED03||ED01,ED03||ST03,ED03||ST01,ED03||ST02,ED03||ED05,ED03||ED06,ED03||ST04,ED03||ED04" & _
               " From ExtensionData,staff,acc090,ACC090NEW" & _
               " where ED02=st01 and st03=a0901(+) AND ST93=A0921(+) and st04='1'" & m_ConSql & _
               " Union" & _
               " select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,'','','',ED02,ED04,ED05,ED06,ED03,ED03||ED01,'','',ED03||ED02,ED03||ED05,ED03||ED06,'',ED03||ED04" & _
               " From ExtensionData,staff" & _
               " where ED02=st01(+) and st01 is null" & m_ConSql & _
               " order by ED03 asc,ED01 asc,ED02 asc"
   Else
   'end 2023/12/26
      strSql = "select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,a0902,ED02,decode(st04,'1','','Y'),st02,ED04,ED05,ED06,ED03,ED03||ED01,ED03||ST03,ED03||ST01,ED03||ST02,ED03||ED05,ED03||ED06,ED03||ST04,ED03||ED04" & _
               " From ExtensionData,staff,acc090" & _
               " where ED02=st01 and st03=a0901(+) and st04='1'" & m_ConSql & _
               " Union" & _
               " select decode(ED03,'1','北','2','中','3','南','4','高','其他'),ED01,'','','',ED02,ED04,ED05,ED06,ED03,ED03||ED01,'','',ED03||ED02,ED03||ED05,ED03||ED06,'',ED03||ED04" & _
               " From ExtensionData,staff" & _
               " where ED02=st01(+) and st01 is null" & m_ConSql & _
               " order by ED03 asc,ED01 asc,ED02 asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
   End If
   rsTmp.Close
   SetDataListWidth
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub GetSelChage()
grd1.Visible = False
If grd1.Rows - 1 > 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 And dblPrevRow <= grd1.Rows - 1 Then
      grd1.col = 2
      grd1.row = dblPrevRow
      For i = 0 To 1
         grd1.col = i
         grd1.CellBackColor = &H8000000F
      Next i
      For i = 0 To grd1.Cols - 1
         grd1.col = i
         grd1.CellBackColor = QBColor(15)
      Next i
   End If
   '尋找目前資料列
   For j = 1 To grd1.Rows - 1
      If grd1.TextMatrix(j, 1) = m_CurrKEY(0) And _
         IIf(grd1.TextMatrix(j, 3) <> "", grd1.TextMatrix(j, 3), grd1.TextMatrix(j, 5)) = m_CurrKEY(1) And _
         grd1.TextMatrix(j, 9) = m_CurrKEY(2) Then
         grd1.col = 0
         grd1.row = j
         dblPrevRow = grd1.row
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
'            grd1.TopRow = j
         Exit For
      End If
   Next j
End If
grd1.Visible = True
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   ClearField
   'Added by Lydia 2023/12/26
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "select ED01,ED02,ED03,ED04,ED05,ED06,ED07,ED08,ED09,ED10,ED11,ED12,0 as kind,st02,st06,st03,NVL(A0922,A0902) AS a0902" & _
               " From ExtensionData,staff,acc090,ACC090NEW" & _
               " where ED02=st01 and st03=a0901(+) AND ST93=A0921(+) and st04='1'" & _
               " and ED01='" & m_CurrKEY(0) & "' and ED02='" & m_CurrKEY(1) & "'" & m_ConSql & _
               " Union" & _
               " select ED01,ED02,ED03,ED04,ED05,ED06,ED07,ED08,ED09,ED10,ED11,ED12,1 as kind,'','','',''" & _
               " From ExtensionData,staff" & _
               " where ED02=st01(+) and st01 is null" & _
               " and ED01='" & m_CurrKEY(0) & "' and ED02='" & m_CurrKEY(1) & "'" & m_ConSql
   Else
   'end 2023/12/26
      strSql = "select ED01,ED02,ED03,ED04,ED05,ED06,ED07,ED08,ED09,ED10,ED11,ED12,0 as kind,st02,st06,st03,a0902" & _
               " From ExtensionData,staff,acc090" & _
               " where ED02=st01 and st03=a0901(+) and st04='1'" & _
               " and ED01='" & m_CurrKEY(0) & "' and ED02='" & m_CurrKEY(1) & "'" & m_ConSql & _
               " Union" & _
               " select ED01,ED02,ED03,ED04,ED05,ED06,ED07,ED08,ED09,ED10,ED11,ED12,1 as kind,'','','',''" & _
               " From ExtensionData,staff" & _
               " where ED02=st01(+) and st01 is null" & _
               " and ED01='" & m_CurrKEY(0) & "' and ED02='" & m_CurrKEY(1) & "'" & m_ConSql
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ED01")) = False Then txtED01 = rsTmp.Fields("ED01")
      m_PKey1 = txtED01 'Add By Sindy 2014/5/6
      If rsTmp.Fields("kind") = 0 Then
         Option1(0).Value = True
         Call Option1_Click(0)
         If IsNull(rsTmp.Fields("st02")) = False Then txtED02_1 = rsTmp.Fields("st02")
         If IsNull(rsTmp.Fields("ED01")) = False Then LblED02_1 = rsTmp.Fields("ED02")
         m_PKey2 = LblED02_1 'Add By Sindy 2014/5/6
         If IsNull(rsTmp.Fields("st06")) = False Then LblED03 = rsTmp.Fields("st06"): LblED03.Tag = rsTmp.Fields("st06")
         If IsNull(rsTmp.Fields("a0902")) = False Then LblDept = rsTmp.Fields("a0902")
         If IsNull(rsTmp.Fields("ED04")) = False Then txtED04 = rsTmp.Fields("ED04"): txtED04.Tag = rsTmp.Fields("ED04")
      Else
         Option1(1).Value = True
         Call Option1_Click(1)
         If IsNull(rsTmp.Fields("ED02")) = False Then txtED02_2 = rsTmp.Fields("ED02")
         m_PKey2 = txtED02_2 'Add By Sindy 2014/5/6
         If IsNull(rsTmp.Fields("ED03")) = False Then txtED03 = rsTmp.Fields("ED03"): txtED03.Tag = rsTmp.Fields("ED03")
      End If
      If IsNull(rsTmp.Fields("ED05")) = False Then txtED05 = rsTmp.Fields("ED05"): txtED05.Tag = rsTmp.Fields("ED05")
      If IsNull(rsTmp.Fields("ED06")) = False Then txtED06 = rsTmp.Fields("ED06"): txtED06.Tag = rsTmp.Fields("ED06")
      
      ' 更新CUID
      UpdateCUID rsTmp
   End If
   rsTmp.Close
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("ED07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ED07")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("ED07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ED08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ED08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("ED08"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ED09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ED09")) = False Then
         strTemp = rsSrcTmp.Fields("ED09")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ED10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ED10")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("ED10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ED11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ED11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("ED11"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ED12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ED12")) = False Then
         strTemp = rsSrcTmp.Fields("ED12")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(5, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "select ED01,ED02,ED03 from ExtensionData" & m_ConWhereSql & " order by ED03 asc,ED01 asc,ED02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_FirstKEY(0) = Trim(rsTmp.Fields(0))
      If IsNull(rsTmp.Fields(1)) = False Then m_FirstKEY(1) = Trim(rsTmp.Fields(1))
      If IsNull(rsTmp.Fields(2)) = False Then m_FirstKEY(2) = Trim(rsTmp.Fields(2))
   End If
   rsTmp.Close
   
   strSql = "select ED01,ED02,ED03 from ExtensionData" & m_ConWhereSql & " order by ED03 desc,ED01 desc,ED02 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_LastKEY(0) = Trim(rsTmp.Fields(0))
      If IsNull(rsTmp.Fields(1)) = False Then m_LastKEY(1) = Trim(rsTmp.Fields(1))
      If IsNull(rsTmp.Fields(2)) = False Then m_LastKEY(2) = Trim(rsTmp.Fields(2))
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
   '鎖住上下筆功能
   TBar1.Buttons(6).Enabled = False
   TBar1.Buttons(7).Enabled = False
   TBar1.Buttons(8).Enabled = False
   TBar1.Buttons(9).Enabled = False
End Sub

Private Sub SetDataListWidth()
grd1.row = 0
grd1.col = 0: grd1.Text = "所別"
grd1.ColWidth(0) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 1: grd1.Text = "分機號碼"
grd1.ColWidth(1) = 800
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 2: grd1.Text = "部門"
grd1.ColWidth(2) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 3: grd1.Text = "員工編號"
grd1.ColWidth(3) = 800
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 4: grd1.Text = "離職"
grd1.ColWidth(4) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 5: grd1.Text = "姓名或場地"
grd1.ColWidth(5) = 1500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 6: grd1.Text = "英文別名"
grd1.ColWidth(6) = 800
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 7: grd1.Text = "樓層"
grd1.ColWidth(7) = 1000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 8: grd1.Text = "備註"
grd1.ColWidth(8) = 2500
grd1.CellAlignment = flexAlignLeftCenter
'以下欄位是為了排序
grd1.col = 9: grd1.Text = "ED03"
grd1.ColWidth(9) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 10: grd1.Text = "ED03||ED01"
grd1.ColWidth(10) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 11: grd1.Text = "ED03||ST03"
grd1.ColWidth(11) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 12: grd1.Text = "ED03||ST01"
grd1.ColWidth(12) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 13: grd1.Text = "ED03||ST02"
grd1.ColWidth(13) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 14: grd1.Text = "ED03||ED05"
grd1.ColWidth(14) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 15: grd1.Text = "ED03||ED06"
grd1.ColWidth(15) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 16: grd1.Text = "ED03||ST04"
grd1.ColWidth(16) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 17: grd1.Text = "ED03||ED04"
grd1.ColWidth(17) = 0
grd1.CellAlignment = flexAlignLeftCenter
End Sub

Private Sub txtED01_GotFocus()
   InverseTextBox txtED01
End Sub

Private Sub txtED01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtED02_1_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtED02_1
      OpenIme
   End If
End Sub

'Modify by Amy 2021/08/02 改Form2.0 原:KeyAscii As Integer
Private Sub txtED02_1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtED02_1_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtED02_1 <> "" Then
      If CheckLengthIsOK(txtED02_1, txtED02_1.MaxLength) = False Then
         Call txtED02_1_GotFocus
         Cancel = True
         Exit Sub
      End If
      
      If m_EditMode <> 4 Then
         '員工編號
         If txtED02_1 > "6" And txtED02_1 < "F" Then
            If ClsPDGetStaff(txtED02_1, strExc(1), strExc(2), , strExc(3)) Then
               LblED02_1 = txtED02_1
               txtED02_1 = strExc(1)
               LblDept = strExc(2)
               LblED03 = strExc(3)
            Else
               Call txtED02_1_GotFocus
               Cancel = True
               Exit Sub
            End If
         '員工姓名
         Else
            txtED02_1 = Trim(txtED02_1) 'Added by Lydia 2024/02/19 因為用複製可能會貼上空白
            '依員工姓名抓取員工編號
            LblED02_1 = GetPrjSalesNM_2(txtED02_1, strExc(2), strExc(3))
            If LblED02_1 <> "" Then
               LblDept = strExc(2)
               LblED03 = strExc(3)
            Else
               MsgBox "找不到此員工 或 此員工已離職！", vbExclamation
               Call txtED02_1_GotFocus
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   End If
   CloseIme
End Sub

Private Sub txtED02_2_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtED02_2
      OpenIme
   End If
End Sub

Private Sub txtED02_2_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtED02_2 <> "" Then
      If CheckLengthIsOK(txtED02_2, txtED02_2.MaxLength) = False Then
         Call txtED02_2_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   CloseIme
End Sub

Private Sub txtED03_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtED03
   End If
End Sub

Private Sub txtED03_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtED03 <> "" Then
      If txtED03 <> "1" And txtED03 <> "2" And txtED03 <> "3" And txtED03 <> "4" Then
         MsgBox "所別只可輸入1 ~ 4", vbExclamation
         Call txtED03_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtED04_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtED04
   End If
End Sub

Private Sub txtED04_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtED04 <> "" Then
      If CheckLengthIsOK(txtED04, txtED04.MaxLength) = False Then
         Call txtED04_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtED05_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtED05
      OpenIme
   End If
End Sub

Private Sub txtED05_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtED05 <> "" Then
      If CheckLengthIsOK(txtED05, txtED05.MaxLength) = False Then
         Call txtED05_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   CloseIme
End Sub

Private Sub txtED06_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtED06
      OpenIme
   End If
End Sub

Private Sub txtED06_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtED06 <> "" Then
      If CheckLengthIsOK(txtED06, txtED06.MaxLength) = False Then
         Call txtED06_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   CloseIme
End Sub
