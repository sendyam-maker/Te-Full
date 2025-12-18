VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140116 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "參考名條/不列印名單/新舊縣市名稱維護作業"
   ClientHeight    =   5745
   ClientLeft      =   180
   ClientTop       =   990
   ClientWidth     =   8955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140116.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   30
      TabIndex        =   10
      Top             =   660
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8916
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "參考名條"
      TabPicture(0)   =   "frm140116.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtRN01"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtRN03"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MSHFlexGrid1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtRN02"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "不列印名單"
      TabPicture(1)   =   "frm140116.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtTBNP01"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MSHFlexGrid1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "新舊縣市名稱"
      TabPicture(2)   =   "frm140116.frx":212C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(15)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(11)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtNOA01"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtNOA02"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "MSHFlexGrid1(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.TextBox txtRN02 
         Height          =   270
         Left            =   -73800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3945
         Index           =   2
         Left            =   450
         TabIndex        =   8
         Top             =   1020
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   6959
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "舊名稱               | 新名稱                "
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
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4335
         Index           =   1
         Left            =   -74850
         TabIndex        =   5
         Top             =   630
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   7646
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   1
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "特定公司名稱"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3675
         Index           =   0
         Left            =   -74910
         TabIndex        =   3
         Top             =   1320
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   6482
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   3
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "公司名稱         | 郵遞區號      | 地址                "
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
         _Band(0).Cols   =   3
      End
      Begin MSForms.TextBox txtNOA02 
         Height          =   300
         Left            =   1260
         TabIndex        =   7
         Top             =   690
         Width           =   2595
         VariousPropertyBits=   679495707
         Size            =   "4577;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNOA01 
         Height          =   300
         Left            =   1260
         TabIndex        =   6
         Top             =   390
         Width           =   2595
         VariousPropertyBits=   679495707
         Size            =   "4577;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTBNP01 
         Height          =   300
         Left            =   -73500
         TabIndex        =   4
         Top             =   330
         Width           =   7215
         VariousPropertyBits=   679495707
         Size            =   "12726;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRN03 
         Height          =   300
         Left            =   -73800
         TabIndex        =   2
         Top             =   990
         Width           =   7155
         VariousPropertyBits=   679495707
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRN01 
         Height          =   300
         Left            =   -73800
         TabIndex        =   0
         Top             =   420
         Width           =   7155
         VariousPropertyBits=   679495707
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "郵遞區號："
         Height          =   180
         Index           =   3
         Left            =   -74730
         TabIndex        =   16
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Index           =   1
         Left            =   -74730
         TabIndex        =   15
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公司名稱："
         Height          =   180
         Index           =   0
         Left            =   -74730
         TabIndex        =   14
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特定公司名稱："
         Height          =   180
         Index           =   2
         Left            =   -74760
         TabIndex        =   13
         Top             =   390
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "舊名稱："
         Height          =   180
         Index           =   11
         Left            =   510
         TabIndex        =   12
         Top             =   450
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新名稱："
         Height          =   180
         Index           =   15
         Left            =   510
         TabIndex        =   11
         Top             =   750
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm140116"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/1 Form2.0已修改
'Created by Sindy 2013/3/11
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的本所案號
Dim m_FirstKEY As String
' 最後一筆資料的本所案號
Dim m_LastKEY As String
' 目前正在顯示的本所案號
Dim m_CurrKEY As String
Dim i As Integer, j As Integer
Dim dblCurrRow As Double


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 13 Then
'      '擋掉Enter動作
'      KeyCode = 0
'      Exit Sub
'   End If
   
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   MoveFormToCenter Me
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
'   m_bOpen = IsUserHasRightOfFunction(Me.Name, strPrint, False)
   
   SSTab1.Tab = 0
   Call CallNewDrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140116 = Nothing
End Sub

Private Sub CallNewDrid()
   m_FirstKEY = ""
   m_LastKEY = ""
   m_CurrKEY = ""
   dblCurrRow = 0
   RefreshRange
   QueryAllData
   ShowFirstRecord
   UpdateToolbarState
   SetKeyReadOnly True
   SetCtrlReadOnly True
   'OnAction vbKeyF4 '按查詢
   OnAction vbKeyF10 '按取消
End Sub

Private Sub MSHFlexGrid1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow MSHFlexGrid1(Index), X, Y, nCol, nRow
   MSHFlexGrid1(Index).col = nCol
   MSHFlexGrid1(Index).row = nRow
End Sub

Private Sub MSHFlexGrid1_Click(Index As Integer)
Dim tmpMouseRow
Dim strText As String
   
   MSHFlexGrid1(Index).Visible = False
   tmpMouseRow = MSHFlexGrid1(Index).row
   If tmpMouseRow <> 0 Then
      '查詢資料
      m_CurrKEY = MSHFlexGrid1(Index).TextMatrix(tmpMouseRow, 1)
      Call UpdateCtrlData(m_CurrKEY)
   End If
   MSHFlexGrid1(Index).Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Call CallNewDrid
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
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   Select Case SSTab1.Tab
   Case 0
      If Me.txtRN01.Enabled = True Then
         Cancel = False
         txtRN01_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      If Me.txtRN02.Enabled = True Then
         Cancel = False
         txtRN02_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      If Me.txtRN03.Enabled = True Then
         Cancel = False
         txtRN03_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Case 1
      If Me.txtTBNP01.Enabled = True Then
         Cancel = False
         txtTBNP01_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Case 2
      If Me.txtNOA01.Enabled = True Then
         Cancel = False
         txtNOA01_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      If Me.txtNOA02.Enabled = True Then
         Cancel = False
         txtNOA02_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   End Select
   
   'Add by Sindy 2021/6/1 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/6/1 END
   
   TxtValidate = True
End Function

' 新增記錄
Private Function AddRecord() As Boolean
Dim strKey1 As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim bolReSave As Boolean
   
   AddRecord = False
   bolReSave = False
   
   Select Case SSTab1.Tab
   Case 0
      strKey1 = Trim(txtRN01)
   Case 1
      strKey1 = Trim(txtTBNP01)
   Case 2
      strKey1 = Trim(txtNOA01)
   End Select
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strKey1) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Call UpdateCtrlData(strKey1)
      Exit Function
   End If
   
ReSave:
   Select Case SSTab1.Tab
   Case 0
      strSql = "insert into ReferenceNames(rn01,rn02,rn03) values(" & CNULL(strKey1) & "," & CNULL(Trim(txtRN02)) & "," & CNULL(Trim(txtRN03)) & ")"
   Case 1
      strSql = "insert into TMBulletinNp(tbnp01,tbnp08) values(" & CNULL(strKey1) & ",'A')"
   Case 2
      strSql = "insert into NewOldAddr(noa01,noa02) values(" & CNULL(strKey1) & "," & CNULL(Trim(txtNOA02)) & ")"
   End Select
   
On Error GoTo ErrHand
   'cnnConnection.BeginTrans
   cnnConnection.Execute strSql
   'cnnConnection.CommitTrans
   
   If ((strKey1) < (m_FirstKEY)) Or ((strKey1) > (m_LastKEY)) Then
      RefreshRange
   End If
   ShowCurrRecord strKey1
   AddRecord = True
   Exit Function
   
ErrHand:
   'cnnConnection.RollbackTrans
   If Err.Number = -2147217900 And bolReSave = False Then '造字錯誤,必須最後加空白才可存入DB
      bolReSave = True
      Select Case SSTab1.Tab
      Case 0
         strKey1 = Trim(txtRN01) & " "
      Case 1
         strKey1 = Trim(txtTBNP01) & " "
      Case 2
         strKey1 = Trim(txtNOA01) & " "
      End Select
      GoTo ReSave
   End If
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
   ModRecord = False
   
   Select Case SSTab1.Tab
   Case 0
      strSql = "update ReferenceNames Set rn02='" & Trim(txtRN02) & "',rn03='" & Trim(txtRN03) & "' WHERE rn01='" & m_CurrKEY & "'"
   Case 1
      Exit Function
   Case 2
      strSql = "update NewOldAddr Set noa02='" & Trim(txtNOA02) & "' WHERE noa01='" & m_CurrKEY & "'"
   End Select
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   ShowCurrRecord m_CurrKEY
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   
   DelRecord = False
   
   Select Case SSTab1.Tab
   Case 0
      m_CurrKEY = txtRN01
      strSql = "DELETE FROM ReferenceNames WHERE rn01 = '" & m_CurrKEY & "'  "
   Case 1
      m_CurrKEY = txtTBNP01
      strSql = "DELETE FROM TMBulletinNp WHERE tbnp01 = '" & m_CurrKEY & "' and tbnp08='A'"
   Case 2
      m_CurrKEY = txtNOA01
      strSql = "DELETE FROM NewOldAddr WHERE noa01 = '" & m_CurrKEY & "'  "
   End Select
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   If (m_CurrKEY = m_LastKEY) Or (m_CurrKEY = m_FirstKEY) Then
      RefreshRange
   End If
   ShowCurrRecord m_CurrKEY
   DelRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Public Function QueryRecord() As Boolean
Dim strKey As String
   
   Select Case SSTab1.Tab
   Case 0
      strKey = Trim(txtRN01.Text)
   Case 1
      strKey = Trim(txtTBNP01.Text)
   Case 2
      strKey = Trim(txtNOA01.Text)
   End Select
   QueryRecord = False
   If IsRecordExist(strKey) = True Then
      QueryRecord = True
      Call UpdateCtrlData(strKey)
      Exit Function
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            If AddRecord = True Then
'               RefreshRange
            Else
               Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
'            RefreshRange
'            ClearField
'            ShowCurrRecord m_CurrKEY
         Else
            Exit Function
         End If
      Case 4: '查詢
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Call UpdateCtrlData(m_CurrKEY)
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case SSTab1.Tab
   Case 0
      Select Case m_EditMode
         Case 1: If Me.Visible = True Then txtRN01.SetFocus
         Case 2: If Me.Visible = True Then txtRN02.SetFocus
         Case 4: If Me.Visible = True Then txtRN01.SetFocus
      End Select
   Case 1
      Select Case m_EditMode
         Case 1: If Me.Visible = True Then txtTBNP01.SetFocus
         Case 4: If Me.Visible = True Then txtTBNP01.SetFocus
      End Select
   Case 2
      Select Case m_EditMode
         Case 1: If Me.Visible = True Then txtNOA01.SetFocus
         Case 2: If Me.Visible = True Then txtNOA02.SetFocus
         Case 4: If Me.Visible = True Then txtNOA01.SetFocus
      End Select
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
      
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE ltrim(rtrim(rn01))=ltrim(rtrim('" & strKEY01 & " " & "'))"
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE ltrim(rtrim(tbnp01))=ltrim(rtrim('" & strKEY01 & " " & "')) and tbnp08='A'"
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE ltrim(rtrim(noa01))=ltrim(rtrim('" & strKEY01 & " " & "'))"
   End Select
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY = strKEY01
   Else
      Select Case SSTab1.Tab
      Case 0
         strSql = "SELECT * FROM ReferenceNames " & _
                  "WHERE rn01='" & m_CurrKEY & "'"
      Case 1
         strSql = "SELECT * FROM TMBulletinNp " & _
                  "WHERE tbnp01='" & m_CurrKEY & "' and tbnp08='A'"
      Case 2
         strSql = "SELECT * FROM NewOldAddr " & _
                  "WHERE noa01='" & m_CurrKEY & "'"
      End Select
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY = rsTmp.Fields(0)
         rsTmp.Close
         'Call UpdateCtrlData(m_CurrKEY)
         QueryAllData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      Select Case SSTab1.Tab
      Case 0
         strSql = "SELECT * FROM ReferenceNames " & _
                  "WHERE rn01=(SELECT MIN(rn01) FROM ReferenceNames)"
      Case 1
         strSql = "SELECT * FROM TMBulletinNp " & _
                  "WHERE tbnp01=(SELECT MIN(tbnp01) FROM TMBulletinNp WHERE tbnp08='A') and tbnp08='A'"
      Case 2
         strSql = "SELECT * FROM NewOldAddr " & _
                  "WHERE noa01=(SELECT MIN(noa01) FROM NewOldAddr)"
      End Select
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY = rsTmp.Fields(0)
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   
   'Call UpdateCtrlData(m_CurrKEY)
   QueryAllData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY = m_FirstKEY
   Call UpdateCtrlData(m_CurrKEY)
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY = m_FirstKEY Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE rn01=(SELECT MAX(rn01) FROM ReferenceNames " & _
                           "WHERE rn01<'" & m_CurrKEY & "')"
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE tbnp01=(SELECT MAX(tbnp01) FROM TMBulletinNp " & _
                           "WHERE tbnp01<'" & m_CurrKEY & "' and tbnp08='A') and tbnp08='A'"
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE noa01=(SELECT MAX(noa01) FROM NewOldAddr " & _
                           "WHERE noa01<'" & m_CurrKEY & "')"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY = rsTmp.Fields(0)
      rsTmp.Close
      Call UpdateCtrlData(m_CurrKEY)
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE rn01=(SELECT Min(rn01) FROM ReferenceNames)"
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE tbnp01=(SELECT Min(tbnp01) FROM TMBulletinNp WHERE tbnp08='A') and tbnp08='A'"
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE noa01=(SELECT Min(noa01) FROM NewOldAddr)"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY = rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   Call UpdateCtrlData(m_CurrKEY)
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY = m_LastKEY Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE rn01=(SELECT MIN(rn01) FROM ReferenceNames " & _
                            "WHERE rn01>'" & m_CurrKEY & "')"
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE tbnp01=(SELECT MIN(tbnp01) FROM TMBulletinNp " & _
                            "WHERE tbnp01>'" & m_CurrKEY & "' and tbnp08='A') and tbnp08='A'"
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE noa01=(SELECT MIN(noa01) FROM NewOldAddr " & _
                            "WHERE noa01>'" & m_CurrKEY & "')"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY = rsTmp.Fields(0)
      rsTmp.Close
      Call UpdateCtrlData(m_CurrKEY)
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE rn01=(SELECT max(rn01) FROM ReferenceNames)"
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE tbnp01=(SELECT max(tbnp01) FROM TMBulletinNp WHERE tbnp08='A') and tbnp08='A'"
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE noa01=(SELECT max(noa01) FROM NewOldAddr)"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY = rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   Call UpdateCtrlData(m_CurrKEY)
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY = m_LastKEY
   Call UpdateCtrlData(m_CurrKEY)
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
         SetKeyReadOnly False
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         SSTab1.TabEnabled(0) = False
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
         SSTab1.TabEnabled(SSTab1.Tab) = True
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
         SSTab1.TabEnabled(0) = False
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
         SSTab1.TabEnabled(SSTab1.Tab) = True
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否確定要刪除此筆資料?"
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
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
         SSTab1.TabEnabled(0) = False
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
         SSTab1.TabEnabled(SSTab1.Tab) = True
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
            SetKeyReadOnly True
            SetCtrlReadOnly True
            UpdateToolbarState
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = True
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
                  Call UpdateCtrlData(m_CurrKEY)
                  SetKeyReadOnly True
                  SetCtrlReadOnly True
                  UpdateToolbarState
                  SSTab1.TabEnabled(0) = True
                  SSTab1.TabEnabled(1) = True
                  SSTab1.TabEnabled(2) = True
               End If
            Case Else
               If m_EditMode <> 0 Then
                  m_EditMode = 0
                  Call UpdateCtrlData(m_CurrKEY)
                  SetKeyReadOnly True
                  SetCtrlReadOnly True
                  UpdateToolbarState
                  SSTab1.TabEnabled(0) = True
                  SSTab1.TabEnabled(1) = True
                  SSTab1.TabEnabled(2) = True
               End If
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE rn01=(SELECT MIN(rn01) FROM ReferenceNames)"
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE tbnp01=(SELECT MIN(tbnp01) FROM TMBulletinNp WHERE tbnp08='A') and tbnp08='A'"
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE noa01=(SELECT MIN(noa01) FROM NewOldAddr)"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_FirstKEY = rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE rn01=(SELECT MAX(rn01) FROM ReferenceNames)"
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE tbnp01=(SELECT MAX(tbnp01) FROM TMBulletinNp WHERE tbnp08='A') and tbnp08='A'"
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE noa01=(SELECT MAX(noa01) FROM NewOldAddr)"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_LastKEY = rsTmp.Fields(0)
   End If
   rsTmp.Close
      
   Set rsTmp = Nothing
End Sub

' 將點選的資料更新到畫面欄位中
Private Sub UpdateCtrlData(strKey As String)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim tmpMouseRow
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   SSTab1.Enabled = False
   
   ClearField
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT * FROM ReferenceNames " & _
               "WHERE ltrim(rtrim(rn01))=ltrim(rtrim('" & strKey & " " & "'))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CurrKEY = strKey
         If IsNull(rsTmp.Fields("rn01")) = False Then: txtRN01 = rsTmp.Fields("rn01")
         If IsNull(rsTmp.Fields("rn02")) = False Then: txtRN02 = rsTmp.Fields("rn02")
         If IsNull(rsTmp.Fields("rn03")) = False Then: txtRN03 = rsTmp.Fields("rn03")
      End If
   Case 1
      strSql = "SELECT * FROM TMBulletinNp " & _
               "WHERE ltrim(rtrim(tbnp01))=ltrim(rtrim('" & strKey & " " & "')) and tbnp08='A'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CurrKEY = strKey
         If IsNull(rsTmp.Fields("tbnp01")) = False Then: txtTBNP01 = rsTmp.Fields("tbnp01")
      End If
   Case 2
      strSql = "SELECT * FROM NewOldAddr " & _
               "WHERE ltrim(rtrim(noa01))=ltrim(rtrim('" & strKey & " " & "'))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CurrKEY = strKey
         If IsNull(rsTmp.Fields("noa01")) = False Then: txtNOA01 = rsTmp.Fields("noa01")
         If IsNull(rsTmp.Fields("noa02")) = False Then: txtNOA02 = rsTmp.Fields("noa02")
      End If
   End Select
   
   If rsTmp.RecordCount > 0 Then
      For i = 1 To MSHFlexGrid1(SSTab1.Tab).Rows - 1
         If MSHFlexGrid1(SSTab1.Tab).TextMatrix(i, 1) = strKey Then
            tmpMouseRow = i
            Exit For
         End If
      Next
      MSHFlexGrid1(SSTab1.Tab).Visible = False
      '反白
      If dblCurrRow > 0 Then
         MSHFlexGrid1(SSTab1.Tab).row = dblCurrRow
         For i = 0 To MSHFlexGrid1(SSTab1.Tab).Cols - 1
            MSHFlexGrid1(SSTab1.Tab).col = i
            MSHFlexGrid1(SSTab1.Tab).CellBackColor = QBColor(15)
         Next i
         MSHFlexGrid1(SSTab1.Tab).TextMatrix(dblCurrRow, 0) = ""
      End If
      '反藍
      If tmpMouseRow > 0 Then
         MSHFlexGrid1(SSTab1.Tab).row = tmpMouseRow
         For i = 0 To MSHFlexGrid1(SSTab1.Tab).Cols - 1
            MSHFlexGrid1(SSTab1.Tab).col = i
            MSHFlexGrid1(SSTab1.Tab).CellBackColor = &HFFC0C0
         Next i
         MSHFlexGrid1(SSTab1.Tab).TextMatrix(tmpMouseRow, 0) = "V"
      End If
      dblCurrRow = tmpMouseRow
      MSHFlexGrid1(SSTab1.Tab).Visible = True
   End If
   
   rsTmp.Close
   Me.Enabled = True
   SSTab1.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 讀取全部資料
Private Sub QueryAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   SSTab1.Enabled = False
   
   Call GridHead1(SSTab1.Tab)
   Select Case SSTab1.Tab
   Case 0
      strSql = "SELECT '' as V,rn01 as 公司名稱,rn02 as 郵遞區號,rn03 as 地址 FROM ReferenceNames " & _
               "order by rn01 asc"
   Case 1
      strSql = "SELECT '' as V,tbnp01 as 特定公司名稱 FROM TMBulletinNp WHERE tbnp08='A' " & _
               "order by tbnp01 asc"
   Case 2
      strSql = "SELECT '' as V,noa01 as 舊名稱,noa02 as 新名稱 FROM NewOldAddr " & _
               "order by noa01 asc"
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set MSHFlexGrid1(SSTab1.Tab).Recordset = rsTmp
      Call UpdateCtrlData(m_CurrKEY)
   End If
   rsTmp.Close
   Me.Enabled = True
   SSTab1.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Me.Enabled = False
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            If SSTab1.Tab = 1 Then
               TBar1.Buttons(2).Enabled = False
            Else
               TBar1.Buttons(2).Enabled = True
            End If
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
   Me.Enabled = True
End Sub

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTit As String
Dim strMsg As String
Dim intIndex As Integer, strText As String
   
   CheckDataValid = False
   
   Select Case SSTab1.Tab
   Case 0
      If IsEmptyText(txtRN01) = True Then
         strTit = "檢核資料"
         strMsg = "公司名稱不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtRN01.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(txtRN02) = True Then
         strTit = "檢核資料"
         strMsg = "郵遞區號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtRN02.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(txtRN03) = True Then
         strTit = "檢核資料"
         strMsg = "地址不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtRN03.SetFocus
         GoTo EXITSUB
      End If
   Case 1
      If IsEmptyText(txtTBNP01) = True Then
         strTit = "檢核資料"
         strMsg = "特定公司名稱不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTBNP01.SetFocus
         GoTo EXITSUB
      End If
   Case 2
      If IsEmptyText(txtNOA01) = True Then
         strTit = "檢核資料"
         strMsg = "舊名稱不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtNOA01.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(txtNOA02) = True Then
         strTit = "檢核資料"
         strMsg = "新名稱不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtNOA02.SetFocus
         GoTo EXITSUB
      End If
      If txtNOA01 = txtNOA02 Then
         strTit = "檢核資料"
         strMsg = "舊名稱與新名稱不可相同"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtNOA02.SetFocus
         GoTo EXITSUB
      End If
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   Select Case SSTab1.Tab
   Case 0
      txtRN01.Locked = bEnable
      If bEnable Then txtRN01.BackColor = &H8000000F Else txtRN01.BackColor = &H80000005
   Case 1
      txtTBNP01.Locked = bEnable
      If bEnable Then txtTBNP01.BackColor = &H8000000F Else txtTBNP01.BackColor = &H80000005
   Case 2
      txtNOA01.Locked = bEnable
      If bEnable Then txtNOA01.BackColor = &H8000000F Else txtNOA01.BackColor = &H80000005
   End Select
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   Select Case SSTab1.Tab
   Case 0
      txtRN02.Locked = bEnable
      txtRN03.Locked = bEnable
   Case 1
   Case 2
      txtNOA02.Locked = bEnable
   End Select
End Sub

Private Sub ClearField()
   txtRN01 = Empty
   txtRN02 = Empty
   txtRN03 = Empty
   txtTBNP01 = Empty
   txtNOA01 = Empty
   txtNOA02 = Empty
End Sub

Private Sub GridHead1(Index As Integer)
   MSHFlexGrid1(Index).Clear
   MSHFlexGrid1(Index).Rows = 2
   Select Case Index
   Case 0
      With MSHFlexGrid1(Index)
         .Visible = False
         .Cols = 4
         .row = 0
         .col = 0: .ColWidth(0) = 200: .Text = "V"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(0) = flexAlignCenterCenter
         
         .col = 1: .ColWidth(1) = 2800: .Text = "公司名稱"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(1) = flexAlignCenterCenter
         
         .col = 2: .ColWidth(2) = 800: .Text = "郵遞區號"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(2) = flexAlignCenterCenter
         
         .col = 3: .ColWidth(3) = 4500: .Text = "地址"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(3) = flexAlignCenterCenter
         
         .Visible = True
      End With
   Case 1
      With MSHFlexGrid1(Index)
         .Visible = False
         .Cols = 2
         .row = 0
         .col = 0: .ColWidth(0) = 200: .Text = "V"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(0) = flexAlignCenterCenter
         
         .col = 1: .ColWidth(1) = 8000: .Text = "特定公司名稱"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(1) = flexAlignCenterCenter
         
         .Visible = True
      End With
   Case 2
      With MSHFlexGrid1(Index)
         .Visible = False
         .Cols = 3
         .row = 0
         .col = 0: .ColWidth(0) = 200: .Text = "V"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(0) = flexAlignCenterCenter
         
         .col = 1: .ColWidth(1) = 1500: .Text = "舊名稱"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(1) = flexAlignCenterCenter
         
         .col = 2: .ColWidth(2) = 1500: .Text = "新名稱"
         .CellAlignment = flexAlignCenterCenter
         .ColAlignment(2) = flexAlignCenterCenter
         
         .Visible = True
      End With
   End Select
End Sub

Private Sub txtNOA01_GotFocus()
   InverseTextBox txtNOA01
   OpenIme
End Sub

Private Sub txtNOA01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNOA01_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtNOA01
End Sub

Private Sub txtNOA01_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtNOA01.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtNOA01, txtNOA01.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub txtNOA02_GotFocus()
   InverseTextBox txtNOA02
   OpenIme
End Sub

Private Sub txtNOA02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNOA02_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtNOA01
End Sub

Private Sub txtNOA02_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtNOA02.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtNOA02, txtNOA02.MaxLength) Then
      Cancel = True
   End If
End Sub

'Add By Sindy 2021/6/1
Private Sub txtRN01_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtRN01
End Sub

Private Sub txtRN03_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtRN03
End Sub

Private Sub txtTBNP01_GotFocus()
   InverseTextBox txtTBNP01
   OpenIme
End Sub

Private Sub txtTBNP01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtTBNP01_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtTBNP01
End Sub

Private Sub txtTBNP01_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtTBNP01.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtTBNP01, txtTBNP01.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub txtRN01_GotFocus()
   InverseTextBox txtRN01
   OpenIme
End Sub

Private Sub txtRN01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtRN01_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtRN01.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtRN01, txtRN01.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub txtRN02_GotFocus()
   InverseTextBox txtRN02
   CloseIme
End Sub

Private Sub txtRN02_KeyPress(KeyAscii As Integer)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub txtRN02_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtRN02.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtRN02, txtRN02.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub txtRN03_GotFocus()
   InverseTextBox txtRN03
   OpenIme
End Sub

Private Sub txtRN03_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub txtRN03_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtRN03.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtRN03, txtRN03.MaxLength) Then
      Cancel = True
   End If
End Sub
