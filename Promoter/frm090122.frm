VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090122 
   BorderStyle     =   1  '單線固定
   Caption         =   "查名人員維護"
   ClientHeight    =   5076
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7524
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5076
   ScaleWidth      =   7524
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "修正今日統計量"
      Height          =   255
      Index           =   1
      Left            =   4920
      MaskColor       =   &H80000004&
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "修正前2日統計量"
      Height          =   255
      Index           =   0
      Left            =   360
      MaskColor       =   &H80000004&
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   4800
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7215
      _ExtentX        =   12721
      _ExtentY        =   6795
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   1764
      TabCaption(0)   =   "資料"
      TabPicture(0)   =   "frm090122.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textCUID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "GRD1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDB(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDB(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDB(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDB(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "排班順序"
      TabPicture(1)   =   "frm090122.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MGRD2"
      Tab(1).Control(1)=   "Label2"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtDB 
         Height          =   270
         Index           =   3
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   2
         Top             =   795
         Width           =   345
      End
      Begin VB.TextBox txtDB 
         Height          =   270
         Index           =   0
         Left            =   120
         MaxLength       =   6
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtDB 
         Height          =   270
         Index           =   1
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   700
      End
      Begin VB.TextBox txtDB 
         Height          =   270
         Index           =   2
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   1
         Top             =   480
         Width           =   700
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm090122.frx":0038
         Height          =   2145
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   6705
         _ExtentX        =   11832
         _ExtentY        =   3789
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   2
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
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGRD2 
         Bindings        =   "frm090122.frx":004D
         Height          =   2505
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         Width           =   6825
         _ExtentX        =   12044
         _ExtentY        =   4424
         _Version        =   393216
         Cols            =   20
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   $"frm090122.frx":0062
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
         _Band(0).Cols   =   20
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   5080
         TabIndex        =   18
         Top             =   510
         Width           =   975
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1720;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   2070
         TabIndex        =   17
         Top             =   510
         Width           =   975
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1720;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   300
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3420
         Width           =   6540
         VariousPropertyBits=   671105055
         Size            =   "11536;529"
         Value           =   "Create ID:            Create Date:   "
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y: 有一人請假就不分派)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   4
         Left            =   4680
         TabIndex        =   13
         Top             =   840
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否一起請假："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   $"frm090122.frx":012D
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   -74880
         TabIndex        =   9
         Top             =   3200
         Width           =   6855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "查名人："
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   525
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "統計人員："
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   5
         Top             =   525
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   360
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
            Picture         =   "frm090122.frx":01F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":0515
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":0831
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":0A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":0D29
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":1045
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":1361
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":167D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":1999
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":1CB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090122.frx":1FD1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7524
      _ExtentX        =   13272
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
   End
   Begin VB.Label Label1 
      Caption         =   "查名人狀態：N=不分派查名單"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   4800
      Width           =   2535
   End
End
Attribute VB_Name = "frm090122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/01 改成Form2.0 ; GRD1改字型=新細明體-ExtB、MGRD2改字型=新細明體-ExtB、textCUID、Lable3(index)
'Memo by Lydia 2015/05/28 GRD1顯示所有記錄,Click會把該筆記錄帶入textbox,功能鍵可移動
'Created by Lydia 2015/05/26 新增-查名人員維護
Option Explicit
Dim intLastRow As Integer, intCols As Integer

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim oText As TextBox
'Added by Lydia 2015/10/27
Dim PrevCol As String
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Private Sub Cmd1_Click(Index As Integer)
    
    'Added by Lydia 2024/11/18 查名單(網中)
    If strSrvDate(1) >= 查名單網中系統啟用日 Then
          Call PUB_TMAtoTake("2", "", "", Trim(Index), True)
    Else
       If strSrvDate(1) >= 查名單網中系統平行測試 Then
          Call PUB_TMAtoTake("2", "", "", Trim(Index), True)
       End If
    'end 2024/11/18
       '原本共用模組,傳入Quser有值=>個人重新計算,不傳入Quser值=>所有人重新計算
       Call PUB_TMQtake("2", "", , , , , Index)
    End If
    
    'Added by Lydia 2015/09/22 +提示
    If Index = 0 Then
       MsgBox "修正前2日統計量 完成", vbInformation
    Else
       MsgBox "修正今日統計量 完成", vbInformation
    End If
    
    ReadData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2 '新增
         KeyCode = 0: Action 1
      Case vbKeyF3 '修改
         KeyCode = 0: Action 2
      Case vbKeyF4: '查詢
         KeyCode = 0: Action 4
      Case vbKeyF5 '刪除
         KeyCode = 0: Action 3
      Case vbKeyHome '第一筆
         KeyCode = 0: Action 6
      Case vbKeyPageUp '上一筆
         KeyCode = 0: Action 7
      Case vbKeyPageDown '下一筆
         KeyCode = 0: Action 8
      Case vbKeyEnd: '最後筆
         KeyCode = 0: Action 9
      Case vbKeyF9, vbKeyReturn '確定
         KeyCode = 0: Action 11
      Case vbKeyF10 '取消
         KeyCode = 0: Action 12
      Case vbKeyEscape '結束
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            KeyCode = 0: Action 14
         End If
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm090122", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090122", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090122", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090122", strFind, False)
  
   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   Action 6 '預設第一筆
   UpdateToolbarState
   
   If Pub_StrUserSt03 <> "M51" Then
      Me.SSTab1.TabVisible(1) = False
      Cmd1(0).Visible = False: Cmd1(1).Visible = False
   End If
   
   'ShowRecord 0
   ReadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090122 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub
'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
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
      
      Case 1, 2, 3, 4 '維護
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
End Sub

Private Sub TxtLock()
   Select Case m_EditMode
   Case 0 '瀏覽
      For Each oText In txtDB
         oText.Locked = True
      Next
      SSTab1.TabEnabled(1) = True
   Case Else
      For Each oText In txtDB
         oText.Locked = False
         oText.Tag = oText.Text
      Next
      If m_EditMode <> 4 Then
         If m_EditMode = "2" Then
            txtDB(1).Locked = True  '查名人員代號(PK)
            txtDB(2).SetFocus
            txtDB_GotFocus 2
         Else
            txtDB(1).SetFocus
            txtDB_GotFocus 1
         End If
      End If
      SSTab1.TabEnabled(1) = False
   End Select
End Sub
Private Sub Action(Index As Integer)
   
   If TBar1.Buttons(Index).Enabled = False Then Exit Sub

On Error GoTo ErrHand

   SSTab1.Tab = 0
   Select Case Index
      Case 1 '按下新增
        m_EditMode = 1
        FormReset
        
      Case 2 '按下修改
         If txtDB(1).Text = "" Or txtDB(2).Text = "" Then
             MsgBox "請先選擇資料!!!", vbExclamation + vbOKOnly
             Exit Sub
         Else
            m_EditMode = 2
         End If
      Case 3 '按下刪除
         If txtDB(1).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If
         If DelMsg() = True Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            Else
               ReadData '更新GRD1
            End If
         End If

      Case 4 '按下查詢
         FormReset
         m_EditMode = 4
      Case 6 '第一筆
         ShowRecord 0
      Case 7 '前一筆
         If txtDB(1) <> "" Then
            ShowRecord 1
         Else
            m_EditMode = -1
         End If
      Case 8 '後一筆
         If txtDB(1) <> "" Then
            ShowRecord 2
         Else
            m_EditMode = -1
         End If
      Case 9 '最後筆
         ShowRecord 3
      Case 11 '按下確定
         Select Case m_EditMode
            '新增,修改
            Case 1, 2
               If m_EditMode = 1 Then
                  If RecIsExist(True, txtDB(1), txtDB(2)) = True Then Exit Sub
               End If
               If TxtValidate = False Then
                  Exit Sub
               Else
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     m_EditMode = 0
                     ReadData
                  End If
               End If
            '查詢
            Case 4
               If RecIsExist(False, txtDB(1), txtDB(2)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
                  SetData txtDB(1) ', txtDB(2)
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         txtDB(1) = txtDB(1).Tag
         txtDB(2) = txtDB(2).Tag
         txtDB(3) = txtDB(3).Tag
         Label3(1).Caption = "": Label3(2).Caption = ""
         If txtDB(1) <> "" Then
            If RecIsExist(False, txtDB(1), txtDB(2)) = False Then
               ShowRecord 3
            End If
         End If
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   
   If m_EditMode < 0 Then
      m_EditMode = 0
   Else
      UpdateToolbarState
      TxtLock
   End If
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

' 顯示資料
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
 Dim stKey As String, stKey2 As String
  Dim mDiff As String
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = "SELECT * FROM TMQMember order by 2,1"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))

         If intI = 0 Then
            DataErrorMessage 8
         End If
         mDiff = "MIN"
         
      Case 1 '前一筆
         strExc(0) = "SELECT * FROM TMQMember where rownum<2 and TMQM01<" & CNULL(txtDB(1)) & " and TMQM02<" & CNULL(txtDB(2)) & _
                     " order by 2 desc,1 desc "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))

         If intI = 0 Then
            DataErrorMessage 6
         End If
         mDiff = "-1"
      Case 2 '後一筆
         strExc(0) = "SELECT * FROM TMQMember where rownum<2 and TMQM01>" & CNULL(txtDB(1)) & " and TMQM02>" & CNULL(txtDB(2)) & _
                     " order by 2 , 1 "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))

         If intI = 0 Then
            DataErrorMessage 7
         End If
         mDiff = "+1"
      Case 3 '最後筆
         strExc(0) = "SELECT * FROM TMQMember order by 2 desc,1 desc"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            DataErrorMessage 8
         End If
         mDiff = "MAX"
   End Select

         If intI = 1 Then
            txtDB(1) = RsTemp.Fields("TMQM01"): txtDB_Validate 1, False
            txtDB(2) = RsTemp.Fields("TMQM02"): txtDB_Validate 2, False
            txtDB(3) = "" & RsTemp.Fields("TMQM03")
            ShowRecord = True
            UpdateCUID RsTemp
         Else
            mDiff = ""
         End If
         
   Screen.MousePointer = vbDefault
   
   '功能鍵可移動反白列
   If intLastRow > 0 And mDiff <> "" Then
       GridClick GRD1, intLastRow, 7
          Select Case mDiff
              Case "MIN"
                 GRD1.row = 1
              Case "-1"
                 GRD1.row = intLastRow - 1
              Case "+1"
                 GRD1.row = intLastRow + 1
              Case "MAX"
                 GRD1.row = GRD1.Rows - 1
          End Select
       GridClick GRD1, intLastRow, 7
   End If
   
   Exit Function
  
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Private Function ReadData(Optional ByVal pKey As String, Optional ByVal pKey2 As String) As Boolean
   
   Dim stCon As String
   
   If pKey <> "" Then
      stCon = stCon & " and TMQM01='" & pKey & "'"
   End If
   If pKey2 <> "" Then
      stCon = stCon & " and TMQM02='" & pKey2 & "'"
   End If
  
   FormReset
   
   strExc(0) = "select TMQM01,(s1.st02) s1name,TMQM02,(s2.st02) s2name,(TMQSR17) type,TMQM03,TMQM04,TMQM05,TMQM06,TMQM07,TMQM08,TMQM09 " & _
               " from TMQMember a,staff s1,staff s2,TMQSUMR where 1=1 " & stCon & " and tmqm01=s1.st01(+) and tmqm02=s2.st01(+) " & _
               " and TMQM02=TMQSR01(+) order by TMQM02,TMQM01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ReadData = True
   End If
   Set GRD1.Recordset = RsTemp.Clone

   strExc(0) = "select ST02||' ('||TMQSR01||')' sname,TMQSR17,tmqsr12,tmqsr02,tmqsr07,tmqsr13,tmqsr03,tmqsr08,tmqsr14,tmqsr04,tmqsr09,tmqsr15,tmqsr05,tmqsr10,tmqsr16,tmqsr06,tmqsr11" & _
               ",TMQSR01,nvl(tmqm02||tmqm01,tmqsr01) or2 from TMQSumR a,staff s1,tmqmember where tmqsr01=tmqm01(+) and tmqsr01=s1.st01(+) order by or2 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   MGRD2.FixedCols = 0
   Set MGRD2.Recordset = RsTemp.Clone
   MGRD2.FixedCols = 2
   GridHead

End Function

Private Sub SetData(Optional ByVal p01 As String)
   Dim rsA As New ADODB.Recordset
   Dim strA1 As String, intA As Integer
   
   If Len(p01) > 0 Then
     strA1 = strA1 & " and TMQM01=" & CNULL(p01)
   End If
   
   strA1 = "select TMQM01,(s1.st02) s1name,TMQM02,(s2.st02) s2name,(TMQSR17) type,TMQM03,TMQM04,TMQM05,TMQM06,TMQM07,TMQM08,TMQM09 " & _
           " from TMQMember a,staff s1,staff s2,TMQSUMR where 1=1 " & strA1 & " and tmqm01=s1.st01(+) and tmqm02=s2.st01(+) " & _
           " and TMQM02=TMQSR01(+) order by TMQM02"
   
   If rsA.State <> adStateClosed Then rsA.Close
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, strA1)
   
   With rsA
     For Each oText In txtDB
        If oText.Index > 0 Then oText = "" & .Fields("TMQM" & Format(oText.Index, "00"))
     Next
   End With
   UpdateCUID rsA
   
   txtDB(1).Tag = txtDB(1)
   txtDB(2).Tag = txtDB(2)
   txtDB(3).Tag = txtDB(3)
   
   Label3(1).Caption = rsA.Fields("s1name")
   Label3(2).Caption = rsA.Fields("s2name")
   
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
   If IsNull(rsSrcTmp.Fields("TMQM04")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TMQM04")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("TMQM04"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TMQM05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TMQM05")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("TMQM05"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TMQM06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TMQM06")) = False Then
         strTemp = rsSrcTmp.Fields("TMQM06")
         strCTime = Format(strTemp, "00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TMQM07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TMQM07")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("TMQM07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TMQM08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TMQM08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("TMQM08"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("TMQM09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("TMQM09")) = False Then
         strTemp = rsSrcTmp.Fields("TMQM09")
         strUTime = Format(strTemp, "00:00")
      End If
   End If
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
  
End Sub

Private Sub FormReset()
   Dim oText As Control
   Dim oLabel As Control
   
   For Each oText In txtDB
      oText.Text = ""
   Next
   
   For Each oLabel In Label3
      oLabel.Caption = ""
   Next
   
   textCUID = ""
     '清除反白列
    If intLastRow > 0 Then
       If GRD1.CellBackColor <> GRD1.BackColor Then
         GridClick GRD1, intLastRow, 7
       End If
    End If
    
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 3 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
Dim ChkStaff As Boolean
Dim strCusTemp As String
   Select Case Index
   Case 1, 2
      If txtDB(Index) <> "" Then
         ChkStaff = ClsPDGetStaff(Trim(txtDB(Index)), strCusTemp)
         If ChkStaff = False Then
            Label3(Index).Caption = ""
            MsgBox "請輸入在職員工代號!", vbExclamation
            If m_EditMode <> 0 Then Cancel = True
         Else
            Label3(Index).Caption = strCusTemp
         End If
      End If

   End Select
   
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
   
   If txtDB(1) = "" Then
      MsgBox "查名人代號不可空白！", vbExclamation
      txtDB(2).SetFocus
      Exit Function
   End If
   If txtDB(2) = "" Then
      MsgBox "統計人員代號不可空白！", vbExclamation
      txtDB(2).SetFocus
      Exit Function
   End If
   
   For idx = 1 To 2
      txtDB_Validate idx, bCancel
      If bCancel = True Then
         txtDB(idx).SetFocus
         Exit Function
      End If
   Next
   
    If txtDB(1) = txtDB(2) Then
       If txtDB(3) <> "" Then
          MsgBox "一般查名人員不可設此欄!", vbExclamation
          Exit Function
       End If
    Else
       'Remove by Lydia 2017/06/23 統一設定
'       strSql = "select tmqm02,tmqm03 from tmqmember where TMQM02=" & CNULL(txtDB(2)) & " and TMQM01<>" & CNULL(txtDB(1)) & " group by tmqm02,tmqm03"
'       intI = 1
'       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'       If intI = 1 Then
'          If RsTemp.RecordCount > 1 Or txtDB(3).Text <> "" & RsTemp(1) Then
'              If MsgBox("統計人員是否一起請假的設定不一致,是否繼續?", vbExclamation + vbYesNo) = vbNo Then
'                 Exit Function
'              End If
'          End If
'       End If
    End If
      
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   Dim intN As Integer, intNum(0 To 4) As Integer
   Dim kindX As String
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   If m_EditMode = 1 Then  '新增
      strSql = "insert into TMQMember(TMQM01,TMQM02,TMQM03,TMQM04,TMQM05,TMQM06) values ('" & txtDB(1) & "','" & txtDB(2) & "','" & txtDB(3) & "'," & CNULL(strUserNum) & _
               "," & strSrvDate(1) & "," & CNULL(Left(Format(ServerTime, "000000"), 4), True) & ") "
   Else '修改
      strSql = "update TMQMember set TMQM02='" & txtDB(2) & "',TMQM03='" & txtDB(3) & "',TMQM07=" & CNULL(strUserNum) & _
               ",TMQM08=" & strSrvDate(1) & ",TMQM09=" & CNULL(Left(Format(ServerTime, "000000"), 4), True) & " where TMQM01=" & CNULL(txtDB(1))
   End If
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   If m_EditMode = 1 Then  '新增
        '增加統計量和排序的記錄
         For intN = 0 To 4
             Randomize Timer
             intNum(intN) = (Fix(Rnd() * 99) + 1) Mod 100
         Next intN
        
        kindX = IIf(txtDB(1).Text <> txtDB(2).Text, "2", "1")
        
        strSql = "insert into tmqsumr(TMQSR01,TMQSR02,TMQSR03,TMQSR04,TMQSR05,TMQSR06,TMQSR07,TMQSR08,TMQSR09,TMQSR10,TMQSR11,TMQSR12,TMQSR13,TMQSR14,TMQSR15,TMQSR16,TMQSR18) " & _
                 "VALUES ('" & txtDB(1) & "',0,0,0,0,0," & intNum(0) & "," & intNum(1) & "," & intNum(2) & "," & intNum(3) & "," & intNum(4) & ",0,0,0,0,0,'" & kindX & "')"
        cnnConnection.Execute strSql, intI
        '當是統計人員的第一筆，新增排班等級1的記錄
        If kindX = "2" Then
             strSql = "select count(*) from tmqmember where TMQM02=" & CNULL(txtDB(2))
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strSql)
             If RsTemp(0) = 1 Then
                 strSql = "insert into tmqsumr(TMQSR01,TMQSR02,TMQSR03,TMQSR04,TMQSR05,TMQSR06,TMQSR07,TMQSR08,TMQSR09,TMQSR10,TMQSR11,TMQSR12,TMQSR13,TMQSR14,TMQSR15,TMQSR16,TMQSR18) " & _
                      "VALUES ('" & txtDB(2) & "',0,0,0,0,0," & intNum(0) & "," & intNum(1) & "," & intNum(2) & "," & intNum(3) & "," & intNum(4) & ",0,0,0,0,0,'1')"
                 cnnConnection.Execute strSql, intI
             End If
        End If
   Else
   'Added by Lydia 2015/10/27 將拿單方式在小組與個人之間互換
        If txtDB(2).Text <> txtDB(2).Tag Then
           '小組改成個人
           If txtDB(1) = txtDB(2) Then
                strSql = "update tmqsumr set tmqsr18='1' where tmqsr01=" & CNULL(txtDB(1))
                cnnConnection.Execute strSql, intI
                strSql = "select count(*) from tmqmember where tmqm02='" & txtDB(2).Tag & "' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                'Memo by Lydia 2017/06/23 沒有人就刪除小組
                If RsTemp(0) = 0 Then
                   strSql = "delete from tmqsumr where tmqsr01='" & txtDB(2).Tag & "' "
                   cnnConnection.Execute strSql, intI
                End If
           Else '個人成小組
               kindX = "2"
               strSql = "update tmqsumr set tmqsr18='2' where tmqsr01=" & CNULL(txtDB(1))
               cnnConnection.Execute strSql, intI
               strSql = "select count(*) from tmqmember where TMQM02=" & CNULL(txtDB(2)) & " and tmqm01<>" & CNULL(txtDB(1))
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If RsTemp(0) = 0 Then
                    For intN = 0 To 4
                        Randomize Timer
                        intNum(intN) = (Fix(Rnd() * 99) + 1) Mod 100
                    Next intN
                    strSql = "insert into tmqsumr(TMQSR01,TMQSR02,TMQSR03,TMQSR04,TMQSR05,TMQSR06,TMQSR07,TMQSR08,TMQSR09,TMQSR10,TMQSR11,TMQSR12,TMQSR13,TMQSR14,TMQSR15,TMQSR16,TMQSR18) " & _
                         "VALUES ('" & txtDB(2) & "',0,0,0,0,0," & intNum(0) & "," & intNum(1) & "," & intNum(2) & "," & intNum(3) & "," & intNum(4) & ",0,0,0,0,0,'1')"
                    cnnConnection.Execute strSql, intI
                End If
           End If
        End If
   End If
   
   'Added by Lydia 2017/06/23 統一是否一起請假和狀態
    '統一是否一起請假(tmqm03)
     If txtDB(1) <> txtDB(2) Then
        strSql = "update tmqmember set tmqm03=" & IIf(Trim(txtDB(3)) = "", "NULL", "'Y'") & " where tmqm01 in (select tmqm01 from tmqmember where tmqm02='" & txtDB(2) & "') "
        cnnConnection.Execute strSql, intI
     End If
     '狀態
     strSql = "select tmqm02,count(tmqm01) cnt1,sum(decode(tmqm03,'Y',1,0)) cnt2,sum(decode(tmqsr17,null,0,1)) cnt3 from tmqmember,tmqsumr where tmqm01<>tmqm02 and tmqm01=tmqsr01(+) group by tmqm02"
     intI = 1
     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
     If intI = 1 Then
        RsTemp.MoveFirst
        Do While Not RsTemp.EOF
           '特定小組成員有一人請假,小組就不拿單;或非特定小組的成員全部請假,小組就不拿單
            If (Val("" & RsTemp.Fields("cnt2")) > 0 And Val("" & RsTemp.Fields("cnt3")) > 0) Or Val("" & RsTemp.Fields("cnt1")) = Val("" & RsTemp.Fields("cnt3")) Then
               strSql = "update tmqsumr set tmqsr17='N' where tmqsr01='" & RsTemp.Fields("tmqm02") & "' "
               cnnConnection.Execute strSql, intI
            Else
               strSql = "update tmqsumr set tmqsr17=null where tmqsr01='" & RsTemp.Fields("tmqm02") & "' "
               cnnConnection.Execute strSql, intI
            End If
           RsTemp.MoveNext
        Loop
     End If
    'end 2017/06/23
    
   cnnConnection.CommitTrans
   'Added by Lydia 2015/09/24 因為有針對新人在一天內重複新增後刪除又新增,做重計拿單量
   'Modified by Lydia 2015/10/26 因為有拿單方式在小組與個人之間互換的情況,做全體重計拿單量
'    Call PUB_TMQtake("2", txtDB(1), , , , , 0)
'    Call PUB_TMQtake("2", txtDB(1), , , , , 1)
     Call PUB_TMQtake("2", "", , , , , 0)
     Call PUB_TMQtake("2", "", , , , , 1)
   'end 2015/09/24

   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Sub Grd1_Click()
   
If m_EditMode = 0 Then
    GridClick GRD1, intLastRow, 7
    
    '帶入textbox
    If GRD1.TextMatrix(intLastRow, 0) <> "" Then
       ' SetData GRD1.TextMatrix(intLastRow, 0), GRD1.TextMatrix(intLastRow, 2)
        SetData GRD1.TextMatrix(intLastRow, 0)
    End If
End If

End Sub

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   strSql = "delete from TMQMember where TMQM01=" & CNULL(txtDB(1))
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   '刪除查名人的統計檔記錄
   strSql = "delete from TMQSUMR where TMQSR01=" & CNULL(txtDB(1))
   cnnConnection.Execute strSql, intI
   '當是統計人員的最後一筆時，刪除統計量和排序的記錄
   strSql = "select * from tmqmember where TMQM02=" & CNULL(txtDB(2))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
      strSql = " delete from tmqsumr where tmqsr01=" & CNULL(txtDB(2))
      cnnConnection.Execute strSql, intI
   End If
   
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function
'
Private Function RecIsExist(Optional ByVal bMsg As Boolean = True, Optional ByVal inS1 As String, Optional ByVal inS2 As String) As Boolean
   Dim iR As Integer
   Dim rsQa As ADODB.Recordset
strExc(0) = ""

If Trim(inS1) <> "" Then
   strExc(0) = strExc(0) & "and TMQM01='" & Trim(inS1) & "' "
End If

If Left(strExc(0), 3) = "and" Then strExc(0) = Mid(strExc(0), 4, Len(strExc(0)) - 4)

   strExc(1) = " select * from TMQMember where " & strExc(0) & " order by 1"
   iR = 1
   Set rsQa = ClsLawReadRstMsg(iR, strExc(1))
   If iR = 1 Then
      RecIsExist = True
      If bMsg = True Then MsgBox "查名人員記錄已存在!!", vbCritical
   Else
      RecIsExist = False
   End If
   Set rsQa = Nothing
   
End Function
Private Sub GridHead()
   Dim iR As Integer

    With GRD1
      .row = 0
      .col = 0: .ColWidth(0) = 1000: .Text = "查名人代號"
      .col = 1: .ColWidth(1) = 800: .Text = "姓名"
      .col = 2: .ColWidth(2) = 1200: .Text = "統計人員代號"
      .col = 3: .ColWidth(3) = 1000: .Text = "姓名"
      .col = 4: .ColWidth(4) = 1000: .Text = "查名人狀態"
      .col = 5: .ColWidth(5) = 1200: .Text = "是否一起請假"
      .ColAlignment = flexAlignCenterCenter
      For iR = 6 To GRD1.Cols - 1
         .col = iR: .ColWidth(iR) = 0
      Next
      .LeftCol = 2
    End With
    
    With MGRD2
      .FormatString = "統計人員|查名人狀態|當日文A單量|文A統計單量|文A亂數順序|當日文B單量|文B統計單量|文B亂數順序|當日圖A單量|圖A統計單量|圖A亂數順序|當日圖B單量|圖B統計單量|圖B亂數順序|當日圖C單量|圖C統計單量|圖C亂數順序"
      .col = 0: .ColWidth(0) = 1600
      For iR = 1 To 16
         .col = iR
      Next iR
      For iR = 17 To MGRD2.Cols - 1
         .col = iR:   .ColWidth(iR) = 0
      Next
      .ColAlignment = flexAlignCenterCenter
    End With

End Sub

'Added by Lydia 2015/10/27 + Grid點選排序
Private Sub MGRD2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MGRD2, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGRD2.col = nCol
   MGRD2.row = nRow
   If Me.MGRD2.row < 1 And Me.MGRD2.Text <> "V" Then
     ' If InStr("委查日期,期限日期,查覆日期,覆核日期", Me.Mgrd2.Text) > 0 Then
      If InStr("統計人員,查名人狀態", Me.MGRD2.Text) = 0 Then
         If m_blnColOrderAsc = True Then
            Me.MGRD2.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGRD2.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MGRD2.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGRD2.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
   PrevCol = Me.MGRD2.Text
End Sub

'end 2015/10/27
