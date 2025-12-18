VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140416 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部關聯企業分類維護"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8295
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7695
      Top             =   450
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
            Picture         =   "frm140416.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140416.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   90
      TabIndex        =   3
      Top             =   720
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "資料維護"
      TabPicture(0)   =   "frm140416.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "textCUID"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtField(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtField(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm140416.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Index           =   2
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1020
         Width           =   6240
      End
      Begin VB.TextBox txtField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   0
         Top             =   540
         Width           =   600
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
         Bindings        =   "frm140416.frx":212C
         Height          =   3825
         Left            =   -74910
         TabIndex        =   4
         Top             =   420
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   6747
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSForms.TextBox textCUID 
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   3990
         Width           =   7860
         VariousPropertyBits=   671105055
         Size            =   "13864;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "修改說明內容時需考慮已使用資料"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "關聯代號："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   585
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "說明內容："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frm140416"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/23 改成Form2.0 ; textCUID
'Created by Lydia 2016/11/09 國外部關聯企業分類維護
Option Explicit

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim oText As TextBox

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
         KeyCode = 0: Action 14
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
  
   MoveFormToCenter Me
   
   Action 6 '預設第一筆
   Call SetGrid(True)
   UpdateToolbarState
   
   textCUID.BackColor = &H8000000F
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140416 = Nothing
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
         If m_bUpdate And txtField(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtField(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtField(1) <> "" Then
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
      For Each oText In txtField
         oText.Enabled = False
      Next
      SSTab1.TabEnabled(1) = True
   Case Else
      If m_EditMode = 4 Or m_EditMode = 1 Then
         For Each oText In txtField
           oText.Enabled = True
         Next
         txtField(1).SetFocus
      Else
         txtField(2).Enabled = True
         txtField(2).SetFocus
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
        SSTab1.TabEnabled(1) = False
        UpdateCUID 0
      Case 2 '按下修改
        m_EditMode = 2
        SSTab1.TabEnabled(1) = False
      Case 3 '按下刪除
         If txtField(1).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If
         '檢查國外部關聯企業資料FRelation
         strExc(0) = "select count(*) from FRelation where instr(FR03,'" & txtField(1) & "') > 0 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val(RsTemp(0)) > 0 Then
               MsgBox "尚有" & RsTemp(0) & "筆國外部關聯企業資料用到此代號，不可刪除!", vbExclamation
               Exit Sub
            End If
         End If
         If DelMsg() = True Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               ShowRecord 3
               Call SetGrid(False)
            End If
         End If

      Case 4 '按下查詢
         FormReset
         m_EditMode = 4
      Case 6 '第一筆
         ShowRecord 0
      Case 7 '前一筆
         ShowRecord 1
      Case 8 '後一筆
         ShowRecord 2
      Case 9 '最後筆
         ShowRecord 3
      Case 11 '按下確定
         Select Case m_EditMode
            '新增,修改
            Case 1, 2
               If TxtValidate = False Then
                  Exit Sub
               Else
                 If txtField(1).Text <> txtField(1).Tag Then
                    If RecIsExist = True Then Exit Sub
                 End If
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     m_EditMode = 0
                     For Each oText In txtField
                         oText.Tag = oText.Text
                     Next
                     ReadData txtField(1)
                     Call SetGrid(False)
                  End If
               End If
               SSTab1.TabEnabled(1) = True
            '查詢
            Case 4
               If ReadData(txtField(1)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         SSTab1.TabEnabled(1) = True
         For Each oText In txtField
             oText.Text = oText.Tag
         Next
         If txtField(1) <> "" Then
            If ReadData(txtField(1)) = False Then
               ShowRecord 3
            End If
         End If
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   UpdateToolbarState
   TxtLock
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

' 顯示資料
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
 Dim stKEY As String

On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(FT01),0) FROM FType "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(FT01),0) FROM FType where FT01 < " & CNULL(txtField(1))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(FT01),0) FROM FType where FT01>" & CNULL(txtField(1))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(FT01),0) FROM FType "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
   End Select
   
   If stKEY <> "" Then
      ReadData stKEY
      ShowRecord = True
   End If
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Private Function ReadData(Optional ByVal pKey01 As String) As Boolean
Dim stCon As String
Dim rsAD As New ADODB.Recordset
   'Modified by Lydia 2017/06/28 改成文字
   'If Val(pKey01) <> 0 Then stCon = stCon & "and FT01='" & pKey01 & "' "
   If Trim(pKey01) <> "" Then stCon = stCon & "and FT01='" & pKey01 & "' "
   
   FormReset

   strExc(0) = "select * from FType where 1=1 " & stCon & " order by FT01,FT02"
  
   intI = 1
   Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsAD.MoveFirst
      With rsAD
         For Each oText In txtField
            oText.Text = "" & .Fields("FT" & Format(oText.Index, "00"))
            oText.Tag = oText.Text
         Next
      End With
      ReadData = True
   End If
   
   UpdateCUID 1, rsAD
End Function

Private Sub SetGrid(ByVal bolShow As Boolean)
Dim rsD As New ADODB.Recordset
Dim idR As Integer, intR As Integer
    
    strExc(1) = "select FT01,FT02 from FType order by FT01 "
    intI = 0
    Set rsD = ClsLawReadRstMsg(intI, strExc(1))
    If intI = 1 Then
       Set MGrid1.Recordset = rsD
       MGrid1.FormatString = "關聯代號|說明"
       MGrid1.ColWidth(0) = 1100
       MGrid1.ColWidth(1) = 3600
       'Modified by Lydia 2017/06/28 靠左對齊
'       For idR = 2 To MGrid1.Cols - 1
'          MGrid1.ColWidth(idR) = 0
'       Next
       For intR = 0 To rsD.RecordCount
          For idR = 0 To MGrid1.Cols - 1
             MGrid1.row = intR
             MGrid1.col = idR
             MGrid1.CellAlignment = flexAlignLeftCenter
         Next idR
       Next intR
       'end 2017/06/28
       If bolShow = True Then
          SSTab1.Tab = 1
       Else
          SSTab1.Tab = 0
       End If
    End If
         
End Sub
' 更新 Create 及 Update 的人
Private Sub FormReset()
   Dim oText As TextBox
   
   For Each oText In txtField
      oText.Text = ""
      'oText.Tag = "" 'Remove by Lydia 2020/04/20 影響還原上一筆
   Next
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   TextInverse txtField(Index)
   If Index = 1 Then
      CloseIme
   End If
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 1 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer

   Cancel = False
   If m_EditMode = 0 Or m_EditMode = 4 Then Exit Sub
   If txtField(Index) = "" Then Exit Sub 'Added by Lydia 2020/04/20 欄位空白,不檢查輸入值
   
   Select Case Index
        Case 1
           If txtField(Index) = "" Then
              MsgBox "請輸入關聯代號!", vbCritical, "輸入錯誤"
              txtField(Index).SetFocus
              Cancel = True
           End If
        Case 2
           If txtField(Index) = "" Then
              MsgBox "請輸入說明!", vbCritical, "輸入錯誤"
              txtField(Index).SetFocus
              Cancel = True
           End If
   End Select
   
   If Not CheckLengthIsOK(txtField(Index), iLen) Then
      Cancel = True
   End If
   
   Exit Sub

End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean

   'Added by Lydia 2020/04/20 檢查欄位是否空白
   If Trim(txtField(1)) = "" Then
       MsgBox "關聯代號不可空白!", vbCritical, "輸入錯誤"
       txtField(1).SetFocus
       txtField_GotFocus 1
       Exit Function
   End If
   If Trim(txtField(2)) = "" Then
       MsgBox "說明內容不可空白!", vbCritical, "輸入錯誤"
       txtField(2).SetFocus
       txtField_GotFocus 2
       Exit Function
   End If
   'end 2020/04/20
   
   For Each oText In txtField
        txtField_Validate oText.Index, bCancel
        If bCancel = True Then
           txtField(oText.Index).SetFocus
           Exit Function
        End If
   Next
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
       If m_EditMode = 1 Then '新增
          'Modified by Lydia 2017/06/28 + PUB_StringFilter
          strSql = "insert into FType(FT01,FT02,FT03,FT04,FT05)" & _
             " Values ('" & PUB_StringFilter(txtField(1).Text) & "','" & ChgSQL(PUB_StringFilter(txtField(2).Text)) & "','" & strUserNum & "'," & strSrvDate(1) & "," & Mid(Format(ServerTime, "000000"), 1, 4) & ") "
       Else         '修改
          'Modified by Lydia 2017/06/28 + PUB_StringFilter
          strSql = "update FType set FT02='" & ChgSQL(PUB_StringFilter(txtField(2).Text)) & "', FT06='" & strUserNum & "', FT07=" & strSrvDate(1) & ", FT08=" & Mid(Format(ServerTime, "000000"), 1, 4) & _
                   " where FT01='" & txtField(1).Text & "' "
       End If
       Pub_SeekTbLog strSql '寫入維護記錄檔
       cnnConnection.Execute strSql, intI

   cnnConnection.CommitTrans
   FormSave = True
   
   Exit Function
   
ErrHnd:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Private Sub MGrid1_DblClick()
   If MGrid1.row > 0 And MGrid1.TextMatrix(MGrid1.row, 0) <> "" Then
      If ReadData(MGrid1.TextMatrix(MGrid1.row, 0)) Then
         SSTab1.Tab = 0
      End If
   End If
End Sub

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
      strSql = "delete from FType where FT01='" & txtField(1) & "' "
      Pub_SeekTbLog strSql '寫入維護記錄檔
      cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

'判斷是否存在相同分類的記錄
Private Function RecIsExist() As Boolean
   Dim iR As Integer
   Dim rsQa As ADODB.Recordset
   Dim strCon As String
   
If Trim(txtField(1)) <> "" Then
   strCon = strCon & "and FT01='" & txtField(1) & "' "
End If

If Left(strCon, 3) = "and" Then strCon = Mid(strCon, 4, Len(strCon) - 4)

   strExc(1) = " select * from FType where " & strCon
   iR = 1
   Set rsQa = ClsLawReadRstMsg(iR, strExc(1))
   If iR = 1 Then
      RecIsExist = True
      MsgBox "已存在同樣關聯代號的記錄，請先查詢!!", vbCritical
   Else
      RecIsExist = False
   End If
   Set rsQa = Nothing
   
End Function


' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByVal actType As Integer, Optional ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If actType = 0 Then
      strCName = GetStaffName(strUserNum, True)
      strCDate = Format(strSrvDate(2), "###/##/##")
      strCTime = ""
   Else
        If IsNull(rsSrcTmp.Fields("FT03")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("FT03")) = False Then
              strCName = GetStaffName(rsSrcTmp.Fields("FT03"), True)
           End If
        End If
        If IsNull(rsSrcTmp.Fields("FT04")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("FT04")) = False Then
              strTemp = TAIWANDATE(rsSrcTmp.Fields("FT04"))
              strCDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("FT05")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("FT05")) = False Then
              strTemp = rsSrcTmp.Fields("FT05")
              strCTime = Format(strTemp, "00:00")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("FT06")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("FT06")) = False Then
              strUName = GetStaffName(rsSrcTmp.Fields("FT06"), True)
           End If
        End If
        If IsNull(rsSrcTmp.Fields("FT07")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("FT07")) = False Then
              strTemp = TAIWANDATE(rsSrcTmp.Fields("FT07"))
              strUDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("FT08")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("FT08")) = False Then
              strTemp = rsSrcTmp.Fields("FT08")
              strUTime = Format(strTemp, "00:00")
           End If
        End If
   End If
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & " " & vbTab & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

