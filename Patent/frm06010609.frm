VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010609 
   BorderStyle     =   1  '單線固定
   Caption         =   "行事曆分類資料維護"
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
            Picture         =   "frm06010609.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010609.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabPicture(0)   =   "frm06010609.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtField(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtField(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtField(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtField(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "多筆資料"
      TabPicture(1)   =   "frm06010609.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MGrid1"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
         Bindings        =   "frm06010609.frx":212C
         Height          =   3825
         Left            =   -74910
         TabIndex        =   6
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
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSForms.TextBox txtField 
         Height          =   810
         Index           =   4
         Left            =   1080
         TabIndex        =   3
         Top             =   1800
         Width           =   5580
         VariousPropertyBits=   -1466939365
         ScrollBars      =   2
         Size            =   "9842;1429"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   280
         Index           =   3
         Left            =   1080
         TabIndex        =   2
         Top             =   1380
         Width           =   600
         VariousPropertyBits=   679495707
         Size            =   "1058;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   280
         Index           =   2
         Left            =   1080
         TabIndex        =   1
         Top             =   960
         Width           =   600
         VariousPropertyBits=   679495707
         Size            =   "1058;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtField 
         Height          =   280
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Top             =   540
         Width           =   600
         VariousPropertyBits=   679495707
         Size            =   "1058;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "適用部門：                (1.外專程序 2.外專承辦)"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   585
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分　　類："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "說明："
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "細　　類："
         Height          =   180
         Index           =   5
         Left            =   135
         TabIndex        =   7
         Top             =   1425
         Width           =   900
      End
   End
End
Attribute VB_Name = "frm06010609"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Lydia 2015/12/22 國外部行事曆分類資料維護
'Memo by Lydia 2018/11/05 改成Form2.0 (Textbox)
'Memo by Lydia 2020/01/15 更名為「行事曆分類資料維護」
Option Explicit

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim iType As String '適用部門

'Modified by Lydia 2018/11/06 改成Form2.0
'Dim oText As TextBox
Dim oText As MSForms.TextBox
Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim SyxMsg As String 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)

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
  
   If Pub_StrUserSt03 = "M51" Then
       iType = "0"
   'Modified by Lydia 2016/06/30
   'ElseIf Pub_strUserST05 = "31" Then
   ElseIf Pub_StrUserSt03 = "F22" Then
       iType = "1"
   'Modified by Lydia 2016/06/30
   'ElseIf Pub_strUserST05 = "35" Then
   Else
       iType = "2"
   End If
   
   MoveFormToCenter Me
   
   'Added by Lydia 2018/11/20 模組-抓DB中的欄位實際長度
   For Each oText In txtField
          oText.MaxLength = PUB_GetFieldDefSize("STAFF_CALENDAR_TYPE", "SCT" & Format(oText.Index, "00"))
   Next
      
   Action 6 '預設第一筆
   Call SetGrid(True)
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010609 = Nothing
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
      For Each oText In txtField
         oText.Enabled = True
      Next
      If iType = "0" And (m_EditMode = 4 Or m_EditMode = 1) Then
         txtField(1).SetFocus
      Else
         If txtField(1) = "" Then txtField(1) = iType
         
         txtField(1).Enabled = False
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
      Case 2 '按下修改
        m_EditMode = 2
        SSTab1.TabEnabled(1) = False
      Case 3 '按下刪除
         If txtField(1).Text = "" Or txtField(2).Text = "" Or txtField(3).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
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
                 If txtField(1).Text <> txtField(1).Tag Or txtField(2).Text <> txtField(2).Tag Or txtField(3).Text <> txtField(3).Tag Then
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
                     ReadData txtField(1), txtField(2), txtField(3)
                     Call SetGrid(False)
                  End If
               End If
               SSTab1.TabEnabled(1) = True
            '查詢
            Case 4
               If ReadData(txtField(1), txtField(2), txtField(3)) = False Then
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
         If txtField(2) <> "" Then
            If ReadData(txtField(1), txtField(2), txtField(3)) = False Then
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
 Dim strSel As String

On Error GoTo ErrHand

   If iType = "0" Then
       strSel = IIf(Val(txtField(1)) = 0, "", txtField(1))
       If p_iWay = 0 Or p_iWay = 3 Then
          strSel = ""
       End If
   Else
       strSel = iType
   End If

   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(sct01||sct02||sct03),0) FROM staff_calendar_type "
         If strSel <> "" Then strExc(0) = strExc(0) & "where sct01=" & CNULL(strSel)

         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(sct01||sct02||sct03),0) FROM staff_calendar_type "
         If strSel <> "" Then strExc(0) = strExc(0) & "where sct01=" & CNULL(strSel) & " and sct02||sct03<" & CNULL(FdFmt(txtField(2)) & FdFmt(txtField(3)))
 
 
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(sct01||sct02||sct03),0) FROM staff_calendar_type "
         If strSel <> "" Then strExc(0) = strExc(0) & "where sct01=" & CNULL(strSel) & " and sct02||sct03>" & CNULL(FdFmt(txtField(2)) & FdFmt(txtField(3)))
      
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(sct01||sct02||sct03),0) FROM staff_calendar_type "
         If strSel <> "" Then strExc(0) = strExc(0) & "where sct01=" & CNULL(strSel)

         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
   End Select
     
   
   If stKEY <> "" Then
      ReadData Mid(stKEY, 1, 1), Mid(stKEY, 2, 2), Mid(stKEY, 4, 2)
      ShowRecord = True
   End If
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Private Function ReadData(Optional ByVal pKey01 As String, Optional ByVal pKey02 As String, Optional ByVal pKey03 As String) As Boolean
   Dim stCon As String

   If Val(pKey01) <> 0 Then stCon = stCon & "and sct01='" & pKey01 & "' "
   If Val(pKey02) <> 0 Then stCon = stCon & "and sct02='" & FdFmt(pKey02) & "' "
   If Val(pKey03) <> 0 Then stCon = stCon & "and sct03='" & FdFmt(pKey03) & "' "

   FormReset

   strExc(0) = "select * from staff_calendar_type where 1=1 " & stCon & " order by sct01,sct02"
  
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      With RsTemp
         For Each oText In txtField
            oText.Text = "" & .Fields("sct" & Format(oText.Index, "00"))
            oText.Tag = oText.Text
         Next
      End With
      ReadData = True
   End If
End Function

Private Sub SetGrid(ByVal bolShow As Boolean)
Dim rsD As New ADODB.Recordset
Dim idR As Integer

    If iType = "0" Then
       strExc(1) = ""
    Else
       strExc(1) = "and sct01=" & CNULL(iType)
    End If
    strExc(1) = "select decode(sct01,'1','程序','2','承辦') type,sct01,sct02,sct03,sct04 from staff_calendar_type where 1=1 " & strExc(1) & " order by sct01,sct02"
    intI = 0
    Set rsD = ClsLawReadRstMsg(intI, strExc(1))
    If intI = 1 Then
       Set MGrid1.Recordset = rsD
       MGrid1.FormatString = "適用部門|SCT01|分類|細類|說明"
       MGrid1.ColWidth(0) = 800
       MGrid1.ColWidth(1) = 0
       MGrid1.ColWidth(2) = 720
       MGrid1.ColWidth(3) = 720
       MGrid1.ColWidth(4) = 4600
       For idR = 5 To MGrid1.Cols - 1
          MGrid1.ColWidth(idR) = 0
       Next
       If bolShow = True Then
          SSTab1.Tab = 1
       Else
          SSTab1.Tab = 0
       End If
    End If
         
End Sub
' 更新 Create 及 Update 的人
Private Sub FormReset()
   'Remove by Lydia 2018/11/06 改在上面宣告
   'Dim oText As TextBox

   For Each oText In txtField
      oText.Text = ""
      oText.Tag = ""
   Next
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   'Modified by Lydia 2018/11/21 取消說明反白
   'TextInverse txtField(Index)
   If Index <> 4 Then TextInverse txtField(Index)

   If Index = 4 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

'Added by Lydia 2018/11/21
Private Sub txtField_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If Index = 4 Then
        Call PUB_HandleForm2TextBox(Me.txtField(4), Me.txtField(3), KeyCode, Shift) '模組化-統一控制
    End If
End Sub

'Added by Lydia 2018/11/21
Private Sub txtField_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If SyxMsg <> "txtField_" & Format(Index, "00") Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "txtField_" & Format(Index, "00")
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
    
End Sub

'Modified by Lydia 2018/11/06 改成Form2.0
'Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index <> 4 Then
      KeyAscii = Pub_NumAscii(KeyAscii)
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim strCusTemp As String, strTemp As String

   If m_EditMode = 0 Or m_EditMode = 4 Then Exit Sub
   Select Case Index
   Case 1
      If txtField(Index) <> "1" And txtField(Index) <> "2" Then
         MsgBox "請輸入1-2!", vbCritical, "輸入錯誤"
         txtField(Index).SetFocus
         Cancel = True
      End If

   Case 3
      If txtField(1) <> "" And txtField(2) <> "" And (txtField(1).Text <> txtField(1).Tag Or txtField(2).Text <> txtField(2).Tag Or txtField(3).Text <> txtField(3).Tag) Then
         If RecIsExist Then
            txtField(Index).SetFocus
            Cancel = True
         End If
      End If
   Case 4
      If txtField(Index) = "" Then
         MsgBox "請輸入說明!", vbCritical, "輸入錯誤"
         txtField(Index).SetFocus
         Cancel = True
      End If
   End Select
   
   Exit Sub

End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If iType <> "0" And iType <> txtField(1) Then
      MsgBox "適用部門錯誤！", vbExclamation
      txtField(1).SetFocus
      Exit Function
   End If
   
   txtField(2) = FdFmt(txtField(2))
   txtField(3) = FdFmt(txtField(3))
   For Each oText In txtField
        txtField_Validate oText.Index, bCancel
        If bCancel = True Then
           txtField(oText.Index).SetFocus
           Exit Function
        End If
   Next
    
    'Added by Lydia 2021/04/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True) = False Then
        Exit Function
    End If
    'end 2021/04/14
    
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   '修改時,先刪除後新增
       If m_EditMode > 1 Then
          strSql = "delete from staff_calendar_type where sct01='" & txtField(1).Tag & "' and sct02='" & FdFmt(txtField(2).Tag) & "' and sct03='" & FdFmt(txtField(3).Tag) & "'"
          cnnConnection.Execute strSql, intI
       End If
    
          strSql = "insert into staff_calendar_type(sct01,sct02,sct03,sct04)" & _
             " Values ('" & txtField(1).Text & "','" & FdFmt(txtField(2).Text) & "','" & FdFmt(txtField(3).Text) & "','" & ChgSQL(txtField(4).Text) & "') "
    
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
      If ReadData(MGrid1.TextMatrix(MGrid1.row, 1), MGrid1.TextMatrix(MGrid1.row, 2), MGrid1.TextMatrix(MGrid1.row, 3)) Then
         SSTab1.Tab = 0
      End If
   End If
End Sub

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
      strSql = "delete from staff_calendar_type where sct01='" & txtField(1) & "' and sct02='" & txtField(2) & "' and sct03='" & txtField(3) & "'"
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
   strCon = strCon & "and sct01='" & txtField(1) & "' "
End If
If Trim(txtField(2)) <> "" Then
   strCon = strCon & "and sct02='" & FdFmt(txtField(2)) & "' "
End If
If Trim(txtField(3)) <> "" Then
   strCon = strCon & "and sct03='" & FdFmt(txtField(3)) & "' "
End If

If Left(strCon, 3) = "and" Then strCon = Mid(strCon, 4, Len(strCon) - 4)

   strExc(1) = " select * from staff_calendar_type where " & strCon
   iR = 1
   Set rsQa = ClsLawReadRstMsg(iR, strExc(1))
   If iR = 1 Then
      RecIsExist = True
      MsgBox "已存在同樣分類的記錄，請先查詢!!", vbCritical
   Else
      RecIsExist = False
   End If
   Set rsQa = Nothing
   
End Function
Private Function FdFmt(ByVal Str01 As String) As String
   FdFmt = Right("00" & Trim("" & Str01), 2)
End Function
