VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140411 
   BorderStyle     =   1  '單線固定
   Caption         =   "名片交換記錄"
   ClientHeight    =   3795
   ClientLeft      =   420
   ClientTop       =   4410
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdRemCont 
      Caption         =   "移除↓"
      Height          =   285
      Left            =   6660
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1980
      Width           =   735
   End
   Begin VB.CommandButton cmdAddCont 
      Caption         =   "新增↑"
      Height          =   285
      Left            =   6660
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7425
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
            Picture         =   "frm140411.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140411.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
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
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   2235
      TabIndex        =   15
      Top             =   2940
      Width           =   1815
      Begin VB.TextBox txtUserNo 
         Height          =   264
         Index           =   0
         Left            =   810
         MaxLength       =   6
         TabIndex        =   5
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除 ->"
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   420
         Width           =   735
      End
      Begin MSForms.Label lblName 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   20
         Top             =   420
         Width           =   900
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1587;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2220
      TabIndex        =   25
      Top             =   1320
      Width           =   5880
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "10372;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   19
      Left            =   345
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
      Width           =   720
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "1270;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   1
      Top             =   984
      Width           =   1125
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1984;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   4
      Left            =   315
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   720
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "1270;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   7
      Left            =   675
      TabIndex        =   6
      Top             =   2625
      Visible         =   0   'False
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   180
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Top             =   1308
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1926;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   660
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1926;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboPlace 
      Height          =   330
      Left            =   1080
      TabIndex        =   21
      Top             =   2610
      Width           =   5550
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9790;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstContact 
      Height          =   600
      Left            =   1080
      TabIndex        =   24
      Top             =   1632
      Width           =   5565
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "9816;1058"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboContact 
      Height          =   330
      Left            =   1080
      TabIndex        =   23
      Top             =   2256
      Width           =   5550
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9790;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstUsers 
      Height          =   630
      Index           =   0
      Left            =   1080
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3060
      Width           =   1125
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "1984;1111"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   2220
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   660
      Width           =   5865
      VariousPropertyBits=   -2147467233
      BackColor       =   16777215
      Size            =   "10345;529"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "接洽同仁："
      Height          =   180
      Index           =   10
      Left            =   165
      TabIndex        =   14
      Top             =   3090
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "往來日期：                           ( 西元 )"
      Height          =   180
      Index           =   13
      Left            =   135
      TabIndex        =   13
      Top             =   1035
      Width           =   2685
   End
   Begin VB.Label Label1 
      Caption         =   "場合："
      Height          =   180
      Index           =   5
      Left            =   135
      TabIndex        =   11
      Top             =   2685
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡人："
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   10
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "記錄編號："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "往來對象："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   9
      Top             =   1320
      Width           =   900
   End
End
Attribute VB_Name = "frm140411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/11 經過確認往來記錄代碼不存在，並且最後資料在2008年；所以先隱藏功能選單
'Memo by Lydia 2022/01/11 改成Form2.0 ; textCUID、lbl1、txtCR(index)、lstContact、cboContact、cboPlace、lstUsers(0)、lblName(0)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2009/2/6
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

Dim m_FieldList() As FIELDITEM

Dim TF_CR As Integer
Dim strTmp As String
Dim oText As Control
Dim idx As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Dim iLanguage As Integer '1:中 2:英 3:日

Private Sub cboContact_GotFocus()
   If cboContact.Locked = False Then
      CloseIme
      'Modified by Lydia 2022/01/11 改成Form 2.0 =>  自動下拉選單
      'SendMessage cboContact.hWnd, CB_SHOWDROPDOWN, 1, 0
      cboContact.DropDown
   End If
End Sub

'新增接洽同仁
Private Sub cmdAdd_Click(Index As Integer)
   AddlstUsers Index
   txtCR(19) = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub
'移除接洽同仁
Private Sub cmdRemove_Click(Index As Integer)
   RemovelstUsers Index
   txtCR(19) = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub

Private Sub cmdAddCont_Click()
   If AddList(lstContact, cboContact, 1) = True Then
      txtCR(4) = ComposeList(lstContact, 1)
      cboContact = ""
   End If
   cboContact.SetFocus
End Sub

Private Sub cmdRemCont_Click()
   If RemoveList(lstContact) = True Then
      txtCR(4) = ComposeList(lstContact, 1)
      cboContact.SetFocus
   End If
End Sub


Private Sub Form_Initialize()
   strExc(0) = "select * from ContactRecord where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_CR = RsTemp.Fields.Count
   ReDim m_FieldList(TF_CR) As FIELDITEM
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   InitialField
   m_EditMode = 0
   ShowRecord -2
   SetInputEntry
   UpdateToolbarState
End Sub
' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 1 '新增
            txtCR(1).Locked = True
            txtCR(2).SetFocus
            
         Case 2 '修改
            txtCR(1).Locked = True
            txtCR(2).SetFocus
         
         Case 4 '查詢
            txtCR(1).Locked = False
            txtCR(1).SetFocus
            
         Case Else
            txtCR(1).Locked = True
            txtCR(1).SetFocus
      End Select
   End If
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
         If m_bUpdate And txtCR(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtCR(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtCR(1) <> "" Then
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

Private Sub Form_Unload(Cancel As Integer)
   Set frm140411 = Nothing
End Sub
' 初始化欄位陣列
Private Sub InitialField()
   For Each oText In txtCR
      idx = oText.Index
      m_FieldList(idx).fiName = "CR" & Format(idx, "00")
   Next
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim adoRst As New ADODB.Recordset
   Dim stCon As String
   'Modified by Lydia 2022/01/12 CR05改成代號B13
   'stCon = " and CR05='交換名片'"
   stCon = " and CR05='B13'"
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM ContactRecord" & _
            " WHERE CR01 = '" & txtCR(1) & "'" & stCon
      Case -2
         strExc(0) = "SELECT * FROM ContactRecord where 1=1" & stCon & " order by CR01 ASC"
      Case -1
         strExc(0) = "SELECT * FROM ContactRecord" & _
            " WHERE CR01 <'" & txtCR(1) & "'" & stCon & " order by CR01 DESC"
      Case 1
         strExc(0) = "SELECT * FROM ContactRecord" & _
            " WHERE CR01 >'" & txtCR(1) & "'" & stCon & " order by CR01 ASC"
      Case 2
         strExc(0) = "SELECT * FROM ContactRecord where 1=1" & stCon & " order by CR01 DESC"
   End Select
      
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = 0 Then
         MsgBox "查無資料！", vbInformation
      ElseIf p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         ClearField
         MsgBox "查無資料！", vbInformation
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtCR(1).SetFocus
      txtCR_GotFocus 1
   End If
End Function
' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   ClearField
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtCR
            idx = oText.Index
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            
            'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
            'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
               m_FieldList(idx).fiType = 0
            'Else
            '   m_FieldList(idx).fiType = 1
            'End If
            'end 2017/06/29
            oText.Text = m_FieldList(idx).fiOldData
         Next
         CUID(1) = "" & .Fields("CR12")
         CUID(2) = "" & .Fields("CR13")
         CUID(3) = "" & .Fields("CR14")
         CUID(4) = "" & .Fields("CR15")
         CUID(5) = "" & .Fields("CR16")
         CUID(6) = "" & .Fields("CR17")
         txtCR_Validate 3, False
         SetCboPlace txtCR(7)
         SetlstUsers 0, txtCR(19)
      End If
   End With
   UpdateCUID CUID, textCUID
   txtCR(1).Tag = txtCR(1)
End Sub

Private Sub ClearField()
   Dim oLabel As LABEL
   For Each oText In txtCR
      oText.Text = Empty
   Next
   lbl1 = Empty
   
   If m_EditMode = 1 Then
      '新增時往來日期預設當天
      txtCR(2) = strSrvDate(1)
   End If
   For intI = 1 To TF_CR
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   cboContact.Clear
   lstContact.Clear
   cboPlace.Clear
   txtUserNo(0) = ""
   lblName(0) = ""
   lstUsers(0).Clear
   'Added by Lydia 2022/01/11
   cboContact.Tag = ""
   lstContact.Tag = ""
   cboPlace.Tag = ""
   lstUsers(0).Tag = ""
End Sub

Private Sub setContact(oCombo As Control, oList As Control, Optional p_stList As String)
   Dim arrID
   Dim stPCC01 As String
   stPCC01 = Left(txtCR(3), 8)
   
   oCombo.Clear
   oCombo.Tag = "" 'Added by Lydia 2022/01/11
   Select Case iLanguage
      Case 1 '中 -> 英 -> 日
         strExc(0) = "select pcc02 c1,nvl(pcc05,nvl(pcc03,pcc04)) c2 from potcustcont where pcc01='" & stPCC01 & "' order by 1 desc"
      
      Case 3 '日 -> 英 -> 中
         strExc(0) = "select pcc02 c1,nvl(pcc04,nvl(pcc03,pcc05)) c2 from potcustcont where pcc01='" & stPCC01 & "' order by 1 desc"
         
      Case Else '英 -> 日 -> 中
         strExc(0) = "select pcc02 c1,nvl(pcc03,nvl(pcc04,pcc05)) c2 from potcustcont where pcc01='" & stPCC01 & "' order by 1 desc"
   End Select
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      '設定聯絡人選單
      .MoveFirst
      Do While Not .EOF
         oCombo.AddItem "" & .Fields(1), 0
         'Modified by Lydia 2022/01/11 改成Form 2.0沒有ItemData屬性
         'oCombo.ItemData(0) = .Fields(0)
         oCombo.Tag = .Fields(0) & "," & oCombo.Tag
         .MoveNext
      Loop
      '設定聯絡人清單
      If p_stList <> "" Then
         oList.Clear
         oList.Tag = "" 'Added by Lydia 2022/01/11
         arrID = Split(p_stList, ",")
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("C1") = arrID(intI) Then
                  oList.AddItem "" & .Fields(1), 0
                  'Modified by Lydia 2022/01/11 改成Form 2.0沒有ItemData屬性
                  'oList.ITEMDATA(0) = .Fields(0)
                  oList.Tag = .Fields(0) & "," & oList.Tag
                  Exit Do
               End If
               .MoveNext
            Loop
         Next
      End If
      End With
   End If
End Sub

Private Function ComposeList(oList As Control, Optional p_iOpt As Integer = 0) As String
'Modified by Lydia 2022/01/11 改成Form 2.0
'   Dim iPos As Integer, stItem As String
'   strExc(1) = ""
'   If oList.ListCount > 0 Then
'      For intI = 0 To oList.ListCount - 1
'         If p_iOpt = 0 Then
'            iPos = InStr(oList.List(intI), Chr(1))
'            If iPos > 0 Then
'               stItem = Left(oList.List(intI), iPos - 1)
'            Else
'               stItem = oList.List(intI)
'            End If
'         Else
'            stItem = Format(oList.ITEMDATA(intI), "00")
'         End If
'         If intI = 0 Then
'            strExc(1) = stItem
'         Else
'            strExc(1) = strExc(1) & "," & stItem
'         End If
'      Next
'   End If
'   ComposeList = strExc(1)
   ComposeList = oList.Tag
'end 2022/01/11
End Function

Private Function GetCustData(p_stCust As String) As Boolean
   Dim aiOrder(1 To 3) As Integer
   Select Case Left(p_stCust, 1)
      Case "X"
         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
      Case "Y"
         strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3 from fagent where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "'"
      Case "R"
         strExc(0) = "select pcu36,pcu08,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) pcu03,pcu07,PCU09 N3 from potcustomer where pcu01='" & Left(p_stCust, 8) & "' and pcu02='" & Right(p_stCust, 1) & "'"
      Case Else
         MsgBox "往來對象必須為 X、Y 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
         Exit Function
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   lbl1 = ""
   If intI = 1 Then
      iLanguage = Val("" & RsTemp(0))
      Select Case iLanguage
         Case 1 '中 -> 英 -> 日
            aiOrder(1) = 1
            aiOrder(2) = 2
            aiOrder(3) = 3
            
         Case 3 '日 -> 中 -> 英
            aiOrder(1) = 3
            aiOrder(2) = 1
            aiOrder(3) = 2
         
         Case Else '英 -> 中 -> 日
            aiOrder(1) = 2
            aiOrder(2) = 1
            aiOrder(3) = 3
      End Select
      For intI = 1 To 3
         If Not IsNull(RsTemp(aiOrder(intI))) Then
            lbl1 = RsTemp(aiOrder(intI))
            Exit For
         End If
      Next
      GetCustData = True
   Else
      MsgBox "往來對象輸入錯誤！"
   End If
End Function

Private Sub txtCR_Change(Index As Integer)
   If Index = 3 Then
      If txtCR(3) <> txtCR(3).Tag Then
         cboContact.Clear
         lstContact.Clear
         'Added by Lydia 2022/01/11
         cboContact.Tag = ""
         lstContact.Tag = ""
         'end 2022/0/1/11
      End If
      txtCR(3).Tag = txtCR(3).Text
   End If
End Sub

Private Sub txtCR_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtCR(Index)
End Sub

'Modified by Lydia 2022/01/11 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
'Private Sub txtCR_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtCR_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCR_Validate(Index As Integer, Cancel As Boolean)
   Dim iLen As Integer
   Select Case Index
      Case 3
         If txtCR(Index) <> "" Then
            If Len(txtCR(Index)) > 5 Then
               txtCR(Index) = Left(txtCR(Index) & "000", 9)
               If GetCustData(txtCR(Index)) = False Then
                  If m_EditMode = "1" Or m_EditMode = "2" Then
                     Cancel = True
                     txtCR_GotFocus Index
                  End If
               Else
                  setContact cboContact, lstContact, txtCR(4)
               End If
            Else
               Cancel = True
               MsgBox "往來對象編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
               txtCR_GotFocus Index
            End If
         End If
         
      Case 2, 10
         If txtCR(Index) <> "" Then
            If CheckIsDate(txtCR(Index)) = False Then
               txtCR_GotFocus Index
               Cancel = True
            End If
         End If
   End Select
   
   If Cancel = False Then
      If txtCR(Index).MaxLength > 0 Then
         Select Case Index
            '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
            Case 6, 7, 8
               iLen = txtCR(Index).MaxLength - 1
            Case Else
               iLen = txtCR(Index).MaxLength
         End Select
         If Not CheckLengthIsOK(txtCR(Index), iLen) Then
            Cancel = True
         End If
      End If
   End If
End Sub

Private Function AddList(oList As Control, oCombo As Control, Optional p_iOpt As Integer = 0) As Boolean
   Dim idx As Integer, bFound As Boolean, stNewItem As String
   Dim stSort As String, iPos As Integer
   Dim iNewItemData As String 'Modified by Lydia 2022/01/11 Integer 改成String
   
   If oCombo.Text = "" Then
      Exit Function
   End If
   
   'Modified by Lydia 2022/01/11 0=>""
   iNewItemData = ""
   If p_iOpt = 1 Then
      If oCombo.ListIndex = -1 Then
         MsgBox "聯絡人資料不存在！"
         Exit Function
      Else
         'Modified by Lydia 2022/01/11 改成Form 2.0
         'iNewItemData = oCombo.ItemData(oCombo.ListIndex)
         iNewItemData = PUB_GetItemData(oCombo.Tag, oCombo.ListIndex)
      End If
   End If
   
   '若有控制字元時後面為說明文字不抓
   iPos = InStr(oCombo, Chr(1))
   If iPos > 0 Then
      stNewItem = Left(oCombo, iPos - 1)
   Else
      stNewItem = oCombo
   End If
   If InStr(stNewItem, ";") > 0 Then
      MsgBox "分號[;]為系統保留字，請改用其他符號！", vbExclamation
      oCombo.SetFocus
      Exit Function
   End If

   If stNewItem <> "" Then
      'Modified by Lydia 2022/01/11 改成Form 2.0
'      For idx = 0 To oList.ListCount - 1
'         If oList.List(idx) = stNewItem And oList.ITEMDATA(idx) = iNewItemData Then
'            MsgBox "資料已存在！"
'            AddList = False
'            bFound = True
'            Exit For
'         End If
'      Next
         If InStr(oList.Tag, iNewItemData) > 0 Then
            MsgBox "資料已存在！"
            AddList = False
            bFound = True
         End If
         'end 2022/01/11
      If bFound = False Then
         oList.AddItem stNewItem, 0
         'Modified by Lydia 2022/01/11 改成Form 2.0
         'If p_iOpt <> 0 Then
         '   oList.ItemData(0) = oCombo.ItemData(oCombo.ListIndex)
         'End If
         oList.Tag = iNewItemData & "," & oList.Tag
         AddList = True
      End If
   End If
End Function

Private Function RemoveList(oList As Control) As Boolean
 'Modified by Lydia 2022/01/11 改成Form 2.0
'   Dim ii As Integer
'   If oList.ListCount > 0 Then
'      ii = 0
'      Do While ii < oList.ListCount
'         If oList.Selected(ii) = True Then
'            RemoveList = True
'            oList.RemoveItem ii
'            ii = ii - 1
'         End If
'         ii = ii + 1
'      Loop
'   End If
   oList.Tag = PUB_RemoveListBox2(oList, oList.Tag)
'end 2022/01/11
End Function
Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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
' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         SetCboPlace
         
      Case vbKeyF3 ' 修改
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         SetCtrlReadOnly True
         ClearField
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
         
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
         
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
         
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
         
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  txtCR(1) = txtCR(1).Tag
                  m_EditMode = 0
                  SetInputEntry
                  ShowRecord
                  UpdateToolbarState
               End If
            Case Else
               txtCR(1) = txtCR(1).Tag
               m_EditMode = 0
               SetInputEntry
               ShowRecord
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtCR(1).SetFocus
               txtCR_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
   
   Dim Cancel As Boolean, ii As Integer, jj As Integer

   For Each oText In txtCR
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         Cancel = False
         txtCR_Validate oText.Index, Cancel
         If Cancel = True Then
            oText.SetFocus
            txtCR_GotFocus oText.Index
            Exit Function
         End If
      End If
   Next
   '查詢
   If m_EditMode = 4 Then
      If txtCR(1) = "" Then
         ShowMsg "請輸入欲查詢之往來記錄編號 !"
         txtCR(1).SetFocus
         txtCR_GotFocus 1
         Exit Function
      End If
   '維護
   Else
      If txtCR(2).Text = "" Then
         ShowMsg "往來日期不可為空白 !"
         txtCR(2).SetFocus
         Exit Function
      End If
     
      If txtCR(3).Text = "" Then
         ShowMsg "往來對象不可為空白 !"
         txtCR(3).SetFocus
         Exit Function
      End If
      
      If cboPlace.Text = "" Then
         ShowMsg "場合不可為空白，若不在選項內請自行輸入 !"
         cboPlace.SetFocus
         Exit Function
      ElseIf GetTextLength(cboPlace) > txtCR(7).MaxLength Then
         ShowMsg "場合長度超過限制(" & txtCR(7).MaxLength & "個字元)!"
         cboPlace.SetFocus
         Exit Function
      End If

      If lstUsers(0).ListCount = 0 Then
         ShowMsg "接洽同仁不可空白!"
         txtUserNo(0).SetFocus
         txtUserNo_GotFocus 0
         Exit Function
      End If
      
   End If
   
    'Added by Lydia 2022/01/12 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
   
   TxtValidate = True
   
End Function

Private Sub UpdateFieldNewData()
   txtCR(7) = cboPlace.Text
   For Each oText In txtCR
      idx = oText.Index
      Select Case idx
         Case 2
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         'Added by Lydia 2022/04/18去掉多餘的,
         Case 4, 19   '聯絡人04,接洽同仁19
            If Right(oText.Text, 1) = "," Then
                m_FieldList(idx).fiNewData = Mid(oText.Text, 1, Len(oText.Text) - 1)
            Else
                m_FieldList(idx).fiNewData = oText.Text
            End If
         'end 2022/04/18
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub
' 新增記錄
Private Function AddRecord() As Boolean
   Dim stSQL As String, stCols As String, stValues As String

On Error GoTo ErrHand
   cnnConnection.BeginTrans

   If txtCR(1) = "" Then
      m_FieldList(1).fiNewData = AutoNo("K", 6)
      
   End If

   '畫面有的欄位才更新
   'Modified by Lydia 2022/01/12 CR05改成代號; 參考往來記錄
   'stCols = ",CR05,CR06": stValues = ",'交換名片','交換名片'"
   stCols = ",CR05,CR06": stValues = ",'B13','交換名片'"
   For Each oText In txtCR
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO ContactRecord (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   
   cnnConnection.Execute stSQL
   
   cnnConnection.CommitTrans
   AddRecord = True
   
   txtCR(1) = m_FieldList(1).fiNewData
   txtCR(1).Tag = txtCR(1)
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "delete from ContactRecord where cr01='" & txtCR(1) & "'"
   Pub_SeekTbLog stSQL
   
   cnnConnection.Execute stSQL
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtCR(1).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE ContactRecord SET "
   stSet = ""
   For Each oText In txtCR
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where cr01='" & txtCR(1) & "'; end; "
      Pub_SeekTbLog stSQL
      
      cnnConnection.Execute stSQL, intI
   End If
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtCR
      oText.Locked = bLocked
   Next
   cboContact.Locked = bLocked
   cmdAddCont.Enabled = Not bLocked
   cmdRemCont.Enabled = Not bLocked
   cboPlace.Locked = bLocked
   Frame2.Visible = Not bLocked
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Control)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
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
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
   End Select
End Sub

Private Sub SetCboPlace(Optional sPlace As String)
   cboPlace.Clear
   cboPlace.Tag = "" 'Added by Lydia 2022/01/11
   cboPlace.AddItem "會議場合", 0
   cboPlace.AddItem "彼所/公司", 0
   cboPlace.AddItem "台一", 0
   If sPlace <> "" Then
      cboPlace.AddItem sPlace, 0
      cboPlace.ListIndex = 0
   End If
End Sub

Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
      'Modified by Lydia 2022/01/10 改成Form 2.0
'      For idx = 0 To lstUsers(p_idx).ListCount - 1
'         If lstUsers(p_idx).ITEMDATA(idx) = PUB_Num2Id(txtUserNo(p_idx)) Then
'            MsgBox "員工已存在於接洽同仁清單中！"
'            txtUserNo(p_idx).SetFocus
'            txtUserNo_GotFocus p_idx
'            bFound = True
'            Exit For
'         End If
'      Next
         If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
            MsgBox "員工已存在於接洽同仁清單中！"
            txtUserNo(p_idx).SetFocus
            txtUserNo_GotFocus p_idx
            bFound = True
         End If
         'end 2022/01/10
      If bFound = False Then
         lstUsers(p_idx).AddItem lblName(p_idx), 0
         'Modified by Lydia 2022/01/10 改成Form 2.0
         'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(txtUserNo(p_idx))
         lstUsers(p_idx).Tag = txtUserNo(p_idx) & "," & lstUsers(p_idx).Tag
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
      End If
   End If
End Sub

Private Function ComposeListX(p_index As Integer) As String
   'Modified by Lydia 2022/01/10 改成Form 2.0
'   strExc(1) = ""
'   If lstUsers(p_index).ListCount > 0 Then
'      strExc(1) = PUB_Num2Id(lstUsers(p_index).ITEMDATA(0))
'      For intI = 1 To lstUsers(p_index).ListCount - 1
'         strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ITEMDATA(intI))
'      Next
'   End If
'   ComposeListX = strExc(1)
   ComposeListX = lstUsers(p_index).Tag
   'end 2022/01/10
End Function

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'Added by Lydia 2022/01/11
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
                  'Modified by Lydia 2022/01/11 改成Form 2.0沒有ItemData屬性
                  'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
                   lstUsers(p_idx).Tag = .Fields(0) & "," & lstUsers(p_idx).Tag
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

Private Sub RemovelstUsers(p_idx As Integer)
   'Modified by Lydia 2022/01/11 改成Form 2.0
'   Dim idx As Integer, ii As Integer
'   If lstUsers(p_idx).ListCount > 0 Then
'      ii = 0
'      For idx = 0 To lstUsers(p_idx).ListCount - 1
'         If lstUsers(p_idx).Selected(ii) = True Then
'            lstUsers(p_idx).RemoveItem ii
'            ii = ii - 1
'         End If
'         ii = ii + 1
'      Next
'   End If
   lstUsers(p_idx).Tag = PUB_RemoveListBox2(lstUsers(p_idx), lstUsers(p_idx).Tag)
   'end 2022/01/11
End Sub

Private Sub txtUserNo_Change(Index As Integer)
   Dim strTempName As String
   If Len(txtUserNo(Index)) = 5 Then
      If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
         lblName(Index) = strTempName
      End If
   Else
      lblName(Index) = ""
   End If
End Sub

Private Sub txtUserNo_GotFocus(Index As Integer)
   TextInverse txtUserNo(Index)
End Sub

'Add By Sindy 2010/11/26
Private Sub txtUserNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_Validate(Index As Integer, Cancel As Boolean)
   Dim strTempName As String
   If txtUserNo(Index).Visible = True Then
      If txtUserNo(Index) <> "" And lblName(Index) = "" Then
         If Len(txtUserNo(Index)) = 5 Then
            If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
               lblName(Index) = strTempName
            End If
         End If
         If lblName(Index) = "" Then
            MsgBox "員工編號輸入錯誤！", vbExclamation
            Cancel = True
         End If
      End If
   End If
End Sub
