VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210102 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件資料修改"
   ClientHeight    =   5604
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7560
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5604
   ScaleWidth      =   7560
   Begin VB.TextBox txtEdit 
      Height          =   510
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   4
      Text            =   "frm210102.frx":0000
      Top             =   2145
      Width           =   5925
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   4
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   3
      Top             =   712
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   3
      Left            =   3015
      MaxLength       =   1
      TabIndex        =   2
      Top             =   712
      Width           =   240
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   2
      Left            =   1980
      MaxLength       =   6
      TabIndex        =   1
      Top             =   712
      Width           =   960
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   1
      Left            =   1425
      MaxLength       =   3
      TabIndex        =   0
      Top             =   712
      Width           =   510
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   4620
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
            Picture         =   "frm210102.frx":001F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":033B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":0657
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":0833
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":0B4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":0E6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":1187
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":14A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":17BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":1ADB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210102.frx":1DF7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
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
            Object.Tag             =   "F2"
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
   Begin MSForms.ComboBox cboCU166 
      Height          =   300
      Left            =   1665
      TabIndex        =   41
      Top             =   3060
      Width           =   5775
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "10186;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCU167 
      Height          =   300
      Left            =   4860
      TabIndex        =   40
      Top             =   2760
      Width           =   2580
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "4551;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboContact 
      Height          =   300
      Left            =   1440
      TabIndex        =   36
      Top             =   2715
      Width           =   1770
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "3122;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "國內副本接洽人："
      Height          =   255
      Index           =   14
      Left            =   3330
      TabIndex        =   39
      Top             =   2760
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "國內副本收件人："
      Height          =   255
      Index           =   13
      Left            =   180
      TabIndex        =   38
      Top             =   3090
      Width           =   1470
   End
   Begin VB.Label lblContact 
      AutoSize        =   -1  'True
      Caption         =   "接洽人："
      Height          =   255
      Left            =   180
      TabIndex        =   37
      Top             =   2745
      Width           =   750
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   14
      Left            =   3780
      TabIndex        =   34
      Top             =   5220
      Width           =   1125
      VariousPropertyBits=   27
      Size            =   "1984;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   13
      Left            =   2610
      TabIndex        =   33
      Top             =   5220
      Width           =   1125
      VariousPropertyBits=   27
      Size            =   "1984;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   12
      Left            =   1440
      TabIndex        =   32
      Top             =   5220
      Width           =   1125
      VariousPropertyBits=   27
      Size            =   "1984;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   11
      Left            =   1410
      TabIndex        =   31
      Top             =   4920
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   10
      Left            =   1410
      TabIndex        =   30
      Top             =   4620
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   9
      Left            =   1410
      TabIndex        =   29
      Top             =   4320
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   8
      Left            =   1410
      TabIndex        =   28
      Top             =   4020
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   7
      Left            =   1410
      TabIndex        =   27
      Top             =   3720
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   6
      Left            =   5010
      TabIndex        =   26
      Top             =   3420
      Width           =   2355
      VariousPropertyBits=   27
      Size            =   "4154;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請案號："
      Height          =   255
      Index           =   7
      Left            =   3750
      TabIndex        =   25
      Top             =   3420
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "客戶案件案號："
      Height          =   255
      Index           =   2
      Left            =   30
      TabIndex        =   24
      Top             =   2160
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　(英)："
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   23
      Top             =   1290
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "分所案號："
      Height          =   255
      Index           =   17
      Left            =   3780
      TabIndex        =   22
      Top             =   720
      Width           =   1200
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   5
      Left            =   1410
      TabIndex        =   21
      Top             =   3420
      Width           =   2115
      VariousPropertyBits=   27
      Size            =   "3731;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   4
      Left            =   1440
      TabIndex        =   20
      Top             =   1890
      Width           =   2115
      VariousPropertyBits=   27
      Size            =   "3731;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   3
      Left            =   1425
      TabIndex        =   19
      Top             =   1590
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update："
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   18
      Top             =   5220
      Width           =   1200
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   2
      Left            =   1440
      TabIndex        =   17
      Top             =   1290
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請人２："
      Height          =   255
      Index           =   12
      Left            =   180
      TabIndex        =   16
      Top             =   4020
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請人１："
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   15
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請日："
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   14
      Top             =   3420
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   255
      Index           =   9
      Left            =   180
      TabIndex        =   13
      Top             =   1890
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　(日)："
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   12
      Top             =   1590
      Width           =   1200
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   990
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "10636;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請人５："
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   10
      Top             =   4920
      Width           =   1200
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   720
      Width           =   2430
      VariousPropertyBits=   27
      Size            =   "4286;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請人４："
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   8
      Top             =   4620
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請人３："
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱(中)："
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   1005
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   1200
   End
End
Attribute VB_Name = "frm210102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/18 改成Form2.0 (lblDisp,cboContact,cboCU167,cboCU166)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Memo by Lydia 2019/07/01 表單名稱:客戶案件資料維護=>案件資料維護
Option Explicit

'授權員工編號
Dim strSalesNo As String
'授權員工所屬部門代碼
Dim strDeptNo As String
'目前狀態
Dim iCurState As Integer
'前一案號
Dim lst_KEY(1 To 4) As String
'目前案號
Dim cur_KEY(1 To 4) As String
'所屬智權人員
Dim stKeySales As String, stDuty As String
Dim strAppNo1 As String '申請人1編號
'Added by Morgan 2020/6/15
Dim strAppNo2 As String '申請人2編號
Dim strAppNo3 As String '申請人3編號
Dim strAppNo4 As String '申請人4編號
Dim strAppNo5 As String '申請人5編號
'end 2020/6/15
'帶人主管權限
'Dim bolPLimit As Boolean 'Add by Sindy 2010/7/28
Dim strContact As String, strCU167 As String 'Added by Morgan 2022/1/18
'Add by Sindy 2023/12/22
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'2023/12/22 END


Public Sub setSalesNo(stNo As String)
   strSalesNo = stNo
End Sub

Public Sub setDeptNo(stDept As String)
   strDeptNo = stDept
End Sub

Private Sub cboCU166_Click()
   If cboCU166.ListIndex >= 0 And cboCU166.Tag <> "" & cboCU166.ListIndex Then
      cboCU167.Clear
      strExc(0) = "select cu127 from customer where " & ChgCustomer(Left(cboCU166, 9))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         PUB_AddContact Left(cboCU166, 8), cboCU167, "" & RsTemp("cu127"), , True, strCU167 'Modified by Morgan 2025/9/4
      End If
      cboCU166.Tag = "" & cboCU166.ListIndex
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Call SetToolBar(0)
   Call FormReset
   '預設為查詢
   Call SetToolBar(4)
   Call SetInputs(4)
   iCurState = 4
   
   Call Pub_AddPersonRec("frm210102") 'Added by Lydia 2019/07/01 智權部-個人常用區
End Sub


'工具列控制
Private Sub SetToolBar(iStatus As Integer)

   Dim i As Integer
   For i = 1 To 13
      tlbar.Buttons(i).Enabled = False
   Next
   tlbar.Buttons(14).Enabled = True
   
   Select Case iStatus
   
      Case 0
      '瀏覽
         tlbar.Buttons(2).Enabled = True
         tlbar.Buttons(4).Enabled = True
      Case 1
      '新增
      Case 2
      '修改
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
      Case 3
      '刪除
      Case 4
      '查詢
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
      Case Else
      
   End Select
   
End Sub

Private Sub FormReset()

   Dim oText As Object, oLabel As Object
   
   For Each oText In txtKey
      oText.Text = ""
      oText.Locked = True
   Next
   
   For Each oText In txtEdit
      oText.Text = ""
      oText.Locked = True
      oText.Tag = "" 'Added by Morgan 2016/12/14
   Next

   For Each oLabel In lblDisp
      oLabel.Caption = ""
   Next
   cboContact.Clear
   'Added by Morgan 2016/12/14
   cboCU166.Clear
   cboCU167.Clear
   'end 2016/12/14
End Sub

Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)

   Dim oText As Object, oLabel As Object
   
   Select Case iStatus
      
      Case 2
      '修改
         For Each oText In txtKey
            oText.Locked = True
            oText.Enabled = False
         Next
         For Each oText In txtEdit
            oText.Locked = False
         Next
         cboContact.Locked = False
         cboCU166.Locked = False 'Added by Morgan 2016/12/14
         cboCU167.Locked = False 'Added by Morgan 2016/12/14
         If txtKey(1).Text <> "LA" Then
            txtEdit(0).Locked = False
         End If
      Case 4
      '查詢
         For Each oText In txtKey
            oText.Text = ""
            oText.Enabled = True
            oText.Locked = False
         Next
         For Each oText In txtEdit
            oText.Text = ""
            oText.Locked = True
            oText.Tag = "" 'Added by Morgan 2016/12/14
         Next
         For Each oLabel In lblDisp
            oLabel.Caption = ""
         Next
         cboContact.Clear
         cboContact.Locked = True
         'Added by Morgan 2016/12/14
         cboCU166.Clear
         cboCU166.Locked = True
         cboCU167.Clear
         cboCU167.Locked = True
         'end 2016/12/14
      Case Else
      '其他
         For Each oText In txtKey
            oText.Enabled = True
            oText.Locked = True
         Next
         For Each oText In txtEdit
            oText.Locked = True
         Next
         cboContact.Locked = True
         cboCU166.Locked = True 'Added by Morgan 2016/12/14
         cboCU167.Locked = True 'Added by Morgan 2016/12/14
   End Select
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210102 = Nothing
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)

   Select Case Button.Index
      Case 1
      '新增
      Case 2
      '修改
         If stKeySales <> strSalesNo And Val(stKeySales) > 63001 And stDuty <> "2" And Pub_StrUserSt03 <> "M51" Then
            MsgBox "非所屬智權人員客戶之案件，不可修改！", vbCritical
            Exit Sub
         Else
            Call SetToolBar(2)
            Call SetInputs(2)
            txtEdit(0).SetFocus
            Call txtEdit_GotFocus(0)
            iCurState = 2
         End If
      Case 3
      '刪除
      Case 4
      '查詢
         lst_KEY(1) = cur_KEY(1)
         lst_KEY(2) = cur_KEY(2)
         lst_KEY(3) = cur_KEY(3)
         lst_KEY(4) = cur_KEY(4)
         Call SetToolBar(4)
         Call SetInputs(4)
         txtKey(1).SetFocus
         Call txtkey_GotFocus(1)
         iCurState = 4
      Case 11
      '確定
         '查詢
         If iCurState = 4 Then
            If txtKey(1) = "" Then
               MsgBox "系統別不可空白！", vbCritical
               Exit Sub
            End If
            txtKey(1) = Trim(txtKey(1).Text)
            If txtKey(2) = "" Then
               MsgBox "案號不可空白！", vbCritical
               Exit Sub
            End If
            txtKey(2) = Left(Me.txtKey(2).Text & "000000000", 6)
            txtKey(3) = Left(Me.txtKey(3).Text & "0", 1)
            txtKey(4) = Left(Me.txtKey(4).Text & "00", 2)
            cur_KEY(1) = txtKey(1)
            cur_KEY(2) = txtKey(2)
            cur_KEY(3) = txtKey(3)
            cur_KEY(4) = txtKey(4)
         '修改
         ElseIf iCurState = 2 Then
            'Added by Morgan 2025/4/24
            If Left(cboCU166, 8) = Left(strAppNo1, 8) Then
               MsgBox "國內副本收件人不可與申請人1相同！", vbCritical
               Exit Sub
            Else
            'end 2025/4/24
               If UpdateData() = True Then
                  'MsgBox "修改成功", vbInformation
               Else
                  Exit Sub
               End If
            End If
         End If
         
         If doQuery(4) = True Then
            Call SetToolBar(0)
            Call SetInputs
            iCurState = 0
         End If
         txtKey(1).SetFocus
         Call txtkey_GotFocus(1)
         
      Case 12
      '取消
         If iCurState = 4 Then
            cur_KEY(1) = lst_KEY(1)
            cur_KEY(2) = lst_KEY(2)
            cur_KEY(3) = lst_KEY(3)
            cur_KEY(4) = lst_KEY(4)
            If cur_KEY(1) = "" Then
               MsgBox "無前次查詢紀錄，不可取消！", vbCritical
               Exit Sub
            Else
               Call doQuery(4)
            End If
         ElseIf iCurState = 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               Exit Sub
            Else
               Call doQuery(4)
            End If
         End If
         Call SetToolBar(0)
         Call SetInputs
         iCurState = 0
      Case 14
      '結束
         Unload Me
   End Select
      
End Sub

Private Function UpdateData() As Boolean

   Dim strSql As String, lngEffRec As Long, strContNo As String
   Dim strCustContNo As String
   Dim strCCNo As String, strCCContNo As String 'Added by Morgan 2016/12/14
   
   'Modify by Morgan 2008/8/11 若個案接洽人與客戶檔預設接洽人相同時不設定
   PUB_GetContact strAppNo1, strCustContNo, True
   
   'Modified by Morgan 2022/1/18
   'strContNo = Format(cboContact.ItemData(cboContact.ListIndex), "00")
   strContNo = Format(PUB_GetItemData(strContact, cboContact.ListIndex), "00")
   'end 2022/1/18
   If strContNo = "00" Or strContNo = strCustContNo Then
      strContNo = ""
   Else
      strContNo = strContNo
   End If
   
   'Added by Morgan 2016/12/14
   'Modified by Morgan 2017/3/10 國內副本收件人改下拉選單
   strCCNo = ""
   strCCContNo = ""
   If cboCU166.ListIndex > 0 Then
      strCCNo = Left(cboCU166.Text, 9)
      If cboCU167.ListIndex > 0 Then
         'Modified by Morgan 2022/1/18
         'strCCContNo = Format(cboCU167.ItemData(cboCU167.ListIndex), "00")
         strCCContNo = Format(PUB_GetItemData(strCU167, cboCU167.ListIndex), "00")
      Else
         strCCContNo = ""
      End If
   End If
   'end 2016/12/14
   
On Error GoTo ErrHand
   'Modify by Morgan 2008/8/4 加案件聯絡人
   Select Case cur_KEY(1)
      '專利
      Case "P", "CFP", "FCP"
         'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time(PA95,PA96,PA97)
         'strSql = "UPDATE PATENT SET PA48='" & txtEdit(0) & "',PA95='" & strSalesNo & "',PA96=to_number(to_char(sysdate,'YYYYMMDD')),PA97=to_number(to_char(sysdate,'HH24MI'))" & _
            ",PA149='" & strContNo & "',PA168='" & strCCNo & "',PA169='" & strCCContNo & "' WHERE PA01='" & cur_KEY(1) & "' AND PA02='" & cur_KEY(2) & "' AND PA03='" & cur_KEY(3) & "' AND PA04='" & cur_KEY(4) & "'"
         strSql = "UPDATE PATENT SET PA48='" & txtEdit(0) & "' " & _
            ",PA149='" & strContNo & "',PA168='" & strCCNo & "',PA169='" & strCCContNo & "' WHERE PA01='" & cur_KEY(1) & "' AND PA02='" & cur_KEY(2) & "' AND PA03='" & cur_KEY(3) & "' AND PA04='" & cur_KEY(4) & "'"
      '商標
      Case "T", "CFT", "FCT", "TF"
         'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time
         'strSql = "UPDATE TRADEMARK SET TM35='" & txtEdit(0) & "',TM62='" & strSalesNo & "',TM63=to_number(to_char(sysdate,'YYYYMMDD')),TM64=to_number(to_char(sysdate,'HH24MI'))" & _
            ",TM123='" & strContNo & "',TM132='" & strCCNo & "',TM133='" & strCCContNo & "' WHERE TM01='" & cur_KEY(1) & "' AND TM02='" & cur_KEY(2) & "' AND TM03='" & cur_KEY(3) & "' AND TM04='" & cur_KEY(4) & "'"
         strSql = "UPDATE TRADEMARK SET TM35='" & txtEdit(0) & "' " & _
            ",TM123='" & strContNo & "',TM132='" & strCCNo & "',TM133='" & strCCContNo & "' WHERE TM01='" & cur_KEY(1) & "' AND TM02='" & cur_KEY(2) & "' AND TM03='" & cur_KEY(3) & "' AND TM04='" & cur_KEY(4) & "'"
      '法務
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "CFL", "FCL", "L", "LIN", "ACS"
         'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time
         'strSql = "UPDATE LAWCASE SET LC17='" & txtEdit(0) & "',LC31='" & strSalesNo & "',LC32=to_number(to_char(sysdate,'YYYYMMDD')),LC33=to_number(to_char(sysdate,'HH24MI'))" & _
            ",LC42='" & strContNo & "' WHERE LC01='" & cur_KEY(1) & "' AND LC02='" & cur_KEY(2) & "' AND LC03='" & cur_KEY(3) & "' AND LC04='" & cur_KEY(4) & "'"
         strSql = "UPDATE LAWCASE SET LC17='" & txtEdit(0) & "' " & _
            ",LC42='" & strContNo & "' WHERE LC01='" & cur_KEY(1) & "' AND LC02='" & cur_KEY(2) & "' AND LC03='" & cur_KEY(3) & "' AND LC04='" & cur_KEY(4) & "'"
      
      'Add by Morgan 2008/8/4
      '顧問
      Case "LA"
         'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time
         'strSql = "UPDATE HIRECASE SET HC16='" & strSalesNo & "',HC17=to_number(to_char(sysdate,'YYYYMMDD')),HC18=to_number(to_char(sysdate,'HH24MI'))" & _
            ",HC23='" & strContNo & "' WHERE HC01='" & cur_KEY(1) & "' AND HC02='" & cur_KEY(2) & "' AND HC03='" & cur_KEY(3) & "' AND HC04='" & cur_KEY(4) & "'"
         strSql = "UPDATE HIRECASE SET HC23='" & strContNo & "' WHERE HC01='" & cur_KEY(1) & "' AND HC02='" & cur_KEY(2) & "' AND HC03='" & cur_KEY(3) & "' AND HC04='" & cur_KEY(4) & "'"
      '服務
      '"CFC","CPS","FG","PS","S","TB","TC","TD","TM","TR","TS","TT"
      Case Else
         'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time
         'strSql = "UPDATE SERVICEPRACTICE SET SP29='" & txtEdit(0) & "',SP55='" & strSalesNo & "',SP56=to_number(to_char(sysdate,'YYYYMMDD')),SP57=to_number(to_char(sysdate,'HH24MI'))" & _
            ",SP78='" & strContNo & "',SP86='" & strCCNo & "',SP87='" & strCCContNo & "' WHERE SP01='" & cur_KEY(1) & "' AND SP02='" & cur_KEY(2) & "' AND SP03='" & cur_KEY(3) & "' AND SP04='" & cur_KEY(4) & "'"
         strSql = "UPDATE SERVICEPRACTICE SET SP29='" & txtEdit(0) & "' " & _
            ",SP78='" & strContNo & "',SP86='" & strCCNo & "',SP87='" & strCCContNo & "' WHERE SP01='" & cur_KEY(1) & "' AND SP02='" & cur_KEY(2) & "' AND SP03='" & cur_KEY(3) & "' AND SP04='" & cur_KEY(4) & "'"
   End Select
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   'Modified by Lydia 2019/04/23 觸發Trigger
   'cnnConnection.Execute strSql, lngEffRec
   cnnConnection.Execute "begin user_data.user_enabled:=1; " & strSql & " ; end; ", lngEffRec
   cnnConnection.CommitTrans
   UpdateData = True
   Exit Function
   
ErrHand:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3
      '修改
         If tlbar.Buttons(2).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(2))
         End If
      Case vbKeyF4
      '查詢
         If tlbar.Buttons(4).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(4))
         End If
      Case vbKeyF9, vbKeyReturn
      '確定
         If tlbar.Buttons(11).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(11))
         End If
      Case vbKeyF10
      '取消
         If tlbar.Buttons(12).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
        If tlbar.Buttons(14).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(14))
         End If
    End Select
End Sub

'Modify by Sindy 2023/9/6 mark
''Add By Sindy 2010/7/28 讀取是否為帶人主管權限
'Private Sub GetPLimit(strCU13 As String)
'   bolPLimit = False
'   strExc(0) = "select count(*) from staff " & _
'                     "where st01='" & strCU13 & "' " & _
'                     "and st04='2' and (st52='" & strUserNum & "' or st53='" & strUserNum & "' or st54='" & strUserNum & "' or st55='" & strUserNum & "') "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If RsTemp.Fields(0) > 0 Then
'         bolPLimit = True
'      End If
'   End If
'End Sub

Private Function doQuery(ByVal iAct As Integer) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   Dim rsA As New ADODB.Recordset   'add by sonia 2024/12/4
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery = False
   
   Select Case iAct
      Case 4
      '查詢
         Select Case cur_KEY(1)
            '專利
            Case "P", "CFP", "FCP"
               strSql = "Select CU12,CU13 From PATENT, Customer WHERE CU01=SUBSTR(PA26,1,8) AND CU02=SUBSTR(PA26,9,1)" & _
                  " AND PA01='" & cur_KEY(1) & "' AND PA02='" & cur_KEY(2) & "' AND PA03='" & cur_KEY(3) & "' AND PA04='" & cur_KEY(4) & "'"
            '商標
            Case "T", "CFT", "FCT", "TF"
               strSql = "Select CU12,CU13 From TRADEMARK, Customer WHERE CU01=SUBSTR(TM23,1,8) AND CU02=SUBSTR(TM23,9,1)" & _
                  " AND TM01='" & cur_KEY(1) & "' AND TM02='" & cur_KEY(2) & "' AND TM03='" & cur_KEY(3) & "' AND TM04='" & cur_KEY(4) & "'"
            '法務
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"
               strSql = "Select CU12,CU13 From LAWCASE, Customer WHERE CU01=SUBSTR(LC11,1,8) AND CU02=SUBSTR(LC11,9,1)" & _
                  " AND LC01='" & cur_KEY(1) & "' AND LC02='" & cur_KEY(2) & "' AND LC03='" & cur_KEY(3) & "' AND LC04='" & cur_KEY(4) & "'"
            '顧問
            Case "LA"
               strSql = "Select CU12,CU13 From HIRECASE, Customer WHERE CU01=SUBSTR(HC05,1,8) AND CU02=SUBSTR(HC05,9,1)" & _
                  " AND HC01='" & cur_KEY(1) & "' AND HC02='" & cur_KEY(2) & "' AND HC03='" & cur_KEY(3) & "' AND HC04='" & cur_KEY(4) & "'"
            '服務
            '"CFC","CPS","FG","PS","S","TB","TC","TD","TM","TR","TS","TT"
            Case Else
               strSql = "Select CU12,CU13 From SERVICEPRACTICE, Customer WHERE CU01=SUBSTR(SP08,1,8) AND CU02=SUBSTR(SP08,9,1)" & _
                  " AND SP01='" & cur_KEY(1) & "' AND SP02='" & cur_KEY(2) & "' AND SP03='" & cur_KEY(3) & "' AND SP04='" & cur_KEY(4) & "'"
         End Select
         stMessage = "無此記錄之資料！"
        
   End Select
   
On Error GoTo ErrHand

   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
'      Dim strArea As String
'      strArea = "" & rsQuery.Fields("CU12").Value
      'Modify By Sindy 2023/9/6
      'Modify By Sindy 2023/12/22 +, bolSpecMan, strSpecCode
      Call PUB_SetFormSaleDept(strSalesNo, , , , , bolSpecMan, strSpecCode, , , , , , True)
      'modify by sonia 2024/12/4 改不在PUB_ChkSalePerLimit彈訊息
      'If PUB_ChkSalePerLimit(ShowCurrCP13(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4), ""), strSalesNo, , bolSpecMan, strSpecCode, , , "案件") = True Then
      If PUB_ChkSalePerLimit(ShowCurrCP13(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4), ""), strSalesNo, False, bolSpecMan, strSpecCode, , , "案件") = True Then
      'Call GetPLimit("" & rsQuery.Fields("CU13").Value) 'Add By Sindy 2010/7/28
      'Modify By Sindy 2010/7/28 開放帶人主管不限制
      'If strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Then
      'If bolPLimit = True Or strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Then
      '2023/9/6 END
         lst_KEY(1) = cur_KEY(1)
         lst_KEY(2) = cur_KEY(2)
         lst_KEY(3) = cur_KEY(3)
         lst_KEY(4) = cur_KEY(4)
         If ReQuery() = True Then doQuery = True
'      Else
'         MsgBox "業務區別不同不可查詢！", vbCritical
      'add by sonia 2024/12/4 法律所案件則客戶檔之智權人員(介紹人)也可以查詢但不能修改
      ElseIf InStr(cur_KEY(1), "L") > 0 Then
         strSql = "Select CU13, CU12, ST04, A0908 From Lawcase, Customer, Staff, acc090 Where substr(LC11,1,8)=CU01 And substr(LC11,9,1)=CU02 And CU13=ST01 and st15=a0901 And LC01='" & cur_KEY(1) & "' And LC02='" & cur_KEY(2) & "' And LC03='" & cur_KEY(3) & "' And LC04='" & cur_KEY(4) & "' "
         strSql = strSql & " union Select CU13, CU12, ST04, A0908 From Hirecase, Customer, Staff, acc090 Where substr(HC05,1,8)=CU01 And substr(HC05,9,1)=CU02 And CU13=ST01 and st15=a0901 And HC01='" & cur_KEY(1) & "' And HC02='" & cur_KEY(2) & "' And HC03='" & cur_KEY(3) & "' And HC04='" & cur_KEY(4) & "' "
         rsA.CursorLocation = adUseClient
         rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If "" & rsA("CU13").Value = strUserNum Then
               doQuery = True
            ElseIf "" & rsA("A0908").Value = strUserNum Then
               doQuery = True
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         If doQuery = True Then
            lst_KEY(1) = cur_KEY(1)
            lst_KEY(2) = cur_KEY(2)
            lst_KEY(3) = cur_KEY(3)
            lst_KEY(4) = cur_KEY(4)
            If ReQuery() = True Then doQuery = True
         Else
            MsgBox "您無查詢權限！", vbExclamation, "操作錯誤！"
         End If
      Else
         MsgBox "您無查詢權限！", vbExclamation, "操作錯誤！"
      'end 2024/12/4
      End If
   Else
      MsgBox stMessage, vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
   
End Function

'完整資料查詢
Private Function ReQuery() As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, intI As Integer
   Dim bolCCNo As Boolean 'Added by Morgan 2016/12/14
   
On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   
   ReQuery = False
   
   'Added by Morgan 2016/12/14
   cboCU166.Enabled = True
   cboCU166.BackColor = txtEdit(0).BackColor
   cboCU167.BackColor = txtEdit(0).BackColor
   'end 2016/12/14
         
   'Modified by Morgan 2016/12/14 +E01,R05
   Select Case cur_KEY(1)
      '專利
      Case "P", "CFP", "FCP"
         'Modified by Lydia 2017/03/03 FLOOR(PA97/100)||':'||MOD(PA97,100) => SQLTIME(PA97||'00')
         strSql = "select PA01 K01,PA02 K02,PA03 K03,PA04 K04" & _
            ", PA47 D00, PA05 D01, PA06 D02, PA07 D03, NA03 D04" & _
            ", DECODE(PA10,NULL,NULL,(SUBSTRB(PA10,1,4)-1911)||'/'||SUBSTR(PA10,5,2)||'/'||SUBSTR(PA10,7,2)) D05" & _
            ", PA11 D06, C1.CU04 D07, C2.CU04 D08, C3.CU04 D09, C4.CU04 D10, C5.CU04 D11, S1.ST02 D12" & _
            ", DECODE(PA96,NULL,NULL,(SUBSTRB(PA96,1,4)-1911)||'/'||SUBSTR(PA96,5,2)||'/'||SUBSTR(PA96,7,2)) D13" & _
            ", SQLTIME(PA97||'00') D14" & _
            ", PA48 E00,PA168 E01,PA169 E02, C1.CU13 R01,S2.ST04 R02,PA149 R03,C1.CU01 R04" & _
            ", C2.CU01 R05,C3.CU01 R06,C4.CU01 R07,C5.CU01 R08" & _
            " from PATENT, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE PA01='" & cur_KEY(1) & "' AND PA02='" & cur_KEY(2) & "' AND PA03='" & cur_KEY(3) & "' AND PA04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(PA26,1,8) AND C1.CU02(+)=SUBSTR(PA26,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(PA27,1,8) AND C2.CU02(+)=SUBSTR(PA27,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(PA28,1,8) AND C3.CU02(+)=SUBSTR(PA28,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(PA29,1,8) AND C4.CU02(+)=SUBSTR(PA29,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(PA30,1,8) AND C5.CU02(+)=SUBSTR(PA30,9,1)" & _
            " AND NA01(+) = PA09 AND S1.ST01(+)=PA95 AND S2.ST01(+)=C1.CU13"
      '商標
      Case "T", "CFT", "FCT", "TF"
         'Modified by Lydia 2017/03/03 FLOOR(TM64/100)||':'||MOD(TM64,100) => SQLTIME(TM64||'00')
         strSql = "select TM01 K01,TM02 K02,TM03 K03,TM04 K04" & _
            ",TM34 D00, TM05 D01, TM06 D02, TM07 D03, NA03 D04" & _
            ", DECODE(TM11,NULL,NULL,(SUBSTRB(TM11,1,4)-1911)||'/'||SUBSTR(TM11,5,2)||'/'||SUBSTR(TM11,7,2)) D05" & _
            ", TM12 D06, C1.CU04 D07, C2.CU04 D08, C3.CU04 D09, C4.CU04 D10, C5.CU04 D11, S1.ST02 D12" & _
            ", DECODE(TM63,NULL,NULL,(SUBSTRB(TM63,1,4)-1911)||'/'||SUBSTR(TM63,5,2)||'/'||SUBSTR(TM63,7,2)) D13" & _
            ", SQLTIME(TM64||'00') D14" & _
            ", TM35 E00,TM132 E01,TM133 E02, C1.CU13 R01,S2.ST04 R02,TM123 R03,C1.CU01 R04" & _
            ", C2.CU01 R05,C3.CU01 R06,C4.CU01 R07,C5.CU01 R08" & _
            " from TRADEMARK, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE TM01='" & cur_KEY(1) & "' AND TM02='" & cur_KEY(2) & "' AND TM03='" & cur_KEY(3) & "' AND TM04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(TM23,1,8) AND C1.CU02(+)=SUBSTR(TM23,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(TM78,1,8) AND C2.CU02(+)=SUBSTR(TM78,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(TM79,1,8) AND C3.CU02(+)=SUBSTR(TM79,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(TM80,1,8) AND C4.CU02(+)=SUBSTR(TM80,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(TM81,1,8) AND C5.CU02(+)=SUBSTR(TM81,9,1)" & _
            " AND NA01(+) = TM10 AND S1.ST01(+)=TM62 AND S2.ST01(+)=C1.CU13"
      '法務
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "CFL", "FCL", "L", "LIN", "ACS"
         'Added by Morgan 2016/12/14
         cboCU166.Enabled = False
         cboCU166.BackColor = &HE0E0E0
         cboCU167.BackColor = cboCU166.BackColor
         'end 2016/12/14
         'Modified by Lydia 2017/03/03 FLOOR(LC33/100)||':'||MOD(LC33,100) => SQLTIME(LC33||'00')
         strSql = "select LC01 K01,LC02 K02,LC03 K03,LC04 K04" & _
            ", LC16 D00, LC05 D01, LC06 D02, LC07 D03, NA03 D04" & _
            ", '' D05, '' D06, C1.CU04 D07, C2.CU04 D08, C3.CU04 D09, C4.CU04 D10, C5.CU04 D11, S1.ST02 D12" & _
            ", DECODE(LC32,NULL,NULL,(SUBSTRB(LC32,1,4)-1911)||'/'||SUBSTR(LC32,5,2)||'/'||SUBSTR(LC32,7,2)) D13" & _
            ", SQLTIME(LC33||'00') D14" & _
            ", LC17 E00,'' E01,'' E02, C1.CU13 R01,S2.ST04 R02,LC42 R03,C1.CU01 R04" & _
            ", C2.CU01 R05,C3.CU01 R06,C4.CU01 R07,C5.CU01 R08" & _
            " from LAWCASE, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE LC01='" & cur_KEY(1) & "' AND LC02='" & cur_KEY(2) & "' AND LC03='" & cur_KEY(3) & "' AND LC04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(LC11,1,8) AND C1.CU02(+)=SUBSTR(LC11,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(LC43,1,8) AND C2.CU02(+)=SUBSTR(LC43,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(LC44,1,8) AND C3.CU02(+)=SUBSTR(LC44,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(LC45,1,8) AND C4.CU02(+)=SUBSTR(LC45,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(LC46,1,8) AND C5.CU02(+)=SUBSTR(LC46,9,1)" & _
            " AND NA01(+) = LC15 AND S1.ST01(+)=LC31 AND S2.ST01(+)=C1.CU13"
      '顧問
      Case "LA"
         'Added by Morgan 2016/12/14
         cboCU166.Enabled = False
         cboCU166.BackColor = &HE0E0E0
         cboCU167.BackColor = cboCU166.BackColor
         'end 2016/12/14
         'Modified by Lydia 2017/03/03 FLOOR(HC18/100)||':'||MOD(HC18,100) => SQLTIME(HC18||'00')
         strSql = "select HC01 K01,HC02 K02,HC03 K03,HC04 K04" & _
            ", HC07 D00, HC06 D01, '' D02, '' D03, '' D04" & _
            ", '' D05, '' D06, C1.CU04 D07, C2.CU04 D08, C3.CU04 D09, C4.CU04 D10, C5.CU04 D11, S1.ST02 D12" & _
            ", DECODE(HC17,NULL,NULL,(SUBSTRB(HC17,1,4)-1911)||'/'||SUBSTR(HC17,5,2)||'/'||SUBSTR(HC17,7,2)) D13" & _
            ", SQLTIME(HC18||'00') D14" & _
            ", '' E00,'' E01,'' E02, C1.CU13 R01,S2.ST04 R02,HC23 R03,C1.CU01 R04" & _
            ", C2.CU01 R05,C3.CU01 R06,C4.CU01 R07,C5.CU01 R08" & _
            " from HIRECASE, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE HC01='" & cur_KEY(1) & "' AND HC02='" & cur_KEY(2) & "' AND HC03='" & cur_KEY(3) & "' AND HC04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(HC05,1,8) AND C1.CU02(+)=SUBSTR(HC05,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(HC24,1,8) AND C2.CU02(+)=SUBSTR(HC24,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(HC25,1,8) AND C3.CU02(+)=SUBSTR(HC25,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(HC26,1,8) AND C4.CU02(+)=SUBSTR(HC26,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(HC27,1,8) AND C5.CU02(+)=SUBSTR(HC27,9,1)" & _
            " AND S1.ST01(+)=HC16 AND S2.ST01(+)=C1.CU13"
      '服務
      '"CFC","CPS","FG","PS","S","TB","TC","TD","TM","TR","TS","TT"
      Case Else
         'Modified by Lydia 2017/03/03 FLOOR(SP57/100)||':'||MOD(SP57,100) => SQLTIME(SP57||'00')
         strSql = "select SP01 K01,SP02 K02,SP03 K03,SP04 K04" & _
            ", SP28 D00, SP05 D01, SP06 D02, SP07 D03, NA03 D04" & _
            ", DECODE(SP10,NULL,NULL,(SUBSTRB(SP10,1,4)-1911)||'/'||SUBSTR(SP10,5,2)||'/'||SUBSTR(SP10,7,2)) D05" & _
            ", SP11 D06, C1.CU04 D07, C2.CU04 D08, C3.CU04 D09, C4.CU04 D10, C5.CU04 D11, S1.ST02 D12" & _
            ", DECODE(SP56,NULL,NULL,(SUBSTRB(SP56,1,4)-1911)||'/'||SUBSTR(SP56,5,2)||'/'||SUBSTR(SP56,7,2)) D13" & _
            ", SQLTIME(SP57||'00') D14" & _
            ", SP29 E00, SP86 E01, SP87 E02, C1.CU13 R01,S2.ST04 R02,SP78 R03,C1.CU01 R04" & _
            ", C2.CU01 R05,C3.CU01 R06,C4.CU01 R07,C5.CU01 R08" & _
            " from SERVICEPRACTICE, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE SP01='" & cur_KEY(1) & "' AND SP02='" & cur_KEY(2) & "' AND SP03='" & cur_KEY(3) & "' AND SP04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(SP08,1,8) AND C1.CU02(+)=SUBSTR(SP08,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(SP58,1,8) AND C2.CU02(+)=SUBSTR(SP58,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(SP59,1,8) AND C3.CU02(+)=SUBSTR(SP59,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(SP65,1,8) AND C4.CU02(+)=SUBSTR(SP65,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(SP66,1,8) AND C5.CU02(+)=SUBSTR(SP66,9,1)" & _
            " AND NA01(+) = SP09 AND S1.ST01(+)=SP55 AND S2.ST01(+)=C1.CU13"
   End Select

   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      txtKey(1) = "" & rsQuery.Fields("K01")
      txtKey(2) = "" & rsQuery.Fields("K02")
      txtKey(3) = "" & rsQuery.Fields("K03")
      txtKey(4) = "" & rsQuery.Fields("K04")
      stKeySales = "" & rsQuery.Fields("R01").Value
      stDuty = "" & rsQuery.Fields("R02").Value
      For intI = 0 To 0
         txtEdit(intI) = "" & rsQuery.Fields("E" & Format(intI, "00"))
      Next intI
      'Modified by Lydia 2017/03/03 去掉UpdateTime
      'For intI = 0 To 14
      For intI = 0 To 13
         lblDisp(intI) = "" & rsQuery.Fields("D" & Format(intI, "00"))
      Next intI
      lblDisp(14) = "" & rsQuery.Fields("D14")  'Added by Lydia 2017/03/03
      strAppNo1 = "" & rsQuery.Fields("R04")
      'Added by Morgan 2020/6/15
      strAppNo2 = "" & rsQuery.Fields("R05")
      strAppNo3 = "" & rsQuery.Fields("R06")
      strAppNo4 = "" & rsQuery.Fields("R07")
      strAppNo5 = "" & rsQuery.Fields("R08")
      'end 2020/6/15
      
      PUB_AddContact strAppNo1, cboContact, "" & rsQuery.Fields("R03"), True, True, strContact
      
      'Added by Morgan 2016/12/14
      'Modified by Morgan 2017/3/10 國內副本收件人改下拉選單
      cboCU166.Clear
      cboCU166.AddItem "", 0
      '與文雄確認先從嚴控管以避免誤設而錯寄,遇特殊情形再例外設定
      '1.不可為相同客戶 2.只能是客戶的關係企業 3.必須是相同的智權人員
      'Modified by Morgan 2020/6/15 +其他申請人--Ex:P124691
      strExc(0) = "select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where substr(cu01,1,6)='" & Left(strAppNo1, 6) & "' and cu02='0' and cu01<>'" & Left(strAppNo1, 8) & "' and cu13='" & stKeySales & "'"
      If strAppNo2 <> "" Then
         strExc(0) = strExc(0) & " union select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where substr(cu01,1,6)='" & Left(strAppNo2, 6) & "' and cu02='0' and cu13='" & stKeySales & "'"
      End If
      If strAppNo3 <> "" Then
         strExc(0) = strExc(0) & " union select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where substr(cu01,1,6)='" & Left(strAppNo3, 6) & "' and cu02='0' and cu13='" & stKeySales & "'"
      End If
      If strAppNo4 <> "" Then
         strExc(0) = strExc(0) & " union select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where substr(cu01,1,6)='" & Left(strAppNo4, 6) & "' and cu02='0' and cu13='" & stKeySales & "'"
      End If
      If strAppNo5 <> "" Then
         strExc(0) = strExc(0) & " union select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where substr(cu01,1,6)='" & Left(strAppNo5, 6) & "' and cu02='0' and cu13='" & stKeySales & "'"
      End If
      strExc(0) = strExc(0) & " order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         intI = 0
         Do While Not RsTemp.EOF
            cboCU166.AddItem RsTemp("CNo") & " " & RsTemp("CName")
            If RsTemp("CNo") = "" & rsQuery.Fields("E01") Then
               intI = cboCU166.ListCount - 1
            End If
            RsTemp.MoveNext
         Loop
         If intI > 0 Then
            cboCU166.Tag = "" & intI
            cboCU166.ListIndex = intI
         
         'Removed by Morgan 2025/4/24 移到下面(若案件讓與給副本收件人時會帶不出來 Ex:P-124432)
         'ElseIf Not IsNull(rsQuery.Fields("E01")) Then
         '   strExc(0) = "select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where cu01||cu02='" & rsQuery.Fields("E01") & "'"
         '   intI = 1
         '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '   If intI = 1 Then
         '      cboCU166.AddItem RsTemp("CNo") & " " & RsTemp("CName"), 1
         '      cboCU166.Tag = "1"
         '      cboCU166.ListIndex = 1
         '   End If
         'end 2025/4/24
         End If
      End If
      
      'Added by Morgan 2025/4/24
      If cboCU166 = "" Then
         If Not IsNull(rsQuery.Fields("E01")) Then
            strExc(0) = "select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where cu01||cu02='" & rsQuery.Fields("E01") & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               cboCU166.AddItem RsTemp("CNo") & " " & RsTemp("CName"), 1
               cboCU166.Tag = "1"
               cboCU166.ListIndex = 1
            End If
         End If
      End If
      'end 2025/4/24
      
      cboCU167.Clear
      If cboCU166.ListIndex >= 0 Then
         PUB_AddContact Left(cboCU166, 8), cboCU167, "" & rsQuery.Fields("E02"), , True, strCU167
      End If
      'end 2016/12/14
      
      ReQuery = True
   Else
      MsgBox "案件〔" & cur_KEY(1) & cur_KEY(2) & cur_KEY(3) & cur_KEY(4) & "〕已被刪除！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Screen.MousePointer = vbDefault
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
   Screen.MousePointer = vbDefault
   
End Function

Private Sub txtEdit_GotFocus(Index As Integer)
   TextInverse txtEdit(Index)
   If txtEdit(Index).Locked = False Then
      Select Case Index
         Case 0, 1
         '客戶案件案號
            'edit by nickc 2007/06/06 切換輸入法改用API
            'txtEdit(Index).IMEMode = 2
            CloseIme
      End Select
   End If
End Sub
'Added by Morgan 2016/12/14
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
      '客戶案件案號
         'edit by nickc 2008/01/22 放不夠寬
         'If CheckLengthIsOK(txtEdit(0), 16) = False Then
         'Modified by Lydia 2017/03/03 改成100
         'If CheckLengthIsOK(txtEdit(0), 30) = False Then
         'Modified by Lydia 2017/06/14 改常數
         'If CheckLengthIsOK(txtEdit(0), 100) = False Then
         If CheckLengthIsOK(txtEdit(0), 專利客戶案號max) = False Then
            txtEdit_GotFocus 0
            Cancel = True
         End If
'Removed by Morgan 2025/4/24 已改為下拉選單
'      'Added by Morgan 2016/12/14
'      '國內副本收件人
'      Case 1
'         If txtEdit(Index).Tag = txtEdit(Index) Then Exit Sub
'         txtEdit(Index).Tag = ""
'         lblDisp(15) = ""
'         cboCU167.Clear
'         If txtEdit(Index) = "" Then Exit Sub
'         If Len(txtEdit(Index)) < 6 Then
'            MsgBox "國內副本收件人輸入錯誤！", vbCritical
'            Cancel = True
'            Call txtEdit_GotFocus(Index)
'            Exit Sub
'         Else
'            txtEdit(Index) = Left(txtEdit(Index) & "000", 9)
'            strExc(0) = "select nvl(cu04, nvl(cu05, cu06)),cu127,cu13 from customer where " & ChgCustomer(txtEdit(Index).Text)
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               lblDisp(15) = "" & RsTemp(0)
'               '與文雄確認先從嚴控管以避免誤設而錯寄,遇特殊情形再例外設定
'               If Left(txtEdit(Index), 8) = Left(strAppNo1, 8) Then
'                  MsgBox "國內副本收件人不可與申請人1相同！", vbCritical
'                  Cancel = True
'                  Call txtEdit_GotFocus(Index)
'                  Exit Sub
'               ElseIf Left(txtEdit(Index), 6) <> Left(strAppNo1, 6) Then
'                  MsgBox "國內副本收件人只能是申請人1的關係企業！", vbCritical
'                  Cancel = True
'                  Call txtEdit_GotFocus(Index)
'                  Exit Sub
'               ElseIf RsTemp("cu13") <> stKeySales Then
'                  MsgBox "國內副本收件人必須和申請人1是相同的智權人員！", vbCritical
'                  Cancel = True
'                  Call txtEdit_GotFocus(Index)
'                  Exit Sub
'               Else
'                  PUB_AddContact Left(txtEdit(Index).Text, 8), cboCU167, "" & RsTemp("cu127"), , True, strCU167
'               End If
'            Else
'               MsgBox "國內副本收件人輸入錯誤！", vbCritical
'               Cancel = True
'               Call txtEdit_GotFocus(Index)
'               Exit Sub
'            End If
'         End If
'         txtEdit(Index).Tag = txtEdit(Index)
'end 2025/4/24
   End Select
End Sub

Private Sub txtkey_GotFocus(Index As Integer)
   TextInverse txtKey(Index)
   If txtKey(Index).Enabled = True Then
      'edit by nickc 2007/06/06 切換輸入法改用API
      'txtKey(Index).IMEMode = 2
      CloseIme
   End If
End Sub

Private Sub txtKey_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
