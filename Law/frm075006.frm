VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075006 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問案件資料維護"
   ClientHeight    =   5820
   ClientLeft      =   612
   ClientTop       =   816
   ClientWidth     =   8964
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8964
   Begin VB.CommandButton cmdIns 
      Caption         =   "各項指示(&N)"
      Height          =   350
      Left            =   6450
      TabIndex        =   16
      Top             =   660
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtCustomer 
      Height          =   300
      Index           =   4
      Left            =   1152
      MaxLength       =   9
      TabIndex        =   8
      Top             =   1735
      Width           =   1335
   End
   Begin VB.TextBox txtCustomer 
      Height          =   300
      Index           =   3
      Left            =   5565
      MaxLength       =   9
      TabIndex        =   7
      Top             =   1418
      Width           =   1335
   End
   Begin VB.TextBox txtCustomer 
      Height          =   300
      Index           =   2
      Left            =   1152
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1418
      Width           =   1335
   End
   Begin VB.TextBox txtCustomer 
      Height          =   300
      Index           =   1
      Left            =   5565
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1101
      Width           =   1335
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "相關卷號(&F)"
      Height          =   350
      Left            =   7680
      TabIndex        =   17
      Top             =   660
      Width           =   1200
   End
   Begin VB.TextBox txtCustomer 
      Height          =   300
      Index           =   0
      Left            =   1152
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1101
      Width           =   1335
   End
   Begin VB.TextBox txtCaseNum 
      Height          =   300
      Left            =   1152
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3003
      Width           =   2925
   End
   Begin VB.TextBox txtCloseYN 
      Height          =   300
      Left            =   1152
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3320
      Width           =   375
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Left            =   5130
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3320
      Width           =   1095
   End
   Begin VB.TextBox txtClsResult 
      Height          =   300
      Left            =   1152
      MaxLength       =   2
      TabIndex        =   14
      Top             =   3637
      Width           =   375
   End
   Begin VB.TextBox txtcp01 
      Height          =   300
      Left            =   1152
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "LA"
      Top             =   784
      Width           =   495
   End
   Begin VB.TextBox txtcp02 
      Height          =   300
      Left            =   1752
      MaxLength       =   6
      TabIndex        =   1
      Top             =   784
      Width           =   1095
   End
   Begin VB.TextBox txtcp03 
      Height          =   300
      Left            =   2952
      MaxLength       =   1
      TabIndex        =   2
      Top             =   784
      Width           =   375
   End
   Begin VB.TextBox txtcp04 
      Height          =   300
      Left            =   3432
      MaxLength       =   2
      TabIndex        =   3
      Top             =   784
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7770
      Top             =   30
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
            Picture         =   "frm075006.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075006.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8964
      _ExtentX        =   15812
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
   Begin MSForms.TextBox txtMemo 
      Height          =   525
      Left            =   1152
      TabIndex        =   15
      Top             =   4590
      Width           =   7590
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13388;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblhc21 
      Height          =   255
      Left            =   6480
      TabIndex        =   58
      Top             =   3977
      Width           =   1590
      VariousPropertyBits=   27
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label IDname 
      Height          =   255
      Left            =   1215
      TabIndex        =   57
      Top             =   5190
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "IDname"
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label UIDname 
      Height          =   255
      Left            =   1215
      TabIndex        =   56
      Top             =   5460
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "UIDname"
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseName 
      Height          =   300
      Left            =   1152
      TabIndex        =   10
      Top             =   2369
      Width           =   6495
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "11456;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   55
      Top             =   1758
      Width           =   2025
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3572;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   54
      Top             =   1441
      Width           =   2025
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3572;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   53
      Top             =   1441
      Width           =   2025
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3572;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   52
      Top             =   1124
      Width           =   2025
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3572;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   51
      Top             =   1124
      Width           =   2025
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3572;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboContact 
      Height          =   300
      Left            =   1152
      TabIndex        =   9
      Top             =   2052
      Width           =   1770
      VariousPropertyBits=   679495711
      DisplayStyle    =   3
      Size            =   "3122;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label19 
      Caption         =   "當 事 人 5："
      Height          =   255
      Left            =   180
      TabIndex        =   50
      Top             =   1758
      Width           =   945
   End
   Begin VB.Label Label17 
      Caption         =   "當 事 人 4："
      Height          =   255
      Left            =   4590
      TabIndex        =   49
      Top             =   1441
      Width           =   945
   End
   Begin VB.Label Label13 
      Caption         =   "當 事 人 3："
      Height          =   255
      Left            =   180
      TabIndex        =   48
      Top             =   1441
      Width           =   945
   End
   Begin VB.Label Label9 
      Caption         =   "當 事 人 2："
      Height          =   255
      Left            =   4590
      TabIndex        =   47
      Top             =   1109
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "接洽人："
      Height          =   255
      Index           =   160
      Left            =   180
      TabIndex        =   46
      Top             =   2075
      Width           =   720
   End
   Begin VB.Label lblhc22 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   45
      Top             =   4294
      Width           =   1155
   End
   Begin VB.Label lblhc20 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   3750
      TabIndex        =   44
      Top             =   3977
      Width           =   1365
   End
   Begin VB.Label lblhc19 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1290
      TabIndex        =   43
      Top             =   3977
      Width           =   1215
   End
   Begin VB.Label Label78 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷備註："
      Height          =   255
      Left            =   180
      TabIndex        =   42
      Top             =   4294
      Width           =   1260
   End
   Begin VB.Label Label79 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷員："
      Height          =   255
      Left            =   5340
      TabIndex        =   41
      Top             =   3977
      Width           =   1080
   End
   Begin VB.Label Label80 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷日："
      Height          =   255
      Left            =   2640
      TabIndex        =   40
      Top             =   3977
      Width           =   1080
   End
   Begin VB.Label Label81 
      AutoSize        =   -1  'True
      Caption         =   "北所銷卷日："
      Height          =   255
      Left            =   180
      TabIndex        =   39
      Top             =   3977
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "當 事 人 1："
      Height          =   255
      Left            =   180
      TabIndex        =   38
      Top             =   1124
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   180
      TabIndex        =   37
      Top             =   807
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "首次聘任日："
      Height          =   255
      Left            =   180
      TabIndex        =   36
      Top             =   2709
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "本次聘任期間：                     "
      Height          =   255
      Left            =   4155
      TabIndex        =   35
      Top             =   2709
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "案件備註："
      Height          =   255
      Left            =   180
      TabIndex        =   34
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "分所案號："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   33
      Top             =   3026
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "是否閉卷："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   32
      Top             =   3343
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   180
      TabIndex        =   31
      Top             =   2392
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "(Y:是)"
      Height          =   255
      Left            =   1575
      TabIndex        =   30
      Top             =   3343
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "閉卷日期："
      Height          =   255
      Left            =   4155
      TabIndex        =   29
      Top             =   3343
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "閉卷原因："
      Height          =   255
      Left            =   180
      TabIndex        =   28
      Top             =   3660
      Width           =   975
   End
   Begin VB.Label lbeResult 
      Height          =   255
      Left            =   1575
      TabIndex        =   27
      Top             =   3660
      Width           =   4095
   End
   Begin VB.Label lbeFirstHire 
      Height          =   255
      Left            =   1320
      TabIndex        =   26
      Top             =   2709
      Width           =   1335
   End
   Begin VB.Label lbeHire 
      Height          =   255
      Left            =   5520
      TabIndex        =   25
      Top             =   2709
      Width           =   2055
   End
   Begin VB.Label Label24 
      Caption         =   "CreateID ："
      Height          =   255
      Left            =   180
      TabIndex        =   24
      Top             =   5190
      Width           =   975
   End
   Begin VB.Label CDT 
      Height          =   255
      Left            =   2220
      TabIndex        =   23
      Top             =   5190
      Width           =   1095
   End
   Begin VB.Label CTM 
      Height          =   255
      Left            =   3405
      TabIndex        =   22
      Top             =   5190
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "UpdateID："
      Height          =   255
      Left            =   180
      TabIndex        =   21
      Top             =   5460
      Width           =   975
   End
   Begin VB.Label UDT 
      Height          =   255
      Left            =   2220
      TabIndex        =   20
      Top             =   5460
      Width           =   1095
   End
   Begin VB.Label UTM 
      Height          =   255
      Left            =   3405
      TabIndex        =   19
      Top             =   5460
      Width           =   975
   End
End
Attribute VB_Name = "frm075006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/16 改成Form2.0 ;lbeCusName(index)、cboContact、txtCaseName、IDName、UIDName、lblhc21、txtMemo
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim LcTmp As String
Dim m_Cpnum As String
'Add By Sindy 2011/5/31
' 第一筆資料的本所案號
Dim m_FirstRow(4) As String
' 最後一筆資料的本所案號
Dim m_LastRow(4) As String
' 目前正在顯示的本所案號
Dim m_CurrRow(4) As String
'2011/5/31 End
' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Public m_EditMode As Integer
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員

'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If txtCP01.Text = "" Or txtCP02.Text = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Exit Sub
   End If
   'Added by Lydia 2020/05/05
   If m_EditMode <> 0 And m_EditMode <> 4 Then
      MsgBox IIf(m_EditMode = 1, "新增中", "修改中") & "不可執行！", vbInformation
      Exit Sub
   End If
   'end 2020/05/05
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   
   frm12040159.SetParent "E", Trim(txtCP01.Text & txtCP02.Text & txtCP03.Text & txtCP04.Text), Me
   frm12040159.Show
End Sub

Private Sub cmdOther_Click()
   Dim i As Integer
   Dim strNum As String
   Dim strTmp As String
   If txtCP01.Text = "" And txtCP02.Text = "" Then
      MsgBox "請輸入本所案號", vbInformation, "顧問案件資料維護"
      Exit Sub
   End If
   Set frm1103_2.m_form = Me
   frm1103_2.intWhereComeFrom = 1
   frm1103_2.lblSystem = txtCP01.Text
   frm1103_2.lblCode(0) = txtCP02.Text
   frm1103_2.lblCode(1) = txtCP03.Text
   frm1103_2.lblCode(2) = txtCP04.Text
   frm1103_2.Show
   Me.Hide
End Sub

'add by nickc 2006/11/10 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case 13:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
   End Select
End Sub

Private Sub Form_Load()
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm075006", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm075006", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm075006", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm075006", strFind, False)
   ' Ken 90.07.16 -- End
   
   m_EditMode = 0
   MoveFormToCenter Me
   txtCP01 = "LA" '2011/5/20 add by sonia
   
   If Not IsEmptyText(m_CurrRow(0)) And Not IsEmptyText(m_CurrRow(1)) And Not IsEmptyText(m_CurrRow(2)) And Not IsEmptyText(m_CurrRow(3)) Then
      ShowCurrRecord m_CurrRow(0), m_CurrRow(1), m_CurrRow(2), m_CurrRow(3)
      UpdateToolbarState
      SetCtrlReadOnly True
   Else
      m_EditMode = 4
      SetCtrlReadOnly True
      SetKeyReadOnly False
      UpdateToolbarState
   End If
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
   End If
   'end 2020/05/05
End Sub

' 按下按鍵
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
      ' 確定, 取消
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      ' 取消或離開
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub txtCaseName_GotFocus()
   TextInverse txtCaseName
   OpenIme
End Sub

'Added by Lydia 2021/09/16
Private Sub txtCaseName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtCaseName  'Form 2.0的TextBox增加右鍵選單功能; 經過測試MouseMove無效,要放在MouseDown
End Sub

Private Sub txtCaseName_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtCaseName, 40) = False Then
      Cancel = True
      txtCaseName_GotFocus
   End If
End Sub

Private Sub txtCaseNum_GotFocus()
   TextInverse txtCaseNum
   CloseIme
End Sub

Private Sub txtCaseNum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCaseNum_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtCaseNum, 50) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "分所號內容太長"
      txtCaseNum.SetFocus
      txtCaseNum_GotFocus
   End If
End Sub

Private Sub txtCloseDate_GotFocus()
   TextInverse txtCloseDate
   CloseIme
End Sub

Private Sub txtCloseDate_Validate(Cancel As Boolean)
   If txtCloseDate <> "" Then
      If Not CheckIsTaiwanDate(txtCloseDate) Then
         Cancel = True
       End If
   End If
   If Cancel Then TextInverse txtCloseDate
End Sub

Private Sub txtCloseYN_GotFocus()
   TextInverse txtCloseYN
   CloseIme
End Sub

Private Sub txtCloseYN_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCloseYN_Validate(Cancel As Boolean)
   If txtCloseYN <> "" Then
   txtCloseYN = UCase(txtCloseYN)
   If txtCloseYN <> "Y" Then
      DataErrorMessage 1, "是否閉卷"
      Cancel = True
   End If
   If Cancel Then TextInverse txtCloseYN
   End If
End Sub

Private Sub txtClsResult_GotFocus()
   TextInverse txtClsResult
   CloseIme
End Sub

Private Sub txtClsResult_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtClsResult_Validate(Cancel As Boolean)
Dim strTempName As String
   lbeResult = ""
   If txtClsResult <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.GetReasonOfRelief(txtClsResult, strTempName) Then
      If ClsLawGetReasonOfRelief(txtClsResult, strTempName) Then
         lbeResult = strTempName
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtClsResult
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtCP01
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtcp01.IMEMode = 2
   CloseIme
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If txtCP01 <> "" Then
      txtCP01 = UCase(txtCP01)
      If txtCP01 = "LA" Then
'         blnCom1 = True
      Else
         DataErrorMessage 1, "系統類別"
'         blnCom1 = False
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtCP01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtCP02
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtcp02.IMEMode = 2
   CloseIme
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
'    '若有輸入本所案號的流水號
'    If txtcp02 <> "" Then
'        'Add By Cheng 2003/02/25
'        If GetSerialNo(Me.txtcp01.Text, Me.txtcp02.Text) Then
'            blnCom2 = True
'        Else
'            Cancel = True
'        End If
'    End If
'    If Cancel Then TextInverse txtcp02
Dim strTemp As String
   
   If txtCP02 <> "" Then
      strTemp = GiveSymbol(txtCP01, txtCP02, txtCP03, txtCP04, LcTmp)
      m_Cpnum = strTemp
      If ClsPDChkCaseNum(txtCP01, txtCP02) Then
         TextInverse txtCP02
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtCP02
End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtCP03
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtcp03.IMEMode = 2
   CloseIme
End Sub

Private Sub txtcp03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub GetPreHire()
  Dim RsTemp As New ADODB.Recordset
  Dim strCP53 As String
  Dim strCP54 As String
  
  m_CP01 = txtCP01.Text
  m_CP02 = txtCP02.Text
  If txtCP03.Text <> "" Then
     m_CP03 = txtCP03.Text
  Else
     m_CP03 = "0"
  End If
  If txtCP04.Text <> "" Then
     m_CP04 = txtCP04.Text
  Else
     m_CP04 = "00"
  End If
  
  strSql = "SELECT cp53,cp54,cp09 from caseprogress, casepropertymap where " & _
           " CP01='" & m_CP01 & "' AND CP02 ='" & m_CP02 & "' AND CP03 ='" & m_CP03 & "'" & _
           " AND CP04 ='" & m_CP04 & "' and cpm03='顧問聘任' and cp01=cpm01(+) AND cp10=cpm02(+) order by cp05 "
           
  RsTemp.CursorLocation = adUseClient
  RsTemp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
  If Not RsTemp.EOF Then
     RsTemp.MoveFirst
     lbeFirstHire.Caption = IIf(IsNull(RsTemp.Fields!cp53), "", RsTemp.Fields!cp53)
     lbeFirstHire.Caption = ChangeTStringToTDateString(ChangeWStringToTString(lbeFirstHire.Caption))
     RsTemp.MoveLast
     strCP53 = IIf(IsNull(RsTemp.Fields!cp53), "", RsTemp.Fields!cp53)
     strCP54 = IIf(IsNull(RsTemp.Fields!cp54), "", RsTemp.Fields!cp54)
     lbeHire.Caption = ChangeTStringToTDateString(ChangeWStringToTString(strCP53)) + "--" + ChangeTStringToTDateString(ChangeWStringToTString(strCP54))
  End If
  RsTemp.Close
  Set RsTemp = Nothing
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtCP04
   CloseIme
End Sub

Private Sub txtcp04_LostFocus()
'   If txtcp04 <> "" Then
'      blnCom4 = True
'   End If
'   If blnIsNew Then
   ' 新增模式下檢查資料是否已存在資料庫中
   If m_EditMode = 1 Then
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.CheckIsExistCaseNum(1, LcTmp, m_Cpnum) Then
         If ClsPDCheckIsExistCaseNum(1, LcTmp, m_Cpnum) Then
         'If CheckIsExistCaseNum(1, LcTmp, m_Cpnum) Then
'            blnCom2 = True
            'TextInverse txtcp02
            'txtcp02.SetFocus
            tlbar.Buttons(11).Enabled = True
         Else
            MsgBox "" + m_Cpnum + "", vbCritical
            TextInverse txtCP02
            txtCP02.SetFocus
         End If
   End If
'   tlbar.Buttons(11).Enabled = True
End Sub

'Modify By Sindy 2011/1/17
Private Sub txtCustomer_GotFocus(Index As Integer)
   TextInverse txtCustomer(Index)
   CloseIme
End Sub

'Modify By Sindy 2011/1/17
Private Sub txtCustomer_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify By Sindy 2011/1/17
Private Sub txtCustomer_Validate(Index As Integer, Cancel As Boolean)
Dim strCusTemp As String, StrCusName As String, i As Integer
   lbeCusName(Index) = ""
   If txtCustomer(Index) <> "" Then
      txtCustomer(Index) = UCase(txtCustomer(Index).Text)
      If Left(txtCustomer(Index).Text, 1) = "Y" Then
         txtCustomer(Index) = "X" & Mid(txtCustomer(Index), 2)
      ElseIf Left(txtCustomer(Index).Text, 1) <> "X" Then
         MsgBox "當事人代碼輸入錯誤!", vbExclamation, "顧問案件資料維護"
         Cancel = True
         TextInverse txtCustomer(Index)
         Exit Sub
      End If
      strCusTemp = txtCustomer(Index)
      If ClsPDGetCustomer(strCusTemp, StrCusName) Then
         lbeCusName(Index) = StrCusName
         If Index = 0 And txtCaseName = "" Then txtCaseName = lbeCusName(0)
      Else
         Cancel = True
         TextInverse txtCustomer(Index)
         Exit Sub
      End If
      'Add By Sindy 2011/1/17 檢查輸入當事人的順序
      If (txtCustomer(1) <> "" And txtCustomer(0) = "") Or _
         (txtCustomer(2) <> "" And txtCustomer(1) = "") Or _
         (txtCustomer(3) <> "" And txtCustomer(2) = "") Or _
         (txtCustomer(4) <> "" And txtCustomer(3) = "") Then
         MsgBox "請依序輸入當事人!", vbExclamation, "法務案件基本資料維護"
         If txtCustomer(1) <> "" And txtCustomer(0) = "" Then txtCustomer(1).SetFocus: Call txtCustomer_GotFocus(1)
         If txtCustomer(2) <> "" And txtCustomer(1) = "" Then txtCustomer(2).SetFocus: Call txtCustomer_GotFocus(2)
         If txtCustomer(3) <> "" And txtCustomer(2) = "" Then txtCustomer(3).SetFocus: Call txtCustomer_GotFocus(3)
         If txtCustomer(4) <> "" And txtCustomer(3) = "" Then txtCustomer(4).SetFocus: Call txtCustomer_GotFocus(4)
         Cancel = True
         Exit Sub
      End If
      'Add By Sindy 2011/1/17 檢查當事人不可重複
      If Index = 0 Then
         If txtCustomer(Index) = txtCustomer(1) Or _
            txtCustomer(Index) = txtCustomer(2) Or _
            txtCustomer(Index) = txtCustomer(3) Or _
            txtCustomer(Index) = txtCustomer(4) Then
            MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
            txtCustomer(Index).SetFocus
            txtCustomer_GotFocus (Index)
            Cancel = True
            Exit Sub
         End If
      End If
      If Index = 1 Then
         If txtCustomer(Index) = txtCustomer(0) Or _
            txtCustomer(Index) = txtCustomer(2) Or _
            txtCustomer(Index) = txtCustomer(3) Or _
            txtCustomer(Index) = txtCustomer(4) Then
            MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
            txtCustomer(Index).SetFocus
            txtCustomer_GotFocus (Index)
            Cancel = True
            Exit Sub
         End If
      End If
      If Index = 2 Then
         If txtCustomer(Index) = txtCustomer(0) Or _
            txtCustomer(Index) = txtCustomer(1) Or _
            txtCustomer(Index) = txtCustomer(3) Or _
            txtCustomer(Index) = txtCustomer(4) Then
            MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
            txtCustomer(Index).SetFocus
            txtCustomer_GotFocus (Index)
            Cancel = True
            Exit Sub
         End If
      End If
      If Index = 3 Then
         If txtCustomer(Index) = txtCustomer(0) Or _
            txtCustomer(Index) = txtCustomer(1) Or _
            txtCustomer(Index) = txtCustomer(2) Or _
            txtCustomer(Index) = txtCustomer(4) Then
            MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
            txtCustomer(Index).SetFocus
            txtCustomer_GotFocus (Index)
            Cancel = True
            Exit Sub
         End If
      End If
      If Index = 4 Then
         If txtCustomer(Index) = txtCustomer(0) Or _
            txtCustomer(Index) = txtCustomer(1) Or _
            txtCustomer(Index) = txtCustomer(2) Or _
            txtCustomer(Index) = txtCustomer(3) Then
            MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
            txtCustomer(Index).SetFocus
            txtCustomer_GotFocus (Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtMemo_GotFocus()
   TextInverse txtMemo
   OpenIme
End Sub

'Added by Lydia 2021/09/16
Private Sub txtMemo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtMemo  'Form 2.0的TextBox增加右鍵選單功能; 經過測試MouseMove無效,要放在MouseDown
End Sub

Private Sub txtMemo_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtMemo, 2000) = False Then
      Cancel = True
      txtMemo_GotFocus
   End If
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

If Me.txtCaseName.Enabled = True Then
   Cancel = False
   txtCaseName_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtCloseDate.Enabled = True Then
   Cancel = False
   txtCloseDate_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtCloseYN.Enabled = True Then
   Cancel = False
   txtCloseYN_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtClsResult.Enabled = True Then
   Cancel = False
   txtClsResult_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtCP01.Enabled = True Then
   Cancel = False
   txtcp01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtCP02.Enabled = True Then
   Cancel = False
   txtcp02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Modify By Sindy 2011/1/17
For ii = 0 To 4
   If Me.txtCustomer(ii).Enabled = True Then
      Cancel = False
      txtCustomer_Validate ii, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
If Pub_ChkAppList(strExc(0), txtCustomer(0) & "," & txtCustomer(1) & "," & txtCustomer(2) & "," & txtCustomer(3) & "," & txtCustomer(4)) = False Then
   txtCustomer(Val(strExc(0))).SetFocus
   txtCustomer_GotFocus Val(strExc(0))
   Exit Function
End If
'end 2024/06/14
   
'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
For ii = 0 To 4
   strExc(1) = ChangeCustomerL(txtCustomer(ii))
   strExc(2) = ChangeCustomerL(txtCustomer(ii).Tag)
   If strExc(1) <> "" And strExc(1) <> strExc(2) Then
      If GetCustomerAndState(strExc(1), strExc(3), , , , txtCP01, strExc(8), False, Me.Name, txtCP02, txtCP03, txtCP04) = False Then
         txtCustomer(ii).SetFocus
         txtCustomer_GotFocus ii
         Exit Function
      End If
   End If
Next ii
'end 2024/06/13
   
If Me.txtMemo.Enabled = True Then
   Cancel = False
   txtMemo_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Morgan 2007/5/10
If Not ((txtCloseYN.Text = "" And txtCloseDate.Text = "" And txtClsResult.Text = "") Or (txtCloseYN.Text <> "" And txtCloseDate.Text <> "" And txtClsResult.Text <> "")) Then
   MsgBox "是否閉卷、閉卷日期、閉卷原因三個欄位須同時空白或有值！", vbExclamation
   Exit Function
End If
'end 2007/5/10

'Added by Lydia 2021/09/16 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   'Add By Sindy 2011/5/31
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyF5 Then
      m_CP01 = txtCP01.Text
      m_CP02 = txtCP02.Text
      If txtCP03.Text <> Empty Then
         m_CP03 = txtCP03.Text
      Else
         m_CP03 = "0"
      End If
      If txtCP04.Text <> Empty Then
         m_CP04 = txtCP04.Text
      Else
         m_CP04 = "00"
      End If
      
      ' 檢查記錄是否不存在
      If IsRecordExist(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
         strTit = "檢查"
         strMsg = "無此資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Exit Sub
      End If
   End If
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         If IsCaseProgressExist(txtCP01, txtCP02, txtCP03, txtCP04) = True Then
            strTit = "檢核資料"
            strMsg = "此本所案號在案件進度檔中仍有資料, 不可刪除!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Else
            'Add By Sindy 2010/7/1
            If ChkCaseCode("NP", txtCP01, txtCP02, txtCP03, txtCP04) = False Then Exit Sub
            '2010/7/1 End
            strTit = "詢問"
            strMsg = "是否要刪除此筆資料?"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
               m_EditMode = 3
               OnWork
               UpdateToolbarState
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
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'edit by nickc 2008/03/28 還沒檢查完資料就先更新，有些資料在檢查時才上，會更新不到
'         UpdateFieldNewData
         OnWork
         UpdateToolbarState
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
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            If Not IsEmptyText(m_FirstRow(0)) And Not IsEmptyText(m_FirstRow(1)) And Not IsEmptyText(m_FirstRow(2)) And Not IsEmptyText(m_FirstRow(3)) Then
               tlbar.Buttons(6).Enabled = True
               tlbar.Buttons(7).Enabled = True
            Else
               tlbar.Buttons(6).Enabled = False
               tlbar.Buttons(7).Enabled = False
            End If
            If Not IsEmptyText(m_LastRow(0)) And Not IsEmptyText(m_LastRow(1)) And Not IsEmptyText(m_LastRow(2)) And Not IsEmptyText(m_LastRow(3)) Then
               tlbar.Buttons(8).Enabled = True
               tlbar.Buttons(9).Enabled = True
            Else
               tlbar.Buttons(8).Enabled = False
               tlbar.Buttons(9).Enabled = False
            End If
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_FirstRow(0) = Empty
   m_FirstRow(1) = Empty
   m_FirstRow(2) = Empty
   m_FirstRow(3) = Empty
   m_CurrRow(0) = Empty
   m_CurrRow(1) = Empty
   m_CurrRow(2) = Empty
   m_CurrRow(3) = Empty
   m_LastRow(0) = Empty
   m_LastRow(1) = Empty
   m_LastRow(2) = Empty
   m_LastRow(3) = Empty
   m_EditMode = 0
   Set frm075006 = Nothing
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   ' 設定 Query 的命令
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & txtCP01 & "' AND " & _
                  "HC02 = (SELECT MIN(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "') AND " & _
                  "HC03 = (SELECT MIN(HC03) FROM HireCase WHERE HC01 = '" & txtCP01 & "' AND HC02 = (SELECT MIN(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "' )) AND " & _
                  "HC04 = (SELECT MIN(HC04) FROM HireCase WHERE HC01 = '" & txtCP01 & "' AND HC02 = (SELECT MIN(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "' ) AND HC03 = (SELECT MIN(HC03) FROM HireCase WHERE HC01 = '" & txtCP01 & "' AND HC02 = (SELECT MIN(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "' ))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_FirstRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_FirstRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_FirstRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_FirstRow(3) = rsTmp.Fields("HC04")
   End If
   rsTmp.Close

   ' 設定 Query 的命令
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & txtCP01 & "' AND " & _
                  "HC02 = (SELECT MAX(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "') AND " & _
                  "HC03 = (SELECT MAX(HC03) FROM HireCase WHERE HC01 = '" & txtCP01 & "' AND HC02 = (SELECT MAX(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "' )) AND " & _
                  "HC04 = (SELECT MAX(HC04) FROM HireCase WHERE HC01 = '" & txtCP01 & "' AND HC02 = (SELECT MAX(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "' ) AND HC03 = (SELECT MAX(HC03) FROM HireCase WHERE HC01 = '" & txtCP01 & "' AND HC02 = (SELECT MAX(HC02) FROM HireCase WHERE HC01 = '" & txtCP01 & "' ))) "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_LastRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_LastRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_LastRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_LastRow(3) = rsTmp.Fields("HC04")
   End If
   rsTmp.Close
  
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strHC01, strHC02, strHC03, strHC04) = True Then
      m_CurrRow(0) = strHC01
      m_CurrRow(1) = strHC02
      m_CurrRow(2) = strHC03
      m_CurrRow(3) = strHC04
   Else
      strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
               "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                     "HC02 = '" & m_CurrRow(1) & "' AND " & _
                     "HC03 = '" & m_CurrRow(2) & "' AND " & _
                     "HC04 = (SELECT MIN(HC04) FROM HireCase " & _
                             "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                   "HC03 = '" & m_CurrRow(2) & "' AND " & _
                                   "HC04 > '" & m_CurrRow(3) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
         If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
         If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
         If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
   
      strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
               "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                     "HC02 = '" & m_CurrRow(1) & "' AND " & _
                     "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                             "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                   "HC03 > '" & m_CurrRow(2) & "') AND " & _
                     "HC04 = (SELECT MIN(HC04) FROM HireCase " & _
                             "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                   "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                                           "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                 "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                                 "HC03 > '" & m_CurrRow(2) & "'))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
         If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
         If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
         If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
                                
      strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
               "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                     "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                             "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "HC02 > '" & m_CurrRow(1) & "') AND " & _
                     "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                             "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                                           "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                 "HC02 > '" & m_CurrRow(1) & "')) AND " & _
                     "HC04 = (SELECT MIN(HC04) FROM HireCase " & _
                             "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                                           "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                 "HC02 > '" & m_CurrRow(1) & "') AND " & _
                                                 "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                                                         "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                               "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                                                                       "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                                             "HC02 > '" & m_CurrRow(1) & "'))) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
         If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
         If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
         If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
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
   m_CurrRow(0) = m_FirstRow(0)
   m_CurrRow(1) = m_FirstRow(1)
   m_CurrRow(2) = m_FirstRow(2)
   m_CurrRow(3) = m_FirstRow(3)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrRow(0) = m_FirstRow(0) And m_CurrRow(1) = m_FirstRow(1) And m_CurrRow(2) = m_FirstRow(2) And m_CurrRow(3) = m_FirstRow(3) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                  "HC02 = '" & m_CurrRow(1) & "' AND " & _
                  "HC03 = '" & m_CurrRow(2) & "' AND " & _
                  "HC04 = (SELECT MAX(HC04) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                "HC03 = '" & m_CurrRow(2) & "' AND " & _
                                "HC04 < '" & m_CurrRow(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                  "HC02 = '" & m_CurrRow(1) & "' AND " & _
                  "HC03 = (SELECT MAX(HC03) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                "HC03 < '" & m_CurrRow(2) & "') AND " & _
                  "HC04 = (SELECT MAX(HC04) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                "HC03 = (SELECT MAX(HC03) FROM HireCase " & _
                                        "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                              "HC03 < '" & m_CurrRow(2) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                  "HC02 = (SELECT MAX(HC02) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 < '" & m_CurrRow(1) & "') AND " & _
                  "HC03 = (SELECT MAX(HC03) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = (SELECT MAX(HC02) FROM HireCase " & _
                                        "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "HC02 < '" & m_CurrRow(1) & "')) AND " & _
                  "HC04 = (SELECT MAX(HC04) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = (SELECT MAX(HC02) FROM HireCase " & _
                                        "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "HC02 < '" & m_CurrRow(1) & "') AND " & _
                                              "HC03 = (SELECT MAX(HC03) FROM HireCase " & _
                                                      "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                            "HC02 = (SELECT MAX(HC02) FROM HireCase " & _
                                                                    "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                                          "HC02 < '" & m_CurrRow(1) & "'))) "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
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
   
   If m_CurrRow(0) = m_LastRow(0) And m_CurrRow(1) = m_LastRow(1) And m_CurrRow(2) = m_LastRow(2) And m_CurrRow(3) = m_LastRow(3) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                  "HC02 = '" & m_CurrRow(1) & "' AND " & _
                  "HC03 = '" & m_CurrRow(2) & "' AND " & _
                  "HC04 = (SELECT MIN(HC04) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                "HC03 = '" & m_CurrRow(2) & "' AND " & _
                                "HC04 > '" & m_CurrRow(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                  "HC02 = '" & m_CurrRow(1) & "' AND " & _
                  "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                "HC03 > '" & m_CurrRow(2) & "') AND " & _
                  "HC04 = (SELECT MIN(HC04) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                                        "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "HC02 = '" & m_CurrRow(1) & "' AND " & _
                                              "HC03 > '" & m_CurrRow(2) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
                                
   strSql = "SELECT HC01,HC02,HC03,HC04 FROM HireCase " & _
            "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                  "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 > '" & m_CurrRow(1) & "') AND " & _
                  "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                                        "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "HC02 > '" & m_CurrRow(1) & "')) AND " & _
                  "HC04 = (SELECT MIN(HC04) FROM HireCase " & _
                          "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                                        "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "HC02 > '" & m_CurrRow(1) & "') AND " & _
                                              "HC03 = (SELECT MIN(HC03) FROM HireCase " & _
                                                      "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                            "HC02 = (SELECT MIN(HC02) FROM HireCase " & _
                                                                    "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                                                                          "HC02 > '" & m_CurrRow(1) & "'))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("HC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("HC01")
      If IsNull(rsTmp.Fields("HC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("HC02")
      If IsNull(rsTmp.Fields("HC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("HC03")
      If IsNull(rsTmp.Fields("HC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("HC04")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrRow(0) = m_LastRow(0)
   m_CurrRow(1) = m_LastRow(1)
   m_CurrRow(2) = m_LastRow(2)
   m_CurrRow(3) = m_LastRow(3)
   
   UpdateCtrlData
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
Dim i As Integer

   txtCP01 = Empty
   txtCP02 = Empty
   txtCP03 = Empty
   txtCP04 = Empty
   'Modify By Sindy 2011/1/17
   For i = 0 To 4
      txtCustomer(i) = Empty
      lbeCusName(i) = Empty
      txtCustomer(i).Tag = Empty 'Added by Lydia 2024/06/13
   Next i
   '2011/1/17 End
   txtCaseName = Empty
   txtCaseNum = Empty
   txtCloseDate = Empty
   txtCloseYN = Empty
   txtClsResult = Empty
   txtMemo = Empty
   lbeFirstHire = Empty
   lbeHire = Empty
   IDname = Empty
   CDT = Empty
   CTM = Empty
   UIDname = Empty
   UDT = Empty
   UTM = Empty
   'add by nickc 2006/07/12
   lblhc19 = Empty
   lblhc20 = Empty
   lblhc21 = Empty
   lblhc22 = Empty
   cboContact.Clear 'Add by Morgan 2008/8/4
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   txtCP01.Locked = bEnable: txtCP02.Locked = bEnable: txtCP03.Locked = bEnable: txtCP04.Locked = bEnable
   'Modify By Sindy 2011/1/14
   For i = 0 To 4
      txtCustomer(i).Locked = bEnable
   Next i
   '2011/1/14 End
   txtCaseName.Locked = bEnable
   txtCaseNum.Locked = bEnable
   txtCloseDate.Locked = bEnable
   txtCloseYN.Locked = bEnable
   txtClsResult.Locked = bEnable
   txtMemo.Locked = bEnable
   cboContact.Locked = bEnable 'Added by Lydia 2021/09/16

End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   txtCP01.Locked = bEnable: txtCP02.Locked = bEnable: txtCP03.Locked = bEnable: txtCP04.Locked = bEnable
End Sub

'Add By Sindy 2011/5/31
' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, strTemp As String
Dim strName As String
   
   strSql = "SELECT * FROM HireCase " & _
            "WHERE HC01 = '" & m_CurrRow(0) & "' AND " & _
                  "HC02 = '" & m_CurrRow(1) & "' AND " & _
                  "HC03 = '" & m_CurrRow(2) & "' AND " & _
                  "HC04 = '" & m_CurrRow(3) & "' "
               
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      txtCP01 = IIf(IsNull(rsTmp.Fields!hc01), "", rsTmp.Fields!hc01)
      txtCP02 = IIf(IsNull(rsTmp.Fields!hc02), "", rsTmp.Fields!hc02)
      txtCP03 = IIf(IsNull(rsTmp.Fields!hc03), "", rsTmp.Fields!hc03)
      txtCP04 = IIf(IsNull(rsTmp.Fields!hc04), "", rsTmp.Fields!hc04)
      'Modify By Sindy 2011/1/14
      For i = 0 To 4
         If i = 0 Then txtCustomer(i) = IIf(IsNull(rsTmp.Fields!hc05), "", rsTmp.Fields!hc05)
         If i = 1 Then txtCustomer(i) = IIf(IsNull(rsTmp.Fields!hc24), "", rsTmp.Fields!hc24)
         If i = 2 Then txtCustomer(i) = IIf(IsNull(rsTmp.Fields!hc25), "", rsTmp.Fields!hc25)
         If i = 3 Then txtCustomer(i) = IIf(IsNull(rsTmp.Fields!hc26), "", rsTmp.Fields!hc26)
         If i = 4 Then txtCustomer(i) = IIf(IsNull(rsTmp.Fields!hc27), "", rsTmp.Fields!hc27)
         If txtCustomer(i) <> "" Then
            If ClsPDGetCustomer(txtCustomer(i), strName) Then
               lbeCusName(i).Caption = strName
            End If
         End If
         txtCustomer(i).Tag = txtCustomer(i).Text 'Added by Lydia 2024/06/13
      Next
      '2011/1/14 End
      txtCaseName = IIf(IsNull(rsTmp.Fields!hc06), "", rsTmp.Fields!hc06)
      txtCaseNum = IIf(IsNull(rsTmp.Fields!hc07), "", rsTmp.Fields!hc07)
      If IsNull(rsTmp.Fields("hc19")) = False Then
         If IsEmptyText(rsTmp.Fields("hc19")) = False Then
            strTemp = TAIWANDATE(rsTmp.Fields("hc19"))
            lblhc19 = Format(strTemp, "###/##/##")
         End If
      End If
      If IsNull(rsTmp.Fields("hc20")) = False Then
         If IsEmptyText(rsTmp.Fields("hc20")) = False Then
            strTemp = TAIWANDATE(rsTmp.Fields("hc20"))
            lblhc20 = Format(strTemp, "###/##/##")
         End If
      End If
      If IsNull(rsTmp.Fields("hc21")) = False Then
         If IsEmptyText(rsTmp.Fields("hc21")) = False Then
            lblhc21 = GetStaffName(rsTmp.Fields("hc21"), True)
         End If
      End If
      lblhc22 = IIf(IsNull(rsTmp.Fields!hc22), "", rsTmp.Fields!hc22)
      txtCloseYN = IIf(IsNull(rsTmp.Fields!hc09), "", rsTmp.Fields!hc09)
      '2012/5/9 MODIFY BY SONIA LA-001103
      'txtCloseDate = IIf(IsNull(rsTmp.Fields!hc10), "", ChangeWStringToTString(rsTmp.Fields!hc10))
      If Not IsNull(rsTmp.Fields!hc10) Then
         txtCloseDate = ChangeWStringToTString(rsTmp.Fields!hc10)
      Else
         txtCloseDate = ""
      End If
      '2012/5/9 END
      txtClsResult = IIf(IsNull(rsTmp.Fields!hc11), "", rsTmp.Fields!hc11)
      txtMemo = IIf(IsNull(rsTmp.Fields!hc12), "", rsTmp.Fields!hc12)
      If IsNull(rsTmp.Fields("hc13")) = False Then
         If IsEmptyText(rsTmp.Fields("hc13")) = False Then
            IDname = GetStaffName(rsTmp.Fields("hc13"), True)
         End If
      End If
      If IsNull(rsTmp.Fields("hc14")) = False Then
         If IsEmptyText(rsTmp.Fields("hc14")) = False Then
            strTemp = TAIWANDATE(rsTmp.Fields("hc14"))
            CDT = Format(strTemp, "###/##/##")
         End If
      End If
      If IsNull(rsTmp.Fields("hc15")) = False Then
         If IsEmptyText(rsTmp.Fields("hc15")) = False Then
            strTemp = rsTmp.Fields("hc15")
            CTM = Format(strTemp, "##:##")
         End If
      End If
      If IsNull(rsTmp.Fields("hc16")) = False Then
         If IsEmptyText(rsTmp.Fields("hc16")) = False Then
            UIDname = GetStaffName(rsTmp.Fields("hc16"), True)
         End If
      End If
      If IsNull(rsTmp.Fields("hc17")) = False Then
         If IsEmptyText(rsTmp.Fields("hc17")) = False Then
            strTemp = TAIWANDATE(rsTmp.Fields("hc17"))
            UDT = Format(strTemp, "###/##/##")
         End If
      End If
      If IsNull(rsTmp.Fields("hc18")) = False Then
         If IsEmptyText(rsTmp.Fields("hc18")) = False Then
            strTemp = rsTmp.Fields("hc18")
            UTM = Format(strTemp, "##:##")
         End If
      End If
      'Modified by Lydia 2021/09/16 改成Form 2.0
      'PUB_AddContact "" & rsTmp.Fields("HC05"), cboContact, "" & rsTmp.Fields("HC23") 'Add by Morgan 2008/8/4
      PUB_AddContact "" & rsTmp.Fields("HC05"), cboContact, "" & rsTmp.Fields("HC23"), , True
      GetPreHire
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False
   
   If IsEmptyText(txtCP03) = True Then: txtCP03 = "0"
   If IsEmptyText(txtCP04) = True Then: txtCP04 = "00"
   
   If IsRecordExist(txtCP01, txtCP02, txtCP03, txtCP04) = True Then
      m_CurrRow(0) = txtCP01
      m_CurrRow(1) = txtCP02
      m_CurrRow(2) = txtCP03
      m_CurrRow(3) = txtCP04
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   ' 當系統別不為原先所輸入的系統別時則需重新取得範圍
   If txtCP01 <> m_CurrRow(0) Then
      RefreshRange
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
Dim strMsg As String
Dim strTit As String
Dim nResponse
Dim StrSQLa As String            '2009/8/19 ADD BY SONIA
Dim rsA As New ADODB.Recordset   '2009/8/19 ADD BY SONIA
   
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'add by nickc 2008/03/28  更新欄位
'            UpdateFieldNewData
            'edit by nickc 2006/06/08
            'AddRecord
            If AddRecord = False Then Exit Sub
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            strChkCuAreaMail = PUB_ChkSameCustSales(Trim(txtCP01), Trim(txtCP02), Trim(txtCP03), Trim(txtCP04), "", Trim(txtCustomer(0)), Trim(txtCustomer(1)), Trim(txtCustomer(2)), Trim(txtCustomer(3)), Trim(txtCustomer(4)), strChkCuAreaMailTo)
            
            'add by nickc 2008/03/28  更新欄位
'            UpdateFieldNewData
            'edit by nickc 2006/06/08
            'ModRecord
            If ModRecord = False Then Exit Sub
            
            'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            If strChkCuAreaMail <> "" Then
               PUB_SendMail strUserNum, strChkCuAreaMailTo, "", "案件收文通知--此案收文非原智權人員(區)！", strChkCuAreaMail
            End If
            'end 2017/06/19
         Else
            GoTo EXITSUB
         End If
      Case 3:
         'edit by nickc 2006/06/08
         'DelRecord
         If DelRecord = False Then Exit Sub
        'add by nickc 2008/03/28  更新欄位
'         UpdateFieldNewData
         
         RefreshRange
      Case 4:
         If TxtValidate = False Then Exit Sub
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: txtCP01.SetFocus
      Case 2: txtCustomer(0).SetFocus
      Case 4: txtCP01.SetFocus
   End Select
End Sub

' 案件進度檔
Private Function IsCaseProgressExist(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   IsCaseProgressExist = False
   strSql = "SELECT * from CaseProgress " & _
            "WHERE CP01 = '" & strHC01 & "' AND " & _
                  "CP02 = '" & strHC02 & "' AND " & _
                  "CP03 = '" & strHC03 & "' AND " & _
                  "CP04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsCaseProgressExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM HireCase " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "'"
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function CheckDataValid() As Boolean
Dim StrCusName As String
Dim strTempName As String
Dim i As Integer
   
   CheckDataValid = False
   
   If txtCP01.Text = "" Then
      MsgBox "本所案號不可空白!", vbExclamation, "顧問案件資料維護"
      TextInverse txtCP01
      Exit Function
   End If
   If txtCP02.Text = "" Then
      MsgBox "本所案號不可空白!", vbExclamation, "顧問案件資料維護"
      TextInverse txtCP02
      Exit Function
   End If
   
   If txtCaseName = "" Then
      DataErrorMessage 5, "案件名稱"
      txtCaseName.SetFocus
      Exit Function
   Else
      If CheckLengthIsOK(txtCaseName, 40) = False Then
         txtCaseName_GotFocus
         Exit Function
      End If
   End If
   
   '若為新增狀態
   If m_EditMode = 1 Then
       '若本所案號的流水號非6碼
       If Len(Me.txtCP02.Text) <> 6 Then
           MsgBox "本所案號的流水號必須為 6 碼，不滿 6 碼者請在前面補零!!!", vbExclamation + vbOKOnly
           txtCP02.SetFocus
           txtcp02_GotFocus
           Exit Function
       End If
   End If
   
   'Modify By Sindy 2011/1/17 +當事人2,3,4,5
   For i = 0 To 4
      If i = 0 And txtCustomer(0) = "" Then
         '資料不可為空
         DataErrorMessage 5, "當事人1"
         txtCustomer(0).SetFocus
         Exit Function
      End If
      '有輸入資料時, 檢查資料是否正確
      If txtCustomer(i) <> "" Then
         txtCustomer(i) = UCase(txtCustomer(i).Text)
         If Left(txtCustomer(i).Text, 1) = "Y" Then
            txtCustomer(i) = "X" & Mid(txtCustomer(i), 2)
         ElseIf Left(txtCustomer(i).Text, 1) <> "X" Then
            MsgBox "當事人" & CStr(i + 1) & "代碼輸入錯誤!", vbExclamation, "顧問案件資料維護"
            TextInverse txtCustomer(i)
            Exit Function
         End If
         If ClsPDGetCustomer(txtCustomer(i), StrCusName) Then
            lbeCusName(i) = StrCusName
         Else
            TextInverse txtCustomer(i)
            Exit Function
         End If
      End If
   Next i
   
   If txtCloseYN <> "" Then
      txtCloseYN = UCase(txtCloseYN)
      If txtCloseYN <> "Y" Then
         DataErrorMessage 1, "是否閉卷"
         TextInverse txtCloseYN
         Exit Function
      End If
   End If
   
   If txtCloseDate <> "" Then
      If Not CheckIsTaiwanDate(txtCloseDate) Then
         TextInverse txtCloseDate
         Exit Function
      End If
   End If
    
   If txtClsResult <> "" Then
      lbeResult = ""
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.GetReasonOfRelief(txtClsResult, strTempName) Then
      If ClsLawGetReasonOfRelief(txtClsResult, strTempName) Then
         lbeResult = strTempName
      Else
         TextInverse txtClsResult
         Exit Function
      End If
   End If
   
   If txtMemo <> "" Then
      If CheckLengthIsOK(txtMemo, 2000) = False Then
         txtMemo_GotFocus
         Exit Function
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 新增記錄
Private Function AddRecord() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim j As Integer
   
   AddRecord = False
   
   m_CP01 = txtCP01.Text
   m_CP02 = txtCP02.Text
   If txtCP03.Text <> "" Then
      m_CP03 = txtCP03.Text
   Else
      m_CP03 = "0"
   End If
   If txtCP04.Text <> "" Then
      m_CP04 = txtCP04.Text
   Else
      m_CP04 = "00"
   End If
   'Add By Cheng 2003/02/25
   '若有輸入當事人, 則要補滿9碼
   'Modify By Sindy 2011/1/17
   For j = 0 To 4
      If Me.txtCustomer(j).Text <> "" Then
         Me.txtCustomer(j).Text = Left(Me.txtCustomer(j).Text & "000000000", 9)
      End If
   Next j
   
   ' 檢查記錄是否已存在
   If IsRecordExist(m_CP01, m_CP02, m_CP03, m_CP04) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   CDT = ChangeWDateStringToWString(Date)
   CTM = Format(time, "HHMM")
   'Modify By Sindy 2011/1/17 +hc24,hc25,hc26,hc27
   strExc(1) = "insert into hirecase(hc01,hc02,hc03,hc04,hc05,hc06,hc07,hc09," + _
               "hc10,hc11,hc12,hc24,hc25,hc26,hc27) values (" & CNULL(m_CP01) & "," + _
               CNULL(m_CP02) + "," + CNULL(m_CP03) + "," + CNULL(m_CP04) + "," + _
               CNULL(txtCustomer(0)) + "," + CNULL(txtCaseName) + "," + CNULL(txtCaseNum) + _
               "," + CNULL(txtCloseYN) + "," + CNULL(DBDATE(txtCloseDate)) + "," + _
               CNULL(txtClsResult) + "," + CNULL(ChgSQL(txtMemo)) + "," + _
               CNULL(txtCustomer(1)) + "," + CNULL(txtCustomer(2)) + "," + _
               CNULL(txtCustomer(3)) + "," + CNULL(txtCustomer(4)) + ")"
   
   On Error GoTo EXITSUB
   cnnConnection.BeginTrans
   'add by nickc 2006/06/07 紀錄分析語法 and 修改原先更新方式
   Pub_SeekTbLog strExc(1)
   cnnConnection.Execute strExc(1)
   cnnConnection.CommitTrans
   
   If ((m_CP01 & m_CP02 & m_CP03 & m_CP04) < (m_FirstRow(0) & m_FirstRow(1) & m_FirstRow(2) & m_FirstRow(3))) Or ((m_CP01 & m_CP02 & m_CP03 & m_CP04) > (m_LastRow(0) & m_LastRow(1) & m_LastRow(2) & m_LastRow(3))) Then
      RefreshRange
   End If
   
   ShowCurrRecord m_CP01, m_CP02, m_CP03, m_CP04
   AddRecord = True
   
EXITSUB:
Exit Function
oErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim j As Integer
   
   ModRecord = False
   
   m_CP01 = txtCP01.Text
   m_CP02 = txtCP02.Text
   If txtCP03.Text <> "" Then
      m_CP03 = txtCP03.Text
   Else
      m_CP03 = "0"
   End If
   If txtCP04.Text <> "" Then
      m_CP04 = txtCP04.Text
   Else
      m_CP04 = "00"
   End If
   'Add By Cheng 2003/02/25
   '若有輸入當事人, 則要補滿9碼
   'Modify By Sindy 2011/1/17
   For j = 0 To 4
      If Me.txtCustomer(j).Text <> "" Then
         Me.txtCustomer(j).Text = Left(Me.txtCustomer(j).Text & "000000000", 9)
      End If
   Next j
   
   UDT = GetTodayDate
   UTM = Format(time, "HHMM")
   'Modify By Sindy 2011/1/17 +hc24,hc25,hc26,hc27
   strExc(1) = " update hirecase set hc05=" + CNULL(txtCustomer(0)) + ",hc06=" + CNULL(txtCaseName) + ",hc07=" + CNULL(txtCaseNum) + _
   ",hc08=NULL,hc09=" + CNULL(txtCloseYN) + ",hc10=" + CNULL(DBDATE(txtCloseDate)) + ",hc11=" + CNULL(txtClsResult) + _
   ",hc12=" + CNULL(ChgSQL(txtMemo)) + _
   ",hc24=" + CNULL(txtCustomer(1)) + _
   ",hc25=" + CNULL(txtCustomer(2)) + _
   ",hc26=" + CNULL(txtCustomer(3)) + _
   ",hc27=" + CNULL(txtCustomer(4)) + _
   " where hc01 ='" & m_CP01 & "' AND HC02 ='" & m_CP02 & "' AND HC03 ='" & m_CP03 & "' AND HC04 ='" & m_CP04 & "' "
   
   On Error GoTo EXITSUB
   cnnConnection.BeginTrans
   'add by nickc 2006/06/07 紀錄分析語法
   Pub_SeekTbLog strExc(1)
   strExc(1) = "begin user_data.user_enabled:=1;" & strExc(1) & "; end;"
   cnnConnection.Execute strExc(1)
   cnnConnection.CommitTrans
   
   'add by nickc 2005/08/23 紀錄修改案號
   pub_ModifyCaseNum = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   
   ShowCurrRecord m_CP01, m_CP02, m_CP03, m_CP04
   ModRecord = True
   
EXITSUB:
   Exit Function
ErrHand:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   DelRecord = False
   
   m_CP01 = txtCP01.Text
   m_CP02 = txtCP02.Text
   If txtCP03.Text <> "" Then
      m_CP03 = txtCP03.Text
   Else
      m_CP03 = "0"
   End If
   If txtCP04.Text <> "" Then
      m_CP04 = txtCP04.Text
   Else
      m_CP04 = "00"
   End If
   
   'Add By Sindy 2010/7/1
   If ChkCaseCode("CP", m_CP01, m_CP02, m_CP03, m_CP04) = False Then Exit Function
   If ChkCaseCode("NP", m_CP01, m_CP02, m_CP03, m_CP04) = False Then Exit Function
   '2010/7/1 End
   
'   If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
   If OnDataDeleteRecord(0, m_CP01 & m_CP02 & m_CP03 & m_CP04) <> 0 Then
      GoTo EXITSUB
   End If
   strExc(1) = "DELETE FROM HIRECASE WHERE HC01 ='" & m_CP01 & "'" & _
               " AND HC02 ='" & m_CP02 & "' AND HC03 ='" & m_CP03 & "'" & _
               " AND HC04 ='" & m_CP04 & "'"
   
   'add by nickc 2006/06/07 紀錄分析語法 and 修改原先方式
   On Error GoTo oErr
   cnnConnection.BeginTrans
   
      Pub_SeekTbLog strExc(1)
      cnnConnection.Execute strExc(1)
      
      'Added by Lydia 2016/11/24 一併刪除各項指示
      strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(m_CP01)) & " AND ITS02=" & CNULL(m_CP01 & m_CP02 & m_CP03 & m_CP04)
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      'end 2016/11/24
   cnnConnection.CommitTrans
'   End If
   
   DelRecord = True
   
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (m_CP01 = m_LastRow(0) And m_CP02 = m_LastRow(1) And m_CP03 = m_LastRow(2) And m_CP04 = m_LastRow(3)) Or (m_CP01 = m_FirstRow(0) And m_CP02 = m_FirstRow(1) And m_CP03 = m_FirstRow(2) And m_CP04 = m_FirstRow(3)) Then
      RefreshRange
   End If
   ShowCurrRecord m_CP01, m_CP02, m_CP03, m_CP04
   
EXITSUB:
   Exit Function
oErr:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function
