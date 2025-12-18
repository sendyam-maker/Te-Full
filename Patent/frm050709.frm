VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050709 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "客戶發明人資料維護"
   ClientHeight    =   5610
   ClientLeft      =   180
   ClientTop       =   990
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "發明人同申請人(&C)"
      Enabled         =   0   'False
      Height          =   400
      Index           =   0
      Left            =   6660
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   690
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8460
      Top             =   720
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
            Picture         =   "frm050709.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050709.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
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
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   2610
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5250
      Width           =   6075
      VariousPropertyBits=   16415
      BackColor       =   16777215
      Size            =   "10716;503"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Top             =   4860
      Width           =   6885
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "12144;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   10
      Top             =   5670
      Visible         =   0   'False
      Width           =   6855
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   675
      Index           =   7
      Left            =   1800
      TabIndex        =   8
      Top             =   3840
      Width           =   6885
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "12144;1191"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   615
      Index           =   6
      Left            =   1800
      TabIndex        =   7
      Top             =   3180
      Width           =   6885
      VariousPropertyBits=   -1466941413
      MaxLength       =   150
      ScrollBars      =   2
      Size            =   "12144;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   615
      Index           =   5
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   6885
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "12144;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   5
      Top             =   2220
      Width           =   6885
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "12144;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
      Width           =   6885
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "12144;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Top             =   1620
      Width           =   6885
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "12144;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   1020
      Width           =   495
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "873;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   4560
      Width           =   645
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1138;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "法定代表人職稱:"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   27
      Top             =   4890
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   180
      Left            =   360
      TabIndex        =   24
      Top             =   1320
      Width           =   228
   End
   Begin MSForms.Label LblIn01 
      Height          =   285
      Left            =   2940
      TabIndex        =   23
      Top             =   720
      Width           =   3660
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6456;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "代表人:"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   5670
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "發明人地址(日):"
      Height          =   180
      Index           =   5
      Left            =   360
      TabIndex        =   21
      Top             =   3840
      Width           =   1248
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "發明人地址(英):"
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   20
      Top             =   3180
      Width           =   1248
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "發明人名稱(日):"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   19
      Top             =   2220
      Width           =   1248
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "發明人名稱(英):"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   1248
   End
   Begin MSForms.Label LblIn10 
      Height          =   285
      Left            =   2490
      TabIndex        =   17
      Top             =   4560
      Width           =   2730
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4815;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "國籍:"
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   4590
      Width           =   408
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "發明人地址(中):"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   2520
      Width           =   1248
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "發明人名稱(中):"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1620
      Width           =   1248
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "發明人代號:"
      Height          =   180
      Left            =   360
      TabIndex        =   13
      Top             =   1050
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號:"
      Height          =   180
      Left            =   360
      TabIndex        =   12
      Top             =   750
      Width           =   948
   End
End
Attribute VB_Name = "frm050709"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim Rs As ADODB.Recordset
Dim in0(0 To 12) As String
'Modify by Morgan 2010/1/11 改全域變數
'Dim Data_Mission As Integer
Public Data_Mission As Integer '1 新增,2 刪除,3 修改,4 查詢

'edit by nickc 2007/02/06 不用 dll 了
'Dim obj0701 As Object
Dim strNo As String
Dim strName As String
'Dim Fld1 As String
'Dim Fld2 As String
'Dim Fld3 As String
'Dim Fld4 As String
Dim DelFlg As Boolean
Dim RsCounts As Integer
Dim nRet As Boolean
Dim InitValue As Boolean
Dim GetNowData As Boolean
Dim ChkData As Boolean
Dim BlnULetter As Boolean
Dim blnKeypreview As Boolean

'Add By Sindy 2012/12/25
' 目前正在顯示的Key值
Dim m_CurrKEY(2) As String
'2012/12/25 End

Public bAddOnly As Boolean '新增後回前表單 Added by Morgan 2020/2/19
Public fmCallForm As Form '前表單 Added by Morgan 2020/2/19


Private Sub cmdOK_Click(Index As Integer)
   'Add By Cheng 2002/01/09
   Select Case Index
      Case 0 '發明人同申請人
         If Me.Text1(0).Text <> "" Then
            ShowCustData
         End If
   End Select
End Sub

Private Sub Form_Activate()
   Static bolActivated As Boolean
   If bolActivated = False And bAddOnly = True Then
      If Data_Mission = 1 Then
         SendKeys "{Tab}"
         SendKeys "{Tab}"
      End If
      bolActivated = True
   End If
End Sub

Private Sub Form_Load()
   Me.Tag = Me.Caption 'Added by Morgan 2020/2/19
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   Data_Mission = 0
'   OpenTable
   InitValue = True
   ChkData = True
'   ShowData
   'Modify By Sindy 2012/12/25
   UpdateToolbarState
   blnKeypreview = True
   Call UseDatamaintain(vbKeyHome)
   '2012/12/25 End
   DelFlg = False
   InitValue = False
   OnOffTxt False
   BlnULetter = False
   
   textCUID.BackColor = &H8000000F 'Add By Sindy 2023/2/10
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF9, vbKeyF10, vbKeyPageUp, vbKeyPageDown
         UseDatamaintain (KeyCode)
         KeyCode = 0
      Case vbKeyEnd, vbKeyHome
         If Data_Mission <> 1 And Data_Mission <> 3 Then
            UseDatamaintain (KeyCode)
            KeyCode = 0
         End If
      Case vbKeyEscape
          If MsgBox("是否確定結束?", vbYesNo + vbCritical) = vbYes Then UseDatamaintain (KeyCode)
      Case vbKeyReturn
         UseDatamaintain (vbKeyF9)
         KeyCode = 0
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If BlnULetter Then
        KeyAscii = UpperCase(KeyAscii)
    End If
End Sub

''將Rs資料顯示到畫面上
'Public Sub ShowData()
'   'Add By Sindy 2012/12/25
'   m_CurrKEY(0) = ""
'   m_CurrKEY(1) = ""
'   '2012/12/25 End
'
'   rs.ReQuery
'   RsCounts = rs.RecordCount
'   If RsCounts > 0 Then
'     rs.MoveFirst
'   End If
'   If rs.EOF Then
'      '無資料
'      Call Clear_AllTxtAry(Text1, 0, 10)
'      LblIn10.Caption = ""
'      LblIn01.Caption = ""
'      tlbar.Buttons.Item(1).Enabled = True
'      tlbar.Buttons.Item(2).Enabled = False
'      tlbar.Buttons.Item(3).Enabled = False
'      tlbar.Buttons.Item(4).Enabled = False
'      tlbar.Buttons.Item(6).Enabled = False
'      tlbar.Buttons.Item(7).Enabled = False
'      tlbar.Buttons.Item(8).Enabled = False
'      tlbar.Buttons.Item(9).Enabled = False
'      tlbar.Buttons.Item(11).Enabled = False
'      tlbar.Buttons.Item(12).Enabled = False
'      tlbar.Buttons.Item(14).Enabled = True
'      Exit Sub
'   End If
'   '有資料
'   tlbar.Buttons.Item(1).Enabled = True
'   tlbar.Buttons.Item(2).Enabled = True
'   tlbar.Buttons.Item(3).Enabled = True
'   tlbar.Buttons.Item(4).Enabled = True
'   tlbar.Buttons.Item(6).Enabled = True
'   tlbar.Buttons.Item(7).Enabled = True
'   tlbar.Buttons.Item(8).Enabled = True
'   tlbar.Buttons.Item(9).Enabled = True
'   tlbar.Buttons.Item(11).Enabled = False
'   tlbar.Buttons.Item(12).Enabled = False
'   tlbar.Buttons.Item(14).Enabled = True
'   If RsCounts > 1 And Not InitValue Then
'       If DelFlg Then
'           QueryData "in01=" + CNULL(Fld3), 2
'       Else
'           QueryData "in01=" + CNULL(Fld1), 2
'       End If
'       Exit Sub
'   Else
'       ShowDetail
'       Exit Sub
'   End If
'End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case Data_Mission
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
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
      Case 1, 2, 3, 4: '維護
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
   'Added by Morgan 2020/2/19
   If TypeName(fmCallForm) = "frm040104_3" Then
      fmCallForm.m_bClickInventor = True
      fmCallForm.Enabled = True
      fmCallForm.ZOrder
   End If
   'end 2020/2/19
   
   Set frm050709 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   Text1(Index).SelStart = 0
   Text1(Index).SelLength = Len(Text1(Index).Text)
   Select Case Index
   Case 0, 1
      BlnULetter = True
   Case 2, 4, 5, 7, 8, 11
      '2008/9/2 modify by sonia
      'Text1(Index).IMEMode = 1
      OpenIme
   '2010/2/12 add by sonia
   Case 3, 6
      BlnULetter = False
      CloseIme
   '2010/2/12 end
   End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As ReturnInteger)
   If KeyAscii = 13 And Data_Mission = 4 Then UseDatamaintain vbKeyF9
   If Index = 10 And (Data_Mission = 1 Or Data_Mission = 3) Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strTmp As String
  
   If Data_Mission = 0 Then Exit Sub 'Add by Morgan 2009/6/29
 
   Select Case Index
      Case 0 '申請人編號
         BlnULetter = False
         'Add By Cheng 2002/01/09
         If Data_Mission = 1 And Text1(0).Text <> "" Then
            strTmp = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
            'Modify by Morgan 2011/3/25 原程式段改寫成共用函數
            Text1(1) = PUB_GetNewIN02(strTmp)
         End If
      
      Case 1 '發明人代號
         BlnULetter = False
         If Data_Mission = 1 And Text1(0).Text <> "" And Text1(1) <> "" Then
            strTmp = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
            strExc(0) = "SELECT COUNT(*) FROM INVENTOR WHERE IN01='" & strTmp & "' AND IN02='" & Text1(1).Text & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If RsTemp.Fields(0) > 0 Then
               MsgBox "資料已存在，請重新輸入 !", vbCritical
               Text1(0).SetFocus
            End If
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim i As Integer, strTmp As String
'Memo by Amy 2014/10/23 下列判斷未於錯誤跳離,導致連續按enter也會存入ex:同一申請人有輸ID,判斷ID不可重覆

   If Data_Mission = 0 Then Exit Sub 'Add by Morgan 2009/6/29
 
   Select Case Index
      Case 0
          If Text1(0) = "" Then Me.LblIn01.Caption = "": Exit Sub
          If Left(Text1(Index), 1) <> "X" Then ShowMsg MsgText(1101): Cancel = True: Exit Sub
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.GETCUSTOMER(Text1(0).Text, strName) Then
          If ClsPDGetCustomer(Text1(0).Text, strName) Then
             Text1(Index) = IIf(Right(Text1(Index), 2) = "00", Left(Text1(Index), 6), Text1(Index))
             Me.LblIn01.Caption = strName
          Else
             Me.LblIn01.Caption = ""
             Cancel = True
             Exit Sub
          End If
      Case 9
          If Text1(9) = "" Then Exit Sub
        'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.GetNation(Text1(9), strName) Then
         '2008/9/4 add by sonia 國籍不可輸入001~008
         If Val(Text1(9)) >= 1 And Val(Text1(9)) <= 8 Then
            ShowMsg "發明人國籍不可輸入 001 - 008"
            Cancel = True
            Exit Sub
         Else
         '2008/9/4 end
            If ClsPDGetNation(Text1(9), strName) Then
               Me.LblIn10.Caption = strName
            Else
               Me.LblIn10.Caption = ""
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 4, 5, 7, 8, 11
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(Index).IMEMode = 2
         CloseIme
      Case 10 'ID
         'Memo by Morgan 2009/6/29 原本檢查移到存檔前
         'Add by Amy 2014/04/24 +同一申請人有輸ID,判斷ID不可重覆
         If Trim(Text1(10)) <> "" Then
            If ChkIDExist(Text1(0)) = True Then
                Cancel = True
                Exit Sub
            End If
         End If
   End Select
   If Text1(Index).Text <> "" Then
      'Modified by Morgan 2014/6/13 長度檢查不要固定，改抓MaxLength設定
      'Select Case Index
      '   Case 2, 4, 8
      '      Cancel = Not CheckLengthIsOK(Text1(Index).Text, 40)
      '   '2005/6/22 MODIFY BY SONIA  英文名稱由 30 改為 70
      '   'Case 3, 11
      '   Case 11
      '      Cancel = Not CheckLengthIsOK(Text1(Index).Text, 30)
      '   Case 3, 5, 7
      '      Cancel = Not CheckLengthIsOK(Text1(Index).Text, 70)
      '   Case 6
      '      Cancel = Not CheckLengthIsOK(Text1(Index).Text, 150)
      'End Select
      If Text1(Index).MaxLength > 0 Then
         Cancel = Not CheckLengthIsOK(Text1(Index).Text, Text1(Index).MaxLength)
      End If
      'end 2014/6/13
   End If
   If Cancel = True Then TextInverse Text1(Index)
End Sub

'新增資料
Private Sub insertdata()
   in0(0) = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
   'Modified by Morgan 2014/8/12
   '考慮造字問題所有中日文欄位右邊空白都不去除
   in0(1) = Trim$(Text1(1).Text)
   in0(2) = Trim$(Text1(10).Text)
   in0(3) = LTrim$(Text1(2).Text)
   in0(4) = Trim$(Text1(3).Text)
   in0(5) = LTrim$(Text1(4).Text)
   in0(6) = LTrim$(Text1(5).Text)
   in0(7) = Trim$(Text1(6).Text)
   in0(8) = LTrim$(Text1(7).Text)
   in0(9) = LTrim$(Text1(8).Text)
   in0(10) = Trim$(Text1(9).Text)
   in0(11) = LTrim(Text1(11))
   'edit by nickc 2007/02/06 不用 dll 了
   'Set obj0701 = CreateObject("prjTaiedll.class0701")
   'nRet = obj0701.AddData0709(in0)
   'Set obj0701 = Nothing
   nRet = Cls0701AddData0709(in0)
   'Add By Sindy 2012/12/25
   blnKeypreview = True
   m_CurrKEY(0) = Trim(Text1(0).Text) & String(8 - Len(Trim(Text1(0).Text)), "0")
   m_CurrKEY(1) = String(2 - Len(Trim(Text1(1).Text)), "0") & Trim(Text1(1).Text)
   '2012/12/25 End
End Sub

'Modify By Sindy 2011/3/24 原為sub函數改為function函數回傳boolean
Private Function DeleteData() As Boolean
   in0(0) = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
   in0(1) = Trim$(Text1(1).Text)
   
   DeleteData = False
   'Add By Sindy 2011/3/24 檢查專利基本檔裡是否有此發明人, 若有不可刪除
   'Modify By Sindy 2014/11/6
   'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
   strSql = "SELECT count(*) FROM patentInventor WHERE pi06='" & in0(0) & in0(1) & "'"
   '2014/11/6 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         ShowMsg "此發明人尚有案件資料，不可刪除！"
         Exit Function
      End If
   End If
   '2011/3/24 End
   
   'edit by nickc 2007/02/06 不用 dll 了
   'Set obj0701 = CreateObject("prjTaiedll.class0701")
   'nRet = obj0701.EraseData0709(in0)
   'Set obj0701 = Nothing
   nRet = Cls0701EraseData0709(in0)
   DeleteData = True
   
   'Add By Sindy 2012/12/25
   blnKeypreview = True
   Call UseDatamaintain(vbKeyPageDown)
   '2012/12/25 End
End Function

Private Sub UpdateData()
   in0(0) = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
   'Modified by Morgan 2014/8/12
   '考慮造字問題所有中日文欄位右邊空白都不去除
   in0(1) = Trim$(Text1(1).Text)
   in0(2) = Trim$(Text1(10).Text)
   in0(3) = LTrim$(Text1(2).Text)
   in0(4) = Trim$(Text1(3).Text)
   in0(5) = LTrim$(Text1(4).Text)
   in0(6) = LTrim$(Text1(5).Text)
   in0(7) = Trim$(Text1(6).Text)
   in0(8) = LTrim$(Text1(7).Text)
   in0(9) = LTrim$(Text1(8).Text)
   in0(10) = Trim$(Text1(9).Text)
   in0(11) = LTrim(Text1(11))
   'edit by nickc 2007/02/06 不用 dll 了
   'Set obj0701 = CreateObject("prjTaiedll.class0701")
   'nRet = obj0701.ModifyData0709(in0)
   'Set obj0701 = Nothing
   nRet = Cls0701ModifyData0709(in0)
   'Add By Sindy 2012/12/25
   blnKeypreview = True
   '2012/12/25 End
End Sub

Private Sub ShowDetail()
   'Add By Sindy 2012/12/25
   Call Clear_AllTxt
'   If rs.State <> adStateClosed Then rs.Close
   strSql = "select * from inventor where in01='" & m_CurrKEY(0) & "' and in02='" & m_CurrKEY(1) & "'"
   Set Rs = ClsPDReadRst(strSql, True)
   If Rs.RecordCount > 0 Then
   '2012/12/25 End
      Text1(0) = IIf(IsNull(Rs(0)), "", Rs(0))
      Text1(0) = IIf(Right(Text1(0), 2) = "00", Left(Text1(0), 6), Text1(0))
      Text1(1).Text = IIf(IsNull(Rs(1)), "", Rs(1))
      Text1(2).Text = IIf(IsNull(Rs(3)), "", Rs(3))
      Text1(3).Text = IIf(IsNull(Rs(4)), "", Rs(4))
      Text1(4).Text = IIf(IsNull(Rs(5)), "", Rs(5))
      Text1(5).Text = IIf(IsNull(Rs(6)), "", Rs(6))
      Text1(6).Text = IIf(IsNull(Rs(7)), "", Rs(7))
      Text1(7).Text = IIf(IsNull(Rs(8)), "", Rs(8))
      Text1(8).Text = IIf(IsNull(Rs(9)), "", Rs(9))
      Text1(9).Text = IIf(IsNull(Rs(10)), "", Rs(10))
      Text1(10).Text = IIf(IsNull(Rs(2)), "", Rs(2))
      Text1(11).Text = IIf(IsNull(Rs("in12")), "", Rs("in12"))
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GETCUSTOMER(Text1(0), strName) Then
      If ClsPDGetCustomer(Text1(0), strName) Then
         Me.LblIn01.Caption = strName
      Else
         Me.LblIn01.Caption = ""
      End If
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNation(Text1(9), strName) Then
      If ClsPDGetNation(Text1(9), strName) Then
         Me.LblIn10.Caption = strName
      Else
         Me.LblIn10.Caption = ""
      End If
      
      'Add By Sindy 2023/2/10
      ' 更新CUID
      UpdateCUID Rs
      '2023/2/10 END
   'Add By Sindy 2012/12/25
   End If
   Rs.Close
   Set Rs = Nothing
   '2012/12/25 End
End Sub

'Add By Sindy 2023/2/10
' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("in13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("in13")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("in13"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("in14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("in14")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("in14"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("in15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("in15")) = False Then
         strTemp = rsSrcTmp.Fields("in15")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("in16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("in16")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("in16"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("in17")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("in17")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("in17"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("in18")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("in18")) = False Then
         strTemp = rsSrcTmp.Fields("in18")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              strCDate & " " & _
              strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              strUDate & " " & _
              strUTime
End Sub

'Modify By Sindy 2012/12/25
Private Sub QueryData()
   m_CurrKEY(0) = Trim(Text1(0).Text) & String(8 - Len(Trim(Text1(0).Text)), "0")
   m_CurrKEY(1) = String(2 - Len(Trim(Text1(1).Text)), "0") & Trim(Text1(1).Text)
   strSql = "select * from inventor where in01='" & m_CurrKEY(0) & "' and in02='" & m_CurrKEY(1) & "'"
   Set Rs = ClsPDReadRst(strSql, True)
   If Rs.RecordCount <= 0 Then
      ShowMsg MsgText(9007)
      blnKeypreview = True
      Call UseDatamaintain(vbKeyHome)
   End If
   ShowDetail
End Sub
'Private Sub QueryData(StrCriteria As String, i As Integer)
'On Error Resume Next
'
'   rs.Find StrCriteria
'   i = i - 1
'   If rs.EOF Then
'      ShowMsg MsgText(9007)
'      'Add By Cheng 2002/01/09
'      rs.MoveFirst
'      ShowDetail
'   Else
'      If i <> 0 Then
'        If DelFlg Then
'           QueryData "in02=" + CNULL(Fld4), i
'        Else
'           QueryData "in02=" + CNULL(Fld2), i
'        End If
'      Else
'        'Add By Sindy 2011/3/24
'        If DelFlg = False Then
'            If rs.Fields("in01") <> (Text1(0).Text & String(8 - Len(Text1(0).Text), "0")) Then
'                ShowMsg MsgText(9007)
'                rs.MoveFirst
'                ShowDetail
'                Exit Sub
'            End If
'        End If
'        '2011/3/24 End
'        ShowDetail
'        OnOff_Button tlbar, True
'      End If
'   End If
'End Sub

'Public Sub OpenTable()
'   strSql = "select * from inventor order by in01,in02"
'   'edit by nickc 2007/02/02 不用 dll 了
'   'Set rs = objPublicData.ReadRst(strSQL, True)
'   Set rs = ClsPDReadRst(strSql, True)
'End Sub

Private Sub OnOffTxt(OnOffValue As Boolean)
Dim i As Integer
    
    For i = 0 To 11
        Text1(i).Locked = Not OnOffValue
    Next
End Sub

Private Function ChkInData() As Boolean
Dim strA As String
   
   If Text1(0) <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If Not objPublicData.GETCUSTOMER(Text1(0), strName) Then
        If Not ClsPDGetCustomer(Text1(0), strName) Then
           Text1(0).SetFocus
           ChkInData = False
           Exit Function
        End If
   Else
        ShowMsg "申請人編號不可空白 !"
        Text1(0).SetFocus
        ChkInData = False
        Exit Function
   End If
   '********************************************************************
   If Text1(1) = "" Then
        ShowMsg "發明人代號不可空白 !"
        Text1(1).SetFocus
        ChkInData = False
        Exit Function
   End If
   '********************************************************************
   If Text1(2).Text = "" And Text1(3).Text = "" And Text1(4).Text = "" Then
          ShowMsg "發明人名稱不可同時空白 !"
          Text1(2).SetFocus
          ChkInData = False
          Exit Function
    End If
   '********************************************************************
   If Text1(9) <> "" Then
       'edit by nickc 2007/02/02 不用 dll 了
       'If Not objPublicData.GetNation(Text1(9), strA) Then
       If Not ClsPDGetNation(Text1(9), strA) Then
          Text1(9).SetFocus
          ChkInData = False
          Exit Function
       End If
    Else
       ShowMsg "國籍不可空白 !"
       Text1(9).SetFocus
       ChkInData = False
       Exit Function
    End If
    ChkInData = True
End Function

Private Sub Clear_AllTxt()
Dim i As Integer
   
   For i = 0 To Text1.UBound
      Text1(i) = ""
   Next
   Me.LblIn01.Caption = ""
   Me.LblIn10.Caption = ""
   
   textCUID = "" 'Add By Sindy 2023/2/10
End Sub

Public Sub UseDatamaintain(j As Integer)
Dim i As Integer
Dim MsgAns As Boolean
Dim Rss As ADODB.Recordset
Dim bCancel As Boolean 'Add by Amy 2014/04/24

On Error GoTo HndErr
   'Add By Cheng 2002/01/09
   Me.cmdok(0).Enabled = False
   
'    nRet = True
'   If Data_Mission <> 2 And Not GetNowData Then
'      Fld1 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
'      Fld2 = Trim$(Text1(1).Text)
'   End If
'   If Data_Mission = 2 And GetNowData Then
'      Fld1 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
'      Fld2 = Trim$(Text1(1).Text)
'      If RsCounts > 1 Then
'          rs.MoveNext
'          If rs.EOF Then rs.MoveFirst
'          Fld3 = rs(0)
'          Fld4 = rs(1)
'      End If
'   End If
   Select Case j
      Case vbKeyF2
         If blnKeypreview Then
         Data_Mission = 1  '新增
         Me.Caption = Me.Tag & "-新增" 'Added by Morgan 2020/2/19
         GetNowData = True
         ChkData = False
         Call Clear_AllTxt
         Me.LblIn01.Caption = ""
         Me.LblIn10.Caption = ""
         ChkData = True
         Call OnOff_Button(tlbar, False)
         OnOffTxt True
         blnKeypreview = False
         Text1(0).SetFocus
         
         'Add By Cheng 2002/01/09
         Me.cmdok(0).Enabled = True
         
         End If
      Case vbKeyF3
         If blnKeypreview Then
         Data_Mission = 3  '修改
         Me.Caption = Me.Tag & "-修改" 'Added by Morgan 2020/2/19
         GetNowData = True
         Call OnOff_Button(tlbar, False)
         OnOffTxt True
         blnKeypreview = False
         Text1(0).Locked = True
         Text1(1).Locked = True
         Text1(10).SetFocus
         End If
      Case vbKeyF5
         If blnKeypreview Then
         Data_Mission = 2  '刪除
         GetNowData = True
         Call OnOff_Button(tlbar, False)
         OnOffTxt False
         blnKeypreview = False
         Text1(0).SetFocus
         End If
      Case vbKeyF4
         If blnKeypreview Then
         Data_Mission = 4
         Me.Caption = Me.Tag & "-查詢" 'Added by Morgan 2020/2/19
         GetNowData = True
         ChkData = False
         Call OnOff_Button(tlbar, False)
         OnOffTxt False
         Text1(0).Locked = False
         Text1(1).Locked = False
         Call Clear_AllTxt
         Me.LblIn01.Caption = ""
         Me.LblIn10.Caption = ""
         blnKeypreview = False
         ChkData = True
         Text1(0).SetFocus
         End If
      Case vbKeyHome
         If blnKeypreview Then
            'rs.MoveFirst
            'Add By Sindy 2012/12/25
'            If rs.State <> adStateClosed Then rs.Close
            strSql = "select in01,min(in02) from inventor where in01=(select min(in01) from inventor) " & _
                     "group by in01"
            Set Rs = ClsPDReadRst(strSql, True)
            If Rs.RecordCount > 0 Then
               m_CurrKEY(0) = Rs.Fields(0)
               m_CurrKEY(1) = Rs.Fields(1)
            Else
               m_CurrKEY(0) = ""
               m_CurrKEY(1) = ""
            End If
            Rs.Close
            Set Rs = Nothing
            '2012/12/25 End
            Call ShowDetail
'            Text1(0).SetFocus
         End If
      
      Case vbKeyPageUp
         If blnKeypreview Then
'            rs.MovePrevious
'            If rs.BOF Then
'               rs.MoveFirst
'               ShowMsg MsgText(9008)
'            End If
            'Add By Sindy 2012/12/25
'            If rs.State <> adStateClosed Then rs.Close
            strSql = "select count(*) from inventor where in01||in02<'" & Trim(Text1(0)) & String(8 - Len(Trim(Text1(0))), "0") & Trim(Text1(1)) & "'"
            Set Rs = ClsPDReadRst(strSql, True)
            If Rs.RecordCount > 0 And Rs.Fields(0) > 0 Then
               strSql = "select max(in01||in02) from inventor where in01||in02<'" & Trim(Text1(0)) & String(8 - Len(Trim(Text1(0))), "0") & Trim(Text1(1)) & "'"
               Set Rs = ClsPDReadRst(strSql, True)
               If Rs.RecordCount > 0 Then
                  m_CurrKEY(0) = Left(Trim(Rs.Fields(0)), 8)
                  m_CurrKEY(1) = Mid(Trim(Rs.Fields(0)), 9, 2)
               End If
            Else
               ShowMsg MsgText(9008)
               Call UseDatamaintain(vbKeyHome)
               Exit Sub
            End If
            Rs.Close
            Set Rs = Nothing
            '2012/12/25 End
            Call ShowDetail
'            Text1(0).SetFocus
         End If
      Case vbKeyPageDown
         If blnKeypreview Then
'            rs.MoveNext
'            If rs.EOF Then
'               rs.MoveLast
'               ShowMsg MsgText(9009)
'            End If
            'Add By Sindy 2012/12/25
'            If rs.State <> adStateClosed Then rs.Close
            strSql = "select count(*) from inventor where in01||in02>'" & Trim(Text1(0)) & String(8 - Len(Trim(Text1(0))), "0") & Trim(Text1(1)) & "'"
            Set Rs = ClsPDReadRst(strSql, True)
            If Rs.RecordCount > 0 And Rs.Fields(0) > 0 Then
               strSql = "select min(in01||in02) from inventor where in01||in02>'" & Trim(Text1(0)) & String(8 - Len(Trim(Text1(0))), "0") & Trim(Text1(1)) & "'"
               Set Rs = ClsPDReadRst(strSql, True)
               If Rs.RecordCount > 0 Then
                  m_CurrKEY(0) = Left(Trim(Rs.Fields(0)), 8)
                  m_CurrKEY(1) = Mid(Trim(Rs.Fields(0)), 9, 2)
               End If
            Else
               ShowMsg MsgText(9009)
               Call UseDatamaintain(vbKeyEnd)
               Exit Sub
            End If
            Rs.Close
            Set Rs = Nothing
            '2012/12/25 End
            Call ShowDetail
'            Text1(0).SetFocus
         End If
      Case vbKeyEnd
         If blnKeypreview Then
            'rs.MoveLast
            'Add By Sindy 2012/12/25
'            If rs.State <> adStateClosed Then rs.Close
            strSql = "select in01,max(in02) from inventor where in01=(select max(in01) from inventor) " & _
                     "group by in01"
            Set Rs = ClsPDReadRst(strSql, True)
            If Rs.RecordCount > 0 Then
               m_CurrKEY(0) = Rs.Fields(0)
               m_CurrKEY(1) = Rs.Fields(1)
            Else
               m_CurrKEY(0) = ""
               m_CurrKEY(1) = ""
            End If
            Rs.Close
            Set Rs = Nothing
            '2012/12/25 End
            Call ShowDetail
'            Text1(0).SetFocus
         End If

      Case vbKeyF9
         If Not blnKeypreview Then
            Select Case Data_Mission
               Case 1
                  
                  If ChkInData Then
                     'Add By Cheng 2002/05/22
                     '重新檢查欄位有效性
                     If TxtValidate = False Then Exit Sub
                     
                     strNo = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
                     strSql = "select * from inventor where in01=" + CNULL(strNo) + " and in02=" + CNULL(Trim$(Text1(1).Text))
                       'edit by nickc 2007/02/02 不用 dll 了
                       'Set Rss = objPublicData.ReadRst(strSQL)
                       Set Rss = ClsPDReadRst(strSql)
                       If Not Rss.EOF Then
                          MsgBox "編號:" + Trim$(Text1(0)) + "-" + Trim$(Text1(1)) + "的資料已存在", vbCritical
                          Rss.Close
                          OnOffTxt True
                          Text1(0).SetFocus
                          Exit Sub
                       End If
                       Rss.Close
                       'Modify by Amy 2014/04/24 +同一申請人有輸ID,判斷ID不可重覆
                       bCancel = False
                       Call Text1_Validate(10, bCancel)
                        If bCancel = True Then
                            OnOffTxt True
                            Text1(10).SetFocus
                            Exit Sub
                       End If
                       'end 2014/04/24
                       
                       '---2008/10/03 ADD BY TONI
                       'Modify by Morgan 2008/12/23 語法修正
                       'Modified by Morgan 2014/8/12 考慮造字問題改用SQL語法去掉空白比對
                       'strSql = "select * from inventor where in01=" & CNULL(strNo) & " and (in04='" + ChgSQL(Trim$(Text1(2).Text)) & "'" & _
                                " OR in05='" & ChgSQL(Trim$(Text1(3).Text)) & "' OR in06='" & ChgSQL(Trim$(Text1(4).Text)) & "')"
                       strSql = "select * from inventor where in01=" & CNULL(strNo) & " and (rtrim(in04)=rtrim('" + ChgSQL(LTrim$(Text1(2).Text)) & "')" & _
                                " OR in05='" & ChgSQL(Trim$(Text1(3).Text)) & "' OR rtrim(in06)=rtrim('" & ChgSQL(LTrim$(Text1(4).Text)) & "'))"
                       Set Rss = ClsPDReadRst(strSql)
                       If Not Rss.EOF Then
                           If Trim(Text1(2).Text) = Trim(Rss.Fields("in04")) Then
                              If MsgBox("發明人名稱中文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                 Rss.Close
                                 Text1(2).SetFocus
                                 OnOffTxt True
                                 Exit Sub
                              End If
                           End If
                           
                           If Trim(Text1(3).Text) = Trim(Rss.Fields("in05")) Then
                              If MsgBox("發明人名稱英文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                 Rss.Close
                                 Text1(3).SetFocus
                                 OnOffTxt True
                                 Exit Sub
                              End If
                           End If
                           
                           If Trim(Text1(4).Text) = Trim(Rss.Fields("in05")) Then
                              If MsgBox("發明人名稱日文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                 Rss.Close
                                 Text1(2).SetFocus
                                 OnOffTxt True
                                 Exit Sub
                              End If
                           End If
                       End If
                       '---end
                      insertdata
                      DelFlg = False
                  Else
                      Exit Sub
                  End If
'                  Fld1 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
'                  Fld2 = Trim$(Text1(1).Text)
                  'Call ShowData
                  Call ShowDetail
               Case 2
                  If DelMsg Then
                     'Modify By Sindy 2011/3/24
                     'DeleteData
                     If DeleteData = False Then
                        DelFlg = False
                     '2011/3/24 End
                     Else
                        DelFlg = True
                     End If
                  End If
                  'Call ShowData
                  Call ShowDetail
                  DelFlg = False
               
               Case 3
                  
                  If ChkInData Then
                     'Add By Cheng 2002/05/22
                     '重新檢查欄位有效性
                     If TxtValidate = False Then Exit Sub
                     'Modify by Amy 2014/04/24 +同一申請人有輸ID,判斷ID不可重覆
                     bCancel = False
                     Call Text1_Validate(10, bCancel)
                     If bCancel = True Then
                        OnOffTxt True
                        Text1(10).SetFocus
                        Exit Sub
                     End If
                     'end 2014/04/24
                    
                     Call UpdateData
                     'Call ShowData
                     Call ShowDetail
                  Else
                     Exit Sub
                  End If
                  DelFlg = False
               
               Case 4
                  If Text1(0).Text <> "" Then
                     'edit by nickc 2007/02/02 不用 dll 了
                     'If Not objPublicData.GetCustomer(Text1(0), strName) Then
                     If Not ClsPDGetCustomer(Text1(0), strName) Then
                        Text1(0).SetFocus
                        Exit Sub
                     End If
                  Else
                       ShowMsg MsgText(9015)
                       Text1(0).SetFocus
                       Exit Sub
                  End If
                     '********************************************************************
                  If Text1(1).Text = "" Then
                       ShowMsg MsgText(9015)
                       Text1(1).SetFocus
                       Exit Sub
                  End If
'                  Fld3 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
'                  Fld2 = Trim$(Text1(1).Text)
'                  rs.MoveFirst
'                  QueryData "in01=" + CNULL(Fld3), 2
                  Call QueryData 'Modify By Sindy 2012/12/25
            End Select
            
            Data_Mission = 0
            Call OnOff_Button(tlbar, True)
            OnOffTxt False
            GetNowData = False
            blnKeypreview = True
            Text1(0).SetFocus
            Me.Caption = Me.Tag   'Added by Morgan 2020/2/19
            If bAddOnly Then Unload Me 'Added by Morgan 2020/2/19
         End If
         
      Case vbKeyF10
         If Not blnKeypreview Then
            
            If Data_Mission <> 4 Then
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Sub
            End If
            
            Call OnOff_Button(tlbar, True)
            DelFlg = False
            Data_Mission = 0
            OnOffTxt False
            'Call ShowData
            Call ShowDetail
            GetNowData = False
            blnKeypreview = True
            Text1(0).SetFocus
            Me.Caption = Me.Tag 'Added by Morgan 2020/2/19
            If bAddOnly Then Unload Me 'Added by Morgan 2020/2/19
         End If
      
      Case vbKeyEscape
         Unload Me
   End Select
   Exit Sub
HndErr:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
          UseDatamaintain (vbKeyF2)
      Case 2
          UseDatamaintain (vbKeyF3)
      Case 3
          UseDatamaintain vbKeyF5
          UseDatamaintain vbKeyF9
      Case 4
          UseDatamaintain (vbKeyF4)
      Case 6
          UseDatamaintain (vbKeyHome)
      Case 7
          UseDatamaintain (vbKeyPageUp)
      Case 8
          UseDatamaintain (vbKeyPageDown)
      Case 9
          UseDatamaintain (vbKeyEnd)
      Case 11
          UseDatamaintain (vbKeyF9)
      Case 12
          UseDatamaintain (vbKeyF10)
      Case 14
          UseDatamaintain (vbKeyEscape)
   End Select
End Sub

'Add By Cheng 2002/01/09
Private Sub ShowCustData()
Dim i As Integer
Dim rsThis As New ADODB.Recordset

   If rsThis.State <> adStateClosed Then rsThis.Close
   Set rsThis = Nothing
   rsThis.CursorLocation = adUseClient
   strSql = "Select CU01 AS 申請人編號,CU11 AS ID,CU04 As 中文名稱,CU05||' '||CU88||' '||CU89||' '||CU90 AS 英文名稱,CU06 AS 日文名稱, " & _
            " CU23 AS 中文地址,CU24||' '||CU25||' '||CU26||' '||CU27||' '||CU28 AS 英文地址,CU29 AS 日文地址,CU10 AS 國籍代號,NVL(NA03,NVL(NA04,NULL)) AS 國籍名稱, CU07 AS 代表人 " & _
            " From Customer,Nation Where CU10=NA01(+) AND CU01 = '" & Left(Me.Text1(0).Text & "00000000", 8) & "'"
   rsThis.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsThis.EOF = False Then
      For i = 0 To Text1.UBound
         Text1(i) = ""
      Next
      Me.LblIn01.Caption = ""
      Me.LblIn10.Caption = ""
         
      If Not IsNull(rsThis("申請人編號")) Then
         Me.Text1(0).Text = rsThis("申請人編號")
         Text1_Validate 0, False
         Text1_LostFocus 0
      End If
      If Not IsNull(rsThis("ID")) Then
         Me.Text1(10).Text = rsThis("ID")
      End If
      If Not IsNull(rsThis("中文名稱")) Then
         Me.Text1(2).Text = rsThis("中文名稱")
      End If
      If Not IsNull(rsThis("英文名稱")) Then
         Me.Text1(3).Text = rsThis("英文名稱")
      End If
      If Not IsNull(rsThis("日文名稱")) Then
         Me.Text1(4).Text = rsThis("日文名稱")
      End If
      If Not IsNull(rsThis("中文地址")) Then
         Me.Text1(5).Text = rsThis("中文地址")
      End If
      If Not IsNull(rsThis("英文地址")) Then
         Me.Text1(6).Text = rsThis("英文地址")
      End If
      If Not IsNull(rsThis("日文地址")) Then
         Me.Text1(7).Text = rsThis("日文地址")
      End If
      If Not IsNull(rsThis("國籍代號")) Then
         Me.Text1(9).Text = rsThis("國籍代號")
      End If
      If Not IsNull(rsThis("國籍名稱")) Then
         Me.LblIn10.Caption = rsThis("國籍名稱")
      End If
      If Not IsNull(rsThis("代表人")) Then
         Me.Text1(8).Text = rsThis("代表人")
      End If
      
   End If
   If rsThis.State <> adStateClosed Then rsThis.Close
   Set rsThis = Nothing
End Sub
   
'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Text1
      If objTxt.Enabled = True Then
         Cancel = False
         Text1_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   'Add by Morgan 2009/6/29 從 Validate 事件移來以避免無法跳離
   If Text1(10) <> "" Then
      If CheckID(0, Text1(10).Text) = False Then
         strExc(0) = "身分證字號錯誤，是否確定 !"
         If MsgBox(strExc(0), vbYesNo + vbCritical) = vbNo Then
            Exit Function
         End If
      End If
   End If
            
   '2011/10/17 add by sonia
   If Text1(5) <> "" Then
      If CheckTaiwanAddr(Text1(5), Text1(9), "發明人地址(中)") = False Then
         Text1(5).SetFocus
         Text1_GotFocus (5)
         Exit Function
      End If
   End If
   '2011/10/17 end
   
   'Add by Sindy 2021/12/07 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If

   TxtValidate = True
End Function

'Add by Amy 2014/04/24 同一申請人ID是否已存在
Private Function ChkIDExist(ByVal strIN01 As String) As Boolean
    Dim stSQL As String
    Dim Rs As ADODB.Recordset
    Dim ii As Integer
    
    strIN01 = Mid(GetNewFagent(strIN01), 1, 8)
    
    stSQL = "Select * From Inventor Where in01=" & CNULL(strIN01) & " And in02<>" & CNULL(Trim$(Text1(1))) & _
               " And in03='" & (Trim$(Text1(10))) & "' "
    intI = 1
    Set Rs = ClsLawReadRstMsg(intI, stSQL)
    If intI = 1 Then
        MsgBox "ID:" + Trim$(Text1(10)) & " 與發明人代號 " & Rs.Fields("in02") & " 重覆,請確認 !", vbCritical
        ChkIDExist = True
        Rs.Close
        Exit Function
    End If
    Rs.Close
    ChkIDExist = False
End Function
 
'Move by Lydia 2022/09/06 從basQuery搬過來; 並且把Public改回Private
'Cls0701 nickc 2007/02/05
Private Function Cls0701AddData0709(in0() As String) As Boolean
Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

    strSql = "insert into inventor(in01,in02,in03,in04,in05,in06,in07,in08,in09,in10,in11,in12) values("
    For i = 0 To UBound(in0()) - 1
        'Modify By Cheng 2002/12/11
'        strSQL = strSQL + CNULL(in0(i)) + ","
        strSql = strSql + CNULL(ChgSQL(in0(i))) + ","
    Next
    strSql = Left(strSql, Len(strSql) - 1)
    strSql = strSql + ")"
   'Debug.Print strSQL
On Error GoTo ErrHand
    cnnConnection.BeginTrans
    cnnConnection.Execute strSql
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    Cls0701AddData0709 = True
    Exit Function
ErrHand:
    'Add By Cheng 2002/11/14
    If Err.NUMBER = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    Cls0701AddData0709 = False
    'edit by nickc 2007/02/05
'ErrorLog
MsgBox Err.Description
End Function

'Move by Lydia 2022/09/06 從basQuery搬過來; 並且把Public改回Private
Private Function Cls0701EraseData0709(in0() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSql = "delete from inventor where in01='" + in0(0) + "' and in02='" + in0(1) + "' "
On Error GoTo ErrHand
    cnnConnection.BeginTrans
    cnnConnection.Execute strSql
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    Cls0701EraseData0709 = True
    Exit Function
ErrHand:
    'Add By Cheng 2002/11/14
    If Err.NUMBER = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    Cls0701EraseData0709 = False
    'edit by nickc 2007/02/05
'ErrorLog
MsgBox Err.Description
End Function

'Move by Lydia 2022/09/06 從basQuery搬過來; 並且把Public改回Private
Private Function Cls0701ModifyData0709(in0() As String) As Integer
Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    strSql = "update inventor set "
    'Modify By Cheng 2002/12/11
'    strSQL = strSQL + "in03=" + CNULL(in0(2)) + ", in04=" + CNULL(in0(3)) + ",in05=" + CNULL(in0(4)) + _
'      ",in06=" + CNULL(in0(5)) + ",in07=" + CNULL(in0(6)) + ",in08=" + CNULL(in0(7)) + ",in09=" + CNULL(in0(8)) + _
'      ",in10=" + CNULL(in0(9)) + ",in11=" + CNULL(in0(10)) + ",in12=" + CNULL(in0(11)) + " where in01='" + in0(0) + "' and in02='" + in0(1) + "'"
    strSql = strSql + "in03=" + CNULL(ChgSQL(in0(2))) + ", in04=" + CNULL(ChgSQL(in0(3))) + ",in05=" + CNULL(ChgSQL(in0(4))) + _
      ",in06=" + CNULL(ChgSQL(in0(5))) + ",in07=" + CNULL(ChgSQL(in0(6))) + ",in08=" + CNULL(ChgSQL(in0(7))) + ",in09=" + CNULL(ChgSQL(in0(8))) + _
      ",in10=" + CNULL(ChgSQL(in0(9))) + ",in11=" + CNULL(ChgSQL(in0(10))) + ",in12=" + CNULL(ChgSQL(in0(11))) + " where in01='" + in0(0) + "' and in02='" + in0(1) + "'"
On Error GoTo ErrHand
    cnnConnection.BeginTrans
    cnnConnection.Execute strSql
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    Cls0701ModifyData0709 = True
    Exit Function
ErrHand:
    'Add By Cheng 2002/11/14
    If Err.NUMBER = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    Cls0701ModifyData0709 = False
    'edit by nickc 2007/02/05
'ErrorLog
MsgBox Err.Description
End Function

