VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010016 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "櫃檯每日信件輸入"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8952
   Begin VB.CommandButton Command3 
      Caption         =   "查詢國外信件清單數量"
      Height          =   285
      Left            =   4680
      TabIndex        =   13
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "取消列印次數"
      Height          =   285
      Left            =   7410
      TabIndex        =   30
      Top             =   1650
      Width           =   1395
   End
   Begin VB.TextBox textLI07 
      Height          =   300
      Left            =   4350
      MaxLength       =   12
      TabIndex        =   14
      Top             =   1650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   8400
      MaxLength       =   1
      TabIndex        =   12
      Top             =   1320
      Width           =   285
   End
   Begin VB.CommandButton Command2 
      Caption         =   "國外信件列印"
      Height          =   285
      Left            =   6060
      TabIndex        =   11
      Top             =   1350
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "非國外信件列印"
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Top             =   1350
      Width           =   1455
   End
   Begin VB.ComboBox textLI06 
      Height          =   300
      ItemData        =   "frm010016.frx":0000
      Left            =   6870
      List            =   "frm010016.frx":0002
      TabIndex        =   5
      Top             =   990
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   6120
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   660
      Width           =   2340
   End
   Begin VB.TextBox textLI02 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   660
      Width           =   765
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3705
      Left            =   30
      TabIndex        =   8
      Top             =   2010
      Width           =   8895
      _ExtentX        =   15706
      _ExtentY        =   6519
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "大陸-一般信件"
      TabPicture(0)   =   "frm010016.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grd1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "非大陸-一般信件"
      TabPicture(1)   =   "frm010016.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "國外信件"
      TabPicture(2)   =   "frm010016.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grd1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "客戶信件"
      TabPicture(3)   =   "frm010016.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grd1(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "退件"
      TabPicture(4)   =   "frm010016.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grd1(4)"
      Tab(4).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3285
         Index           =   0
         Left            =   30
         TabIndex        =   15
         Top             =   360
         Width           =   8805
         _ExtentX        =   15515
         _ExtentY        =   5779
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3285
         Index           =   1
         Left            =   -74970
         TabIndex        =   16
         Top             =   360
         Width           =   8805
         _ExtentX        =   15515
         _ExtentY        =   5779
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3285
         Index           =   2
         Left            =   -74970
         TabIndex        =   17
         Top             =   360
         Width           =   8805
         _ExtentX        =   15515
         _ExtentY        =   5779
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3285
         Index           =   3
         Left            =   -74970
         TabIndex        =   18
         Top             =   360
         Width           =   8805
         _ExtentX        =   15515
         _ExtentY        =   5779
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3285
         Index           =   4
         Left            =   -74970
         TabIndex        =   27
         Top             =   360
         Width           =   8805
         _ExtentX        =   15515
         _ExtentY        =   5779
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   1
      End
   End
   Begin VB.TextBox textLI04 
      Height          =   300
      Left            =   4860
      MaxLength       =   12
      TabIndex        =   4
      Top             =   990
      Width           =   1275
   End
   Begin VB.TextBox textLI01 
      Alignment       =   2  '置中對齊
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   930
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   0
      Top             =   660
      Width           =   1035
   End
   Begin VB.ComboBox cboLI08 
      Height          =   300
      ItemData        =   "frm010016.frx":0090
      Left            =   2010
      List            =   "frm010016.frx":00A3
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   660
      Width           =   1785
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":00E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":03FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":08F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":0C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":1248
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":1564
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":1880
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":1B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":1EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010016.frx":21D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   480
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   847
      ButtonWidth     =   1076
      ButtonHeight    =   794
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "列印"
            Key             =   "keyQuery"
            ImageIndex      =   12
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
   Begin MSForms.TextBox textLI09 
      Height          =   300
      Left            =   570
      TabIndex        =   7
      Top             =   1650
      Width           =   3735
      VariousPropertyBits=   -1466941413
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "6588;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox textLI05 
      Height          =   300
      Left            =   570
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "6588;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textLI03 
      Height          =   300
      Left            =   930
      TabIndex        =   3
      Top             =   990
      Width           =   3375
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "5953;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "列印次數："
      Height          =   180
      Left            =   7470
      TabIndex        =   29
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Index           =   2
      Left            =   0
      TabIndex        =   28
      Top             =   1710
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Index           =   1
      Left            =   5370
      TabIndex        =   26
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "事由："
      Height          =   180
      Left            =   0
      TabIndex        =   25
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "序號："
      Height          =   180
      Left            =   3870
      TabIndex        =   24
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收件人："
      Height          =   180
      Left            =   6150
      TabIndex        =   22
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "文號："
      Height          =   180
      Left            =   4350
      TabIndex        =   21
      Top             =   1050
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公司名稱："
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "信件日期："
      Height          =   180
      Left            =   0
      TabIndex        =   19
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frm010016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0已修改 textLI03/textLI05/textLI09/grd1() (Printer列印未改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_EditMode As Integer
Dim m_SubMode As Integer
' 第一筆資料的key
Dim m_FirstKEY(3) As String
' 最後一筆資料的key
Dim m_LastKEY(3) As String
' 目前正在顯示的key
Dim m_CurrKEY(3) As String
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iLine As Integer
Dim MaxLine As Integer
Dim PLeft(7) As Integer
Dim strTemp(7) As String
Dim strDept As String, strDeptName As String, strSignPer As String, strSysDept As String 'Add By Sindy 2013/3/12
Dim m_PrintDate As String 'Add By Sindy 2014/3/4


Private Sub cboLI08_Change()
   SSTab1.Tab = cboLI08.ListIndex
   If SSTab1.Tab = 2 Then
      Label4.Visible = False
      textLI06.Visible = False
   Else
      Label4.Visible = True
      textLI06.Visible = True
   End If
End Sub

Private Sub cboLI08_Click()
   SSTab1.Tab = cboLI08.ListIndex
   If SSTab1.Tab = 2 Then
      Label4.Visible = False
      textLI06.Visible = False
   Else
      Label4.Visible = True
      textLI06.Visible = True
   End If
End Sub

'Add By Sindy 2013/3/12
Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   Me.grd1(SSTab1.Tab).MousePointer = flexArrowHourGlass
   PrintData
   Me.grd1(SSTab1.Tab).MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2013/3/12
Private Sub Command2_Click()
   Screen.MousePointer = vbHourglass
   Me.grd1(SSTab1.Tab).MousePointer = flexArrowHourGlass
   PrintData3
   Me.grd1(SSTab1.Tab).MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2013/3/12
Private Sub Command3_Click()
'開啟國外信件清單數量視窗
frm010016_1.textLI01 = Me.textLI01
If frm010016_1.StrMenu = True Then
   frm010016_1.Show vbModal
Else
   MsgBox "無資料！"
End If
Unload frm010016_1
Set frm010016_1 = Nothing
End Sub

'Add By Sindy 2014/9/3 取消列印次數
Private Sub Command4_Click()
Dim strNum As String
   
On Error GoTo ErrHand
   
   If Val(Text1) = 0 Then
      MsgBox "請輸入列印次數！", vbInformation, "輸入錯誤"
      Text1.SetFocus
      Exit Sub
   End If
   
   m_PrintDate = InputBox("請輸入欲列印的日期：", "國外信件列印", TransDate(m_CurrKEY(0), 1))
   If m_PrintDate = "" Then '取消
      Exit Sub
   Else
      m_PrintDate = DBDATE(m_PrintDate)
   End If
   '檢查是否為日期格式
   If ChkDate(m_PrintDate) = False Then
      Exit Sub
   End If
   If ChkWorkDay(m_PrintDate) = False Then
      MsgBox "請輸入工作天！", vbInformation, "輸入錯誤"
      Exit Sub
   End If
   
   strExc(0) = "SELECT li01 FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3' and li07='" & Text1 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      strExc(0) = "SELECT max(li07) FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3' and li07>0 and li07 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      strNum = ""
      If intI = 1 Then
         strNum = RsTemp.Fields(0)
      End If
      MsgBox "無此列印次數，請重新確認！" & IIf(strNum <> "", "（目前最大列印次數為：" & strNum & "）", ""), vbInformation, "輸入列印次數錯誤"
      Text1.SetFocus
      Exit Sub
   Else
      If MsgBox("確定要清除列印次數嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Sub
      End If
   End If
   
   cnnConnection.BeginTrans
   
   '清除列印次數及清單數量，欲重新列印
   strSql = "delete from ForeignLetterCount where FLC01=" & m_PrintDate & " and FLC03>=" & Text1
   cnnConnection.Execute strSql
   strSql = "update letterinput set li07=NULL WHERE li01=" & m_PrintDate & " and li08='3' and li07>='" & Text1 & "'"
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   Text1 = ""
   MsgBox "資料已清除！", vbOKOnly
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Sub

Private Sub Form_Initialize()
ReDim m_FieldList(8) As FIELDITEM
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
'            PrintData
      ' 刪除
      Case vbKeyF5:
'         If m_bDelete Then
'            If m_EditMode = 0 Then
'               OnAction KeyCode
'               KeyCode = 0
'            End If
'         End If
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
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'Mark by Amy 2022/01/11 改Form2.0 不使用Enter存檔
'Private Sub Form_KeyPress(KeyAscii As Integer)
'   Select Case KeyAscii
'      Case 13:
'         If m_EditMode <> 0 Then
'            KeyAscii = 0
'            OnAction vbKeyF9
'         End If
'   End Select
'End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   m_bInsert = IsUserHasRightOfFunction("frm010016", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm010016", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm010016", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm010016", strFind, False)
   textLI01.Text = strSrvDate(2)
   cboLI08.Text = cboLI08.List(0)
   InitialField
   RefreshRange
   GetAllData
   ShowLastRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   SetGrd
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
       Set Printer = Printers(i)
       Combo1.AddItem Printer.DeviceName, j
        j = j + 1
       If Printer.DeviceName = strSql Then
           SeekPrint = i
       End If
   Next i
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
   
   '2009/12/9 ADD BY SONIA
   textLI05.AddItem "支票 $"
   textLI05.AddItem "文件"
   textLI05.AddItem "繳款書"
   textLI05.AddItem "扣繳憑單"
   textLI05.AddItem "匯款單"
   textLI05.AddItem "快捷"
   textLI05.AddItem "查無此人"
   textLI05.AddItem "遷移"
   textLI05.AddItem "逾期招領"
   textLI05.AddItem "匯款通知" 'Add By Sindy 2013/4/8
   
   textLI06.AddItem "專利處"
   textLI06.AddItem "商標處"
   textLI06.AddItem "臺一投資"   'modify by sonia 2021/2/25 改名稱
   textLI06.AddItem "法務"
   textLI06.AddItem "財務處"
   textLI06.AddItem "謝經理"
   textLI06.AddItem "劉經理"
   'textLI06.AddItem "張副理"    'cancel by sonia 2023/4/19 阿妙
   textLI06.AddItem "薛經理"
   'textLI06.AddItem "葉特助"    'cancel by sonia 2023/4/19 阿妙 'Modify By Sindy 2015/6/2 葉經理
   textLI06.AddItem "服務處"
   'textLI06.AddItem "王副總"    'cancel by sonia 2023/4/19 阿妙說郵件很少,不必預設
   'textLI06.AddItem "游經理" 'Removed by Morgan 2025/2/20 將退休
   textLI06.AddItem "吳副理"     'modify by sonia 2023/4/19 阿妙(原為吳婉莘) 'Add By Sindy 2015/6/2
   'textLI06.AddItem "唐律師"
   '2009/12/9 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Printer = Printers(SeekPrint)
   Printer.Orientation = SeekPrintL
   Set frm010016 = Nothing
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   grd1(0).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1(0).Cols - 1
      grd1(0).row = 0
      grd1(0).col = iRow
      grd1(0).Text = arrGridHeadText(iRow)
      grd1(0).ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1(0).CellAlignment = flexAlignCenterCenter
   Next
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   grd1(1).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1(1).Cols - 1
      grd1(1).row = 0
      grd1(1).col = iRow
      grd1(1).Text = arrGridHeadText(iRow)
      grd1(1).ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1(1).CellAlignment = flexAlignCenterCenter
   Next
   'Modify By Sindy 2011/6/1
'   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "備註")
'   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 2500)
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 0, 2500)
   grd1(2).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1(2).Cols - 1
      grd1(2).row = 0
      grd1(2).col = iRow
      grd1(2).Text = arrGridHeadText(iRow)
      grd1(2).ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1(2).CellAlignment = flexAlignCenterCenter
   Next
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   grd1(3).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1(3).Cols - 1
      grd1(3).row = 0
      grd1(3).col = iRow
      grd1(3).Text = arrGridHeadText(iRow)
      grd1(3).ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1(3).CellAlignment = flexAlignCenterCenter
   Next
   'Add By Sindy 98/03/20
   arrGridHeadText = Array("流水號", "公司名稱", "文號", "事由", "收件人", "備註")
   arrGridHeadWidth = Array(800, 1500, 1500, 2000, 1000, 2500)
   grd1(4).Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1(4).Cols - 1
      grd1(4).row = 0
      grd1(4).col = iRow
      grd1(4).Text = arrGridHeadText(iRow)
      grd1(4).ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1(4).CellAlignment = flexAlignCenterCenter
   Next
   '98/03/20 End
End Sub

' 98/02/05 Modify by Sindy
'Private Sub grd1_SelChange(Index As Integer)
'Dim tmpRow As Integer
'grd1(Index).Visible = False
'tmpRow = grd1(Index).MouseRow
'grd1(Index).col = 0
'If tmpRow <> 0 Then
'    m_CurrKEY(1) = grd1(Index).TextMatrix(tmpRow, 0)
'    m_CurrKEY(2) = cboLI08.ListIndex + 1
'    UpdateCtrlData
'End If
'grd1(Index).Visible = True
'End Sub

Private Sub grd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grd1(Index), x, y, nCol, nRow
   grd1(Index).col = nCol
   grd1(Index).row = nRow
End Sub

Private Sub grd1_SelChange(Index As Integer)
Dim tmpMouseRow
Dim i, j

   grd1(Index).Visible = False
   tmpMouseRow = grd1(Index).row
   grd1(Index).Visible = True
   If tmpMouseRow <> 0 Then
       grd1(Index).row = tmpMouseRow
       grd1(Index).col = 0
       If grd1(Index).CellBackColor <> &HFFC0C0 Then
            grd1(Index).Visible = False
            For j = 1 To grd1(Index).Rows - 1
                grd1(Index).row = j
                For i = 0 To grd1(Index).Cols - 1
                     grd1(Index).col = i
                     grd1(Index).CellBackColor = QBColor(15)
                Next i
           Next j
           grd1(Index).row = tmpMouseRow
            For i = 0 To grd1(Index).Cols - 1
                grd1(Index).col = i
                grd1(Index).CellBackColor = &HFFC0C0
            Next i
            m_CurrKEY(0) = Val(ChangeTStringToWString(textLI01))
            m_CurrKEY(1) = grd1(Index).TextMatrix(tmpMouseRow, 0)
            m_CurrKEY(2) = cboLI08.ListIndex + 1
            UpdateCtrlData
            grd1(Index).Visible = True
       End If
   End If
End Sub
' 98/02/05 END

Private Sub ChgGrdData(Index As Integer, iRow As Integer)

   grd1(Index).Visible = False
Dim i, j, k

   For i = 0 To 3
       For j = 1 To grd1(i).Rows - 1
           grd1(i).row = j
           For k = 0 To grd1(i).Cols - 1
               grd1(i).col = k
               grd1(i).CellBackColor = QBColor(15)
           Next k
       Next j
   Next i
   SSTab1.Tab = Index
   If SSTab1.Tab = 2 Then
       Label4.Visible = False
       textLI06.Visible = False
   Else
       Label4.Visible = True
       textLI06.Visible = True
   End If
   grd1(Index).row = iRow
   For j = 0 To grd1(Index).Cols - 1
       grd1(Index).col = j
       grd1(Index).CellBackColor = &HFFC0C0
   Next j
   grd1(Index).TopRow = iRow
   grd1(Index).Visible = True
End Sub

Private Sub ChgToNowData()
Dim i, j As Integer
   j = 0
   For i = 1 To grd1(cboLI08.ListIndex).Rows - 1
       If grd1(cboLI08.ListIndex).TextMatrix(i, 0) = textLI02 Then
           j = i
           Exit For
       End If
   Next i
   If j <> 0 Then ChgGrdData cboLI08.ListIndex, j
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim IsRun As Boolean

   cboLI08.ListIndex = SSTab1.Tab
   If SSTab1.Tab = 2 Then
       Label4.Visible = False
       textLI06.Visible = False
   Else
       Label4.Visible = True
       textLI06.Visible = True
   End If
   If grd1(SSTab1.Tab).Rows > 1 Then
       m_CurrKEY(1) = grd1(SSTab1.Tab).TextMatrix(grd1(SSTab1.Tab).Rows - 1, 0)
       m_CurrKEY(2) = Trim(cboLI08.ListIndex + 1)
       UpdateCtrlData
   Else
       ClearField
   End If
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textLI01_GotFocus()
   InverseTextBox textLI01
End Sub

Private Sub textLI01_KeyPress(KeyAscii As Integer)
   If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub textLI01_LostFocus()
   If Trim(textLI01) <> "" And textLI01.Locked = False Then
       m_CurrKEY(0) = ""
       GetAllData
   End If
End Sub

Private Sub textLI01_Validate(Cancel As Boolean)
   If Trim(textLI01) <> "" And m_EditMode = 1 Then
       If CheckIsTaiwanDate(textLI01, False) = False Then
           Cancel = True
           MsgBox "信件日期請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
       ElseIf ChkWorkDay(ChangeTStringToWString(textLI01)) = False Then
           Cancel = True
           MsgBox "信件日期請輸入工作天！", vbInformation, "輸入日期錯誤"
       End If
   End If
End Sub

Private Sub textLI03_GotFocus()
   If Me.SSTab1.Tab = 2 Then CloseIme Else OpenIme
   InverseTextBox textLI03
End Sub

Private Sub textLI03_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textLI03, textLI03.MaxLength) = False Then
       Cancel = True
       Exit Sub
   End If
   CloseIme
End Sub

Private Sub textLI04_GotFocus()
   CloseIme
   InverseTextBox textLI04
End Sub

Private Sub textLI04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textLI04_Validate(Cancel As Boolean)
Dim strCheck As String
   
   If Trim(textLI04) = "" Then Exit Sub 'Add By Sindy 2013/3/12
   
   If CheckLengthIsOK(textLI04, textLI04.MaxLength) = False Then
      Cancel = True
   End If
   'Add By Sindy 2013/3/12 若為國外信件，則檢查是否有輸入正確的系統別
   If Trim(textLI04) <> "" And cboLI08.ListIndex = 2 Then
      strCheck = Trim(textLI04)
      If InStr(1, Trim(textLI04), "-") > 0 Then
         strCheck = Left(Trim(textLI04), InStr(1, Trim(textLI04), "-") - 1)
      End If
      If ClsPDGetSystemKind(strCheck) = False Then
         Cancel = True
         textLI04_GotFocus
      End If
   End If
   '2013/3/12 End
End Sub

Private Sub textLI05_GotFocus()
   OpenIme
   InverseTextBox textLI05
End Sub

Private Sub textLI05_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textLI05, 120) = False Then
       Cancel = True
       Exit Sub
   End If
   CloseIme
End Sub

Private Sub textLI06_GotFocus()
   OpenIme
   InverseTextBox textLI06
End Sub

Private Sub textLI06_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textLI06, 12) = False Then
       Cancel = True
       Exit Sub
   End If
   CloseIme
End Sub

Private Sub textLI09_GotFocus()
   InverseTextBox textLI09
   CloseIme
End Sub

'Modify by Amy 2022/01/11 原:Integer
Private Sub textLI09_KeyPress(KeyAscii As MSForms.ReturnInteger)
   '2011/7/22 MODIFY BY SONIA 開放謝經理操作時可輸小寫
   'KeyAscii = UpperCase(KeyAscii)
   If strUserNum <> "77047" Then KeyAscii = UpperCase(KeyAscii)
   '2011/7/22 END
End Sub

Private Sub textLI09_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textLI09, textLI09.MaxLength) = False Then
       Cancel = True
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: 'OnAction vbKeyF5
      ' 查詢
      Case 4: 'OnAction vbKeyF4
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

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To 8
      strTmp = Format(nIndex, "00")
      'Add By Sindy 98/03/20
      If nIndex = 8 Then
         m_FieldList(nIndex - 1).fiName = "LI15"
      ElseIf nIndex = 7 Then
         m_FieldList(nIndex - 1).fiName = "LI08"
      Else
         m_FieldList(nIndex - 1).fiName = "LI" & strTmp
      End If
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 1, 2
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   m_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         textLI01.Locked = False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         UpdateCtrlData
         m_EditMode = 2
         SetCtrlReadOnly False
         textLI01.Locked = True
         UpdateToolbarState
         SetInputEntry
         'Add By Sindy 2013/3/13 國外信件若有列印次數時,鎖住文號、事由欄位
         If m_EditMode = 2 And cboLI08.ListIndex = 2 And Val(textLI07) > 0 Then
            If MsgBox("此為國外信件並且也已列印清單，確定還要修改資料嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               'textLI04.Locked = True
               'textLI05.Locked = True
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
               Exit Sub
            End If
         End If
         '2013/3/13 End
      ' 刪除
      Case vbKeyF5:
'         strTit = "詢問"
'         strMsg = "是否要刪除此筆資料?"
'         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'         If nResponse = vbYes Then
'            m_EditMode = 3
'            If OnWork = True Then
'                UpdateToolbarState
'            Else
'                Exit Sub
'            End If
'         End If
      ' 查詢
      Case vbKeyF4:
'         Screen.MousePointer = vbHourglass
'         Me.grd1(SSTab1.Tab).MousePointer = flexArrowHourGlass
'         PrintData
'         Me.grd1(SSTab1.Tab).MousePointer = flexDefault
'         Screen.MousePointer = vbDefault
''         SetCtrlReadOnly True
''         ClearField
''         UpdateToolbarState
''         SetInputEntry
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
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
         CloseIme
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT li01,min(li02) li02,li08 FROM letterinput " & _
            "WHERE li01 = (select min(li01) from letterinput)  AND " & _
                  "li08 = (SELECT MIN(li08) FROM letterinput  " & _
                           "where li01 = (select min(li01) from letterinput)) group by li01,li08"
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("li01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("li01")
      If IsNull(rsTmp.Fields("li02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("li02")
      If IsNull(rsTmp.Fields("li08")) = False Then: m_FirstKEY(2) = rsTmp.Fields("li08")
   End If
   rsTmp.Close

   strSql = "SELECT li01,max(li02) li02,li08 FROM letterinput " & _
            "WHERE li01 = (select max(li01) from letterinput)  AND " & _
                  "li08 = (SELECT max(li08) FROM letterinput  " & _
                           "where li01 = (select max(li01) from letterinput)) group by li01,li08"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("li01")) = False Then: m_LastKEY(0) = rsTmp.Fields("li01")
      If IsNull(rsTmp.Fields("li02")) = False Then: m_LastKEY(1) = rsTmp.Fields("li02")
      If IsNull(rsTmp.Fields("li08")) = False Then: m_LastKEY(2) = rsTmp.Fields("li08")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
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

   If Val(m_CurrKEY(0)) = Val(m_FirstKEY(0)) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If

   strSql = "SELECT li01,li02,li08 FROM letterinput " & _
            "WHERE to_char(ltrim(rtrim(li01)))||li08||ltrim(rtrim(to_char(li02,'00000000'))) in (select max(to_char(ltrim(rtrim(li01)))||li08||ltrim(rtrim(to_char(li02,'00000000')))) from letterinput where  to_char(ltrim(rtrim(li01)))||li08||ltrim(rtrim(to_char(li02,'00000000')))<'" & m_CurrKEY(0) & m_CurrKEY(2) & Trim(Format(m_CurrKEY(1), "00000000")) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("li01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("li01")
      If IsNull(rsTmp.Fields("li02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("li02")
      If IsNull(rsTmp.Fields("li08")) = False Then: m_CurrKEY(2) = rsTmp.Fields("li08")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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

   If Val(m_CurrKEY(0)) = Val(m_LastKEY(0)) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If

   strSql = "SELECT li01,li02,li08 FROM letterinput " & _
            "WHERE to_char(ltrim(rtrim(li01)))||li08||ltrim(rtrim(to_char(li02,'00000000'))) in (select min(to_char(ltrim(rtrim(li01)))||li08||ltrim(rtrim(to_char(li02,'00000000')))) from letterinput where  to_char(ltrim(rtrim(li01)))||li08||ltrim(rtrim(to_char(li02,'00000000')))>'" & m_CurrKEY(0) & m_CurrKEY(2) & Trim(Format(m_CurrKEY(1), "00000000")) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("li01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("li01")
      If IsNull(rsTmp.Fields("li02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("li02")
      If IsNull(rsTmp.Fields("li08")) = False Then: m_CurrKEY(2) = rsTmp.Fields("li08")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            Toolbar1.Buttons(1).Enabled = True
         Else
            Toolbar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            Toolbar1.Buttons(2).Enabled = True
         Else
            Toolbar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            Toolbar1.Buttons(3).Enabled = True
         Else
            Toolbar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(4).Enabled = True
         Else
            Toolbar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(6).Enabled = True
            Toolbar1.Buttons(7).Enabled = True
            Toolbar1.Buttons(8).Enabled = True
            Toolbar1.Buttons(9).Enabled = True
         Else
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(7).Enabled = False
            Toolbar1.Buttons(8).Enabled = False
            Toolbar1.Buttons(9).Enabled = False
         End If
         Toolbar1.Buttons(11).Enabled = False
         Toolbar1.Buttons(12).Enabled = False
         Toolbar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         Toolbar1.Buttons(1).Enabled = False
         Toolbar1.Buttons(2).Enabled = False
         Toolbar1.Buttons(3).Enabled = False
         Toolbar1.Buttons(4).Enabled = False
         Toolbar1.Buttons(6).Enabled = False
         Toolbar1.Buttons(7).Enabled = False
         Toolbar1.Buttons(8).Enabled = False
         Toolbar1.Buttons(9).Enabled = False
         Toolbar1.Buttons(11).Enabled = True
         Toolbar1.Buttons(12).Enabled = True
         Toolbar1.Buttons(14).Enabled = False
   End Select
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   cboLI08.Enabled = bEnable
   textLI03.Locked = bEnable
   textLI04.Locked = bEnable
   textLI05.Locked = bEnable
   textLI06.Locked = bEnable
   SSTab1.Enabled = bEnable
   'Add By Sindy 98/03/20
   textLI09.Locked = bEnable
End Sub

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
            If TxtValidate = False Then Exit Function
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            If AddRecord = True Then
                ChgToNowData
            Else
                Exit Function
            End If
      Case 2: '修改
            If TxtValidate = False Then Exit Function
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            '若是更改到信件種類  需先刪除原先的，在新增
            If Trim(cboLI08.ListIndex + 1) <> Trim(m_CurrKEY(2)) Then
                cnnConnection.BeginTrans
                If DelRecord = True Then
                    If AddRecord(False) = True Then
                        ReSortData
                        GetAllData
                        ChgToNowData
                    End If
                End If
                cnnConnection.CommitTrans
            Else
                If ModRecord = False Then Exit Function
            End If
      Case 3: '刪除
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'         UpdateFieldNewData
'         If DelRecord = True Then
'            RefreshRange
'         Else
'            Exit Function
'         End If
      Case 4: '列印
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'         UpdateFieldNewData
'         'If CheckDataValid() = True Then
'         If textCU01 <> "" Then
'            If QueryRecord = False Then
'               strMsg = "無此資料"
'               strTit = "查詢資料"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               UpdateCtrlData
'            End If
'         Else
'            GoTo EXITSUB
'         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

Private Sub ClearField()
Dim nIndex As Integer
   
   textLI02 = Empty
   textLI03 = Empty
   textLI04 = Empty
   textLI05 = Empty
   textLI06 = Empty
   textLI07 = Empty 'Add By Sindy 2013/3/12
   textLI09 = Empty 'Add By Sindy 98/03/20
   For nIndex = 0 To 7
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   strSql = "SELECT * FROM letterinput " & _
            "WHERE li01 = " & Val(m_CurrKEY(0)) & " AND " & _
                  "li02 = " & Val(m_CurrKEY(1)) & " and li08='" & m_CurrKEY(2) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("li08")) = False Then: cboLI08.Text = cboLI08.List(Val(rsTmp.Fields("li08")) - 1): m_CurrKEY(3) = rsTmp.Fields("li08")
      SSTab1.Tab = cboLI08.ListIndex
      If SSTab1.Tab = 2 Then
          Label4.Visible = False
          textLI06.Visible = False
      Else
          Label4.Visible = True
          textLI06.Visible = True
      End If
      If Val(m_CurrKEY(0)) <> Val(ChangeTStringToWString(textLI01)) Then
           GetAllData
      End If
      If IsNull(rsTmp.Fields("li01")) = False Then: textLI01 = ChangeWStringToTString(rsTmp.Fields("li01"))
      If IsNull(rsTmp.Fields("li02")) = False Then: textLI02 = rsTmp.Fields("li02"): m_CurrKEY(1) = textLI02
      If IsNull(rsTmp.Fields("li03")) = False Then: textLI03 = rsTmp.Fields("li03")
      If IsNull(rsTmp.Fields("li04")) = False Then: textLI04 = rsTmp.Fields("li04")
      If IsNull(rsTmp.Fields("li05")) = False Then: textLI05 = rsTmp.Fields("li05")
      If IsNull(rsTmp.Fields("li06")) = False Then: textLI06 = rsTmp.Fields("li06")
      'Add By Sindy 2013/3/12
      If IsNull(rsTmp.Fields("li07")) = False Then: textLI07 = rsTmp.Fields("li07")
      '2013/3/12 End
      'Add By Sindy 98/03/20
      If IsNull(rsTmp.Fields("li15")) = False Then: textLI09 = rsTmp.Fields("li15")
      '98/03/20 End
      ChgToNowData
   End If
   ' 更新暫存區的資料
   UpdateFieldOldData rsTmp
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Sub

'抓當日所有資料
Private Sub GetAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim Rani As Integer
   
   For Rani = 0 To 4 '3
        'Modify By Sindy 98/03/20
        'strSQL = "SELECT li02,li03,li04,li05,li06,li08 FROM letterinput "
        strSql = "SELECT li02,li03,li04,li05,li06,li15,li08 FROM letterinput " & _
                 "WHERE li01 = " & IIf(m_CurrKEY(0) = "", Val(ChangeTStringToWString(textLI01)), Val(m_CurrKEY(0))) & " AND " & _
                       " li08='" & Trim(Rani + 1) & "' order by li02"
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        Set grd1(Rani).Recordset = rsTmp
        rsTmp.Close
    Next Rani
    SetGrd
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   '2007/9/13 modify by sonia 貞瑩說新增時游標停在公司名稱欄
   'Select Case m_EditMode
   '   Case 1: textLI01.SetFocus: textLI01_GotFocus
   '   Case 2: textLI03.SetFocus: textLI03_GotFocus
   'End Select
   textLI03.SetFocus: textLI03_GotFocus
   '2007/9/13 end
End Sub

Private Sub UpdateFieldNewData()
   '若新增資料
   SetFieldNewData "LI01", ChangeTStringToWString(textLI01)
   SetFieldNewData "LI02", textLI02
   SetFieldNewData "LI03", textLI03 'Trim(textLI03) Modify By Sindy 2015/3/20 會有造字不可Trim掉
   SetFieldNewData "LI04", Trim(textLI04)
   SetFieldNewData "LI05", Trim(textLI05)
   SetFieldNewData "LI06", Trim(textLI06)
   SetFieldNewData "LI08", Trim(cboLI08.ListIndex + 1)
   'Add By Sindy 98/03/20
   SetFieldNewData "LI15", textLI09 'Trim(textLI09) Modify By Sindy 2015/3/20 會有造字不可Trim掉
End Sub

Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To 7
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False And rsTmp.RecordCount <> 0 Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord(Optional GoTrans As Boolean = True) As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bFirst As Boolean
Dim rsTmp As New ADODB.Recordset
   
   AddRecord = False
   
   bFirst = True
   'Modify By Sindy 98/03/20
   'strSQL = "INSERT INTO letterinput (li01,li02,li03,li04,li05,li06,li07,li08,li09,li10,li11) select "
   strSql = "INSERT INTO letterinput (li01,li02,li03,li04,li05,li06,li08,li15,li09,li10,li11) select "
   bFirst = True
   For nIndex = 0 To 7
      If nIndex <> 1 Then
            strTmp = Empty
            If m_FieldList(nIndex).fiType = 0 Then
               strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
            Else
               strTmp = m_FieldList(nIndex).fiNewData
            End If
            If strTmp <> Empty Then
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
     Else
         strSql = strSql & ",nvl(max(li02),0)+1 "
     End If
   Next nIndex
   strSql = strSql & ",'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI')) from letterinput where " & IIf(Trim(cboLI08.ListIndex + 1) = "3", " substr(to_char(li01),1,4)=" & Mid(ChangeTStringToWString(textLI01), 1, 4) & " ", " li01=" & ChangeTStringToWString(textLI01) & " ") & " and li08='" & Trim(cboLI08.ListIndex + 1) & "' "
   
On Error GoTo ErrHand
    If GoTrans = True Then
        cnnConnection.BeginTrans
    End If
   
   cnnConnection.Execute strSql
   
   strSql = "select max(li02) from letterinput where " & IIf(Trim(cboLI08.ListIndex + 1) = "3", " substr(to_char(li01),1,4)=" & Mid(ChangeTStringToWString(textLI01), 1, 4) & " ", " li01=" & ChangeTStringToWString(textLI01) & " ") & " and li08='" & Trim(cboLI08.ListIndex + 1) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
        textLI02 = CheckStr(rsTmp.Fields(0))
        'If ((textLI01 & Trim(cboLI08.ListIndex + 1) & textLI02) < (m_FirstKEY(0) & m_FirstKEY(2) & m_FirstKEY(1))) Or ((textLI01 & Trim(cboLI08.ListIndex + 1) & textLI02) > (m_LastKEY(0) & m_LastKEY(2) & m_LastKEY(1))) Then
        If ((ChangeTStringToWString(textLI01) & Trim(cboLI08.ListIndex + 1) & textLI02) < (m_FirstKEY(0) & m_FirstKEY(2) & m_FirstKEY(1))) Or ((ChangeTStringToWString(textLI01) & Trim(cboLI08.ListIndex + 1) & textLI02) > (m_LastKEY(0) & m_LastKEY(2) & m_LastKEY(1))) Then
           RefreshRange
        End If
        If GoTrans = True Then
            cnnConnection.CommitTrans
            GetAllData
            ShowCurrRecord ChangeTStringToWString(textLI01), textLI02, cboLI08.ListIndex + 1
        End If
    Else
        GoTo ErrHand
        Exit Function
    End If
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    If GoTrans = True Then
        MsgBox " 新增失敗！" & vbCrLf & Err.Description
        Resume Next
    Else
        MsgBox " 修改失敗！" & vbCrLf & Err.Description
    End If
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM letterinput " & _
            "WHERE li01 = " & strKEY01 & " AND " & _
                  "li02 = " & strKEY02 & " and li08='" & Trim(strKEY03) & "' "
                  
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

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To 7
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01, strKEY02, strKEY03) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      m_CurrKEY(2) = strKEY03
   Else
      strSql = "SELECT li01,li02,li08 FROM letterinput " & _
               "WHERE li01 = " & m_CurrKEY(0) & " AND " & _
                     "li02 = " & m_CurrKEY(1) & " and " & _
                     "li08='" & Trim(m_CurrKEY(2)) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("li01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("li01")
         If IsNull(rsTmp.Fields("li02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("li02")
         If IsNull(rsTmp.Fields("li08")) = False Then: m_CurrKEY(2) = rsTmp.Fields("li08")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
   strSql = "SELECT li01,li02,li08 FROM letterinput " & _
            "WHERE li01 = " & ChangeTStringToWString(textLI01) & " AND " & _
                  "li02 = (SELECT MIN(li02) FROM letterinput " & _
                           "WHERE li01 = " & ChangeTStringToWString(textLI01) & ") and  " & _
                  "li08 = (SELECT MIN(li08) FROM letterinput  " & _
                           "where li01 = " & ChangeTStringToWString(textLI01) & " and " & _
                           " li02= (SELECT MIN(li02) FROM letterinput " & _
                           " WHERE li01 = " & ChangeTStringToWString(textLI01) & ") ) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("li01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("li01")
         If IsNull(rsTmp.Fields("li02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("li02")
         If IsNull(rsTmp.Fields("li08")) = False Then: m_CurrKEY(2) = rsTmp.Fields("li08")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   'Add by Amy 2022/01/11檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        Exit Function
   End If

   'add by sonia 2018/5/22
   If Me.textLI01.Enabled = True Then
      Cancel = False
      textLI01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2018/5/22
   
   If Me.textLI03.Enabled = True Then
      Cancel = False
      textLI03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textLI04.Enabled = True Then
      Cancel = False
      textLI04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textLI05.Enabled = True Then
      Cancel = False
      textLI05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textLI06.Enabled = True Then
      Cancel = False
      textLI06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Trim(textLI03.Text & textLI04.Text & textLI05.Text & textLI06.Text) = "" Then
       MsgBox "最少輸入一欄資料！", vbInformation, "操作錯誤！"
       textLI03.SetFocus
       Exit Function
   End If
   
   'Add By Sindy 2013/3/12
   If cboLI08.ListIndex = 2 Then
      If Trim(textLI04.Text) <> "" And (Left(Trim(textLI05.Text), 2) = "支票" Or Left(Trim(textLI05.Text), 4) = "匯款通知") Then
          MsgBox "國外信件不可同時輸入文號和事由有支票二個字或匯款通知四個字的內容，請二選一做輸入！", vbInformation, "操作錯誤！"
          Exit Function
      End If
   End If
   '2013/3/12 End
   TxtValidate = True
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strLI01 As String
Dim strLI02 As String
Dim strLI08 As String
   
   ModRecord = False
   
   strLI01 = m_CurrKEY(0)
   strLI02 = m_CurrKEY(1)
   strLI08 = m_CurrKEY(2)
   strSql = "UPDATE letterinput SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To 7
        strTmp = Empty
        If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
           If m_FieldList(nIndex).fiType = 0 Then
              If m_FieldList(nIndex).fiNewData = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
              End If
           Else
              If m_FieldList(nIndex).fiNewData = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
              End If
           End If
        End If
        If strTmp <> Empty Then
           bDifference = True
           If bFirst = True Then
              strSql = strSql & strTmp
              bFirst = False
           Else
              strSql = strSql & "," & strTmp
           End If
        End If
   Next nIndex

   strSql = strSql & ",li12='" & strUserNum & "',li13=to_number(to_char(sysdate,'YYYYMMDD')),li14=to_number(to_char(sysdate,'HH24MI')) " & _
                  "WHERE li01 = " & strLI01 & " AND " & _
                        "li02 = " & strLI02 & " and  li08='" & strLI08 & "' "
On Error GoTo ErrHand
   If bDifference = True Then
      cnnConnection.BeginTrans
      
      cnnConnection.Execute strSql

      cnnConnection.CommitTrans
      
      GetAllData
      ShowCurrRecord strLI01, strLI02, strLI08
   End If
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
    Resume Next
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strLI01 As String
Dim strLI02 As String
Dim strLI08 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   strLI01 = m_CurrKEY(0)
   strLI02 = m_CurrKEY(1)
   strLI08 = m_CurrKEY(2)

   strSql = "DELETE FROM letterinput " & _
            "WHERE li01 = " & strLI01 & " AND " & _
                  "li02 = " & strLI02 & " and li08='" & strLI08 & "'"
   
   cnnConnection.Execute strSql

   DelRecord = True
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "修改失敗！" & vbCrLf & Err.Description
End Function

'重整資料
Private Sub ReSortData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim MaxByOverSeas As Long
Dim NowCount As Long
Dim NowLI08 As String

On Error GoTo ErrHand
   '抓國外信件的上月最大值
    strSql = "SELECT nvl(max(li02),0) FROM letterinput " & _
             "WHERE substr(li01,1,6) = " & Mid(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(m_CurrKEY(0)))), 1, 6) & " AND " & _
                   "li08='3' "
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If Not rsTmp.EOF And Not rsTmp.BOF Then
       MaxByOverSeas = Val(CheckStr(rsTmp.Fields(0)))
    End If
    rsTmp.Close
    '先將所有的序號往後推 10000 號
    cnnConnection.Execute "update letterinput set li02=li02 + 10000 where li01=" & m_CurrKEY(0) & " "
    strSql = "SELECT li01,li02,li08 FROM letterinput " & _
                   "WHERE li01 = " & m_CurrKEY(0) & " order by li08,li02 "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
          rsTmp.MoveFirst
          NowCount = 0
          NowLI08 = ""
          Do While Not rsTmp.EOF
                If NowLI08 <> CheckStr(rsTmp.Fields("li08")) Then
                    If CheckStr(rsTmp.Fields("li08")) <> "3" Then
                        NowCount = 1
                    Else
                        NowCount = MaxByOverSeas + 1
                    End If
                    NowLI08 = CheckStr(rsTmp.Fields("li08"))
                End If
                '將原先畫面那筆資料換成新的序號
                If Trim(Val(textLI02) + 10000) = CheckStr(rsTmp.Fields("li02")) And CheckStr(rsTmp.Fields("li08")) = Trim(cboLI08.ListIndex + 1) Then
                    textLI02 = NowCount
                    m_CurrKEY(1) = Trim(NowCount)
                End If
                cnnConnection.Execute "update letterinput set li02=" & NowCount & " where li01=" & CheckStr(rsTmp.Fields("li01")) & " and li02=" & CheckStr(rsTmp.Fields("li02")) & " and li08='" & CheckStr(rsTmp.Fields("li08")) & "' "
                NowCount = NowCount + 1
                rsTmp.MoveNext
          Loop
      End If
      rsTmp.Close
   Exit Sub
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "修改失敗！" & vbCrLf & Err.Description
End Sub

'非國外信件列印
Sub PrintData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim NowLI08 As String
Dim tmpLi05 As String
   
   'Add By Sindy 2014/3/4
   m_PrintDate = InputBox("請輸入欲列印的日期：", "非國外信件列印", TransDate(m_CurrKEY(0), 1))
   If m_PrintDate = "" Then '取消
      Exit Sub
   Else
      m_PrintDate = DBDATE(m_PrintDate)
   End If
   '檢查是否為日期格式
   If ChkDate(m_PrintDate) = False Then
      Exit Sub
   End If
   If ChkWorkDay(m_PrintDate) = False Then
      MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
      Exit Sub
   End If
   '2014/3/4 END
   
   'Modify By Sindy +and li08<>'3' 不可為國外信件
   strSql = "SELECT * FROM letterinput " & _
            "WHERE li01= " & m_PrintDate & " and li08='" & Trim(cboLI08.ListIndex + 1) & "' and li08<>'3' order by li08,li02 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <> 0 Then
       Set Printer = Printers(Combo1.ListIndex)
       rsTmp.MoveFirst
       NowLI08 = CheckStr(rsTmp.Fields("li08"))
       PrintTitle NowLI08
       iLine = 0
       Do While Not rsTmp.EOF
           If NowLI08 <> CheckStr(rsTmp.Fields("li08")) Then
               NowLI08 = CheckStr(rsTmp.Fields("li08"))
               'add by nickc 2007/08/23 加入印頁數在下方
'               Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
'               Printer.CurrentY = 15500 '2000 + (41 * 320) + 50
'               Printer.Print Trim(Printer.Page)
               Printer.NewPage
               PrintTitle NowLI08
               iLine = 0
           End If
'           If NowLI08 = "3" Then
'               MaxLine = 40
'               strTemp(1) = CheckStr(rsTmp.Fields("li02"))
'               strTemp(2) = CheckStr(rsTmp.Fields("li03"))
'               strTemp(3) = CheckStr(rsTmp.Fields("li04"))
'               strTemp(4) = CheckStr(rsTmp.Fields("li05"))
'               strTemp(5) = ""
'               strTemp(6) = ""
'               strTemp(7) = CheckStr(rsTmp.Fields("li15")) 'Add By Sindy 98/03/23
'           Else
               MaxLine = 20
               strTemp(1) = CheckStr(rsTmp.Fields("li02"))
               strTemp(2) = CheckStr(rsTmp.Fields("li03"))
               strTemp(3) = CheckStr(rsTmp.Fields("li04"))
               strTemp(4) = CheckStr(rsTmp.Fields("li05"))
               strTemp(5) = ""
               strTemp(6) = CheckStr(rsTmp.Fields("li06"))
               strTemp(7) = CheckStr(rsTmp.Fields("li15")) 'Add By Sindy 98/03/23
'           End If
           'add by nickc 2007/08/02
           If GetTextLength(strTemp(4)) > 24 Then
               tmpLi05 = strTemp(4)
               strTemp(4) = StrToStr(tmpLi05, 12)
               tmpLi05 = Replace(tmpLi05, strTemp(4), "", 1, 1)
               PrintDetil NowLI08
               iLine = iLine + 1
               If iLine >= MaxLine Then
                   'add by nickc 2007/08/23 加入印頁數在下方
'                   Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
'                   Printer.CurrentY = 15500 '2000 + (41 * 320) + 50
'                   Printer.Print Trim(Printer.Page)
                   Printer.NewPage
                   PrintTitle NowLI08
               End If
               strTemp(1) = ""
               strTemp(2) = ""
               strTemp(3) = ""
               strTemp(4) = ""
               strTemp(5) = ""
               strTemp(6) = ""
               strTemp(7) = "" 'Add By Sindy 98/03/23
               Do While GetTextLength(tmpLi05) <> 0
                   strTemp(4) = StrToStr(tmpLi05, 24)
                   tmpLi05 = Replace(tmpLi05, strTemp(4), "", 1, 1)
                   PrintDetil NowLI08
                   iLine = iLine + 1
                   If iLine >= MaxLine Then
                       'add by nickc 2007/08/23 加入印頁數在下方
'                       Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
'                       Printer.CurrentY = 15500 '2000 + (41 * 320) + 50
'                       Printer.Print Trim(Printer.Page)
                       Printer.NewPage
                       PrintTitle NowLI08
                   End If
               Loop
           Else
               PrintDetil NowLI08
               iLine = iLine + 1
           End If
           If iLine >= MaxLine Then
               'add by nickc 2007/08/23 加入印頁數在下方
'               Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
'               Printer.CurrentY = 15500 '2000 + (41 * 320) + 50
'               Printer.Print Trim(Printer.Page)
               Printer.NewPage
               PrintTitle NowLI08
           End If
           rsTmp.MoveNext
       Loop
       'add by nickc 2007/08/23 加入印頁數在下方
'       Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
'       Printer.CurrentY = 15500 '2000 + (41 * 320) + 50
'       Printer.Print Trim(Printer.Page)
       Printer.EndDoc
       ShowPrintOk
   Else
       MsgBox "沒有資料可以列印！", , "錯誤！"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Sub PrintTitle(strT1 As String)
Dim i As Integer

On Error GoTo MsgErr

   '畫框
   If Printer.Orientation <> 1 Then
       Printer.Orientation = 1
   End If
   Printer.DrawWidth = 20
   'Printer.Line (800, 1320)-(11000, 15600), , B
   Printer.Line (800, 1320)-(11000, 15500), , B
   iLine = 0
   '定位 畫格子
   Select Case strT1
      Case "1"
           Printer.DrawWidth = 1
           For i = 1 To 20
               Printer.Line (800, 1320)-(11000, 15500 - (680 * i)), , B
           Next i
           PLeft(1) = 900
           PLeft(2) = 1750
           PLeft(3) = 4050
           PLeft(4) = 5050
           PLeft(5) = 7550
           PLeft(6) = 8550
           PLeft(7) = 9550 'Add By Sindy 98/03/23
           Printer.Line (1700, 1320)-(1700, 15500) '公司名稱
           Printer.Line (4000, 1320)-(4000, 15500) '文號
           Printer.Line (5000, 1320)-(5000, 15500) '事由
           Printer.Line (7500, 1320)-(7500, 15500) '收件人
           Printer.Line (8500, 1900)-(8500, 15500) '(8500, 2000)-(8500, 15500)
           Printer.Line (9500, 1320)-(9500, 15500) '備註 Add By Sindy 98/03/23
           Printer.Font.Size = 22
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("大陸 一般信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)) / 2
           Printer.CurrentY = 600
           Printer.Print "大陸 一般信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)
           Printer.Font.Size = 14
           Printer.CurrentX = 800 + ((1700 - 800) / 2) - Printer.TextWidth("流水號") / 2
           Printer.CurrentY = 1500
           Printer.Print "流水號"
           Printer.CurrentX = 1500 + ((4240 - 1700) / 2) - Printer.TextWidth("公司名稱") / 2
           Printer.CurrentY = 1500
           Printer.Print "公司名稱"
           Printer.CurrentX = 3950 + ((5310 - 4240) / 2) - Printer.TextWidth("文號") / 2
           Printer.CurrentY = 1500
           Printer.Print "文號"
           Printer.CurrentX = 4800 + ((8130 - 5310) / 2) - Printer.TextWidth("事由") / 2
           Printer.CurrentY = 1500
           Printer.Print "事由"
           Printer.CurrentX = 7000 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "收件人"
           'Add By Sindy 98/03/23
           Printer.CurrentX = 8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "備註"
           '98/03/23 End
      Case "2"
           Printer.DrawWidth = 1
           For i = 1 To 20
               Printer.Line (800, 1320)-(11000, 15500 - (680 * i)), , B
           Next i
           PLeft(1) = 900
           PLeft(2) = 1750
           PLeft(3) = 4050
           PLeft(4) = 5050
           PLeft(5) = 7550
           PLeft(6) = 8550
           PLeft(7) = 9550 'Add By Sindy 98/03/23
           Printer.Line (1700, 1320)-(1700, 15500) '公司名稱
           Printer.Line (4000, 1320)-(4000, 15500) '文號
           Printer.Line (5000, 1320)-(5000, 15500) '事由
           Printer.Line (7500, 1320)-(7500, 15500) '收件人
           Printer.Line (8500, 1900)-(8500, 15500) '(8500, 2000)-(8500, 15500)
           Printer.Line (9500, 1320)-(9500, 15500) '備註 Add By Sindy 98/03/23
           Printer.Font.Size = 22
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("大陸 一般信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)) / 2
           Printer.CurrentY = 600
           Printer.Print "非大陸 一般信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)
           Printer.Font.Size = 14
           Printer.CurrentX = 800 + ((1700 - 800) / 2) - Printer.TextWidth("流水號") / 2
           Printer.CurrentY = 1500
           Printer.Print "流水號"
           Printer.CurrentX = 1500 + ((4240 - 1700) / 2) - Printer.TextWidth("公司名稱") / 2
           Printer.CurrentY = 1500
           Printer.Print "公司名稱"
           Printer.CurrentX = 3950 + ((5310 - 4240) / 2) - Printer.TextWidth("文號") / 2
           Printer.CurrentY = 1500
           Printer.Print "文號"
           Printer.CurrentX = 4800 + ((8130 - 5310) / 2) - Printer.TextWidth("事由") / 2
           Printer.CurrentY = 1500
           Printer.Print "事由"
           Printer.CurrentX = 7000 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "收件人"
           'Add By Sindy 98/03/23
           Printer.CurrentX = 8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "備註"
           '98/03/23 End
'      Case "3"
'           Printer.DrawWidth = 1
'           For i = 1 To 40
'               Printer.Line (800, 1320)-(11000, 15500 - (320 * i)), , B
'           Next i
'           PLeft(1) = 900
'           PLeft(2) = 1950
'           PLeft(3) = 5950
'           PLeft(4) = 8050
'           PLeft(7) = 9550 'Add By Sindy 98/03/23
'           Printer.Line (1900, 1320)-(1900, 15500)
'           Printer.Line (5900, 2000)-(5900, 15500)
'           Printer.Line (8000, 1320)-(8000, 15500)
'           Printer.Line (9500, 1320)-(9500, 15500) '備註 Add By Sindy 98/03/23
'           Printer.Font.Size = 22
'           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("國外信件     中華民國  " & ChangeTStringToTDateString(strSrvDate(2))) / 2
'           Printer.CurrentY = 600
'           Printer.Print "國外信件     中華民國  " & ChangeTStringToTDateString(strSrvDate(2))
'           Printer.Font.Size = 14
'           Printer.CurrentX = 300 + ((2900 - 800) / 2) - Printer.TextWidth("流水號") / 2
'           Printer.CurrentY = 1500
'           Printer.Print "流水號"
'           Printer.CurrentX = 1300 + ((9000 - 2900) / 2) - Printer.TextWidth("公司名稱") / 2
'           Printer.CurrentY = 1500
'           Printer.Print "公司名稱"
'           Printer.CurrentX = 7700 + ((11000 - 9000) / 2) - Printer.TextWidth("事由") / 2
'           Printer.CurrentY = 1500
'           Printer.Print "事由"
'           'Add By Sindy 98/03/23
'           Printer.CurrentX = 8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
'           Printer.CurrentY = 1500
'           Printer.Print "備註"
'           '98/03/23 End
      Case "4"
           Printer.DrawWidth = 1
           For i = 1 To 20
               Printer.Line (800, 1320)-(11000, 15500 - (680 * i)), , B
           Next i
           PLeft(1) = 900
           PLeft(2) = 1750
           PLeft(3) = 4050
           PLeft(4) = 5050
           PLeft(5) = 7550
           PLeft(6) = 8550
           PLeft(7) = 9550 'Add By Sindy 98/03/23
           Printer.Line (1700, 1320)-(1700, 15500) '公司名稱
           Printer.Line (4000, 1320)-(4000, 15500) '文號
           Printer.Line (5000, 1320)-(5000, 15500) '事由
           Printer.Line (7500, 1320)-(7500, 15500) '收件人
           Printer.Line (8500, 1900)-(8500, 15500) '(8500, 2000)-(8500, 15500)
           Printer.Line (9500, 1320)-(9500, 15500) '備註 Add By Sindy 98/03/23
           Printer.Font.Size = 22
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("客      戶      中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)) / 2
           Printer.CurrentY = 600
           Printer.Print "客      戶      中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)
           Printer.Font.Size = 14
           Printer.CurrentX = 800 + ((1700 - 800) / 2) - Printer.TextWidth("流水號") / 2
           Printer.CurrentY = 1500
           Printer.Print "流水號"
           Printer.CurrentX = 1500 + ((4240 - 1700) / 2) - Printer.TextWidth("公司名稱") / 2
           Printer.CurrentY = 1500
           Printer.Print "公司名稱"
           Printer.CurrentX = 3950 + ((5310 - 4240) / 2) - Printer.TextWidth("文號") / 2
           Printer.CurrentY = 1500
           Printer.Print "文號"
           Printer.CurrentX = 4800 + ((8130 - 5310) / 2) - Printer.TextWidth("事由") / 2
           Printer.CurrentY = 1500
           Printer.Print "事由"
           Printer.CurrentX = 7000 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "收件人"
           'Add By Sindy 98/03/23
           Printer.CurrentX = 8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "備註"
           '98/03/23 End
      'Add By Sindy 98/03/23
      Case "5"
           Printer.DrawWidth = 1
           For i = 1 To 20
               Printer.Line (800, 1320)-(11000, 15500 - (680 * i)), , B
           Next i
           PLeft(1) = 900
           PLeft(2) = 1750
           PLeft(3) = 4050
           PLeft(4) = 5050
           PLeft(5) = 7550
           PLeft(6) = 8550
           PLeft(7) = 9550 'Add By Sindy 98/03/23
           Printer.Line (1700, 1320)-(1700, 15500) '公司名稱
           Printer.Line (4000, 1320)-(4000, 15500) '文號
           Printer.Line (5000, 1320)-(5000, 15500) '事由
           Printer.Line (7500, 1320)-(7500, 15500) '收件人
           Printer.Line (8500, 1900)-(8500, 15500) '(8500, 2000)-(8500, 15500)
           Printer.Line (9500, 1320)-(9500, 15500) '備註 Add By Sindy 98/03/23
           Printer.Font.Size = 22
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("客      戶      中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)) / 2
           Printer.CurrentY = 600
           Printer.Print "退件      中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)
           Printer.Font.Size = 14
           Printer.CurrentX = 800 + ((1700 - 800) / 2) - Printer.TextWidth("流水號") / 2
           Printer.CurrentY = 1500
           Printer.Print "流水號"
           Printer.CurrentX = 1500 + ((4240 - 1700) / 2) - Printer.TextWidth("公司名稱") / 2
           Printer.CurrentY = 1500
           Printer.Print "公司名稱"
           Printer.CurrentX = 3950 + ((5310 - 4240) / 2) - Printer.TextWidth("文號") / 2
           Printer.CurrentY = 1500
           Printer.Print "文號"
           Printer.CurrentX = 4800 + ((8130 - 5310) / 2) - Printer.TextWidth("事由") / 2
           Printer.CurrentY = 1500
           Printer.Print "事由"
           Printer.CurrentX = 7000 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "收件人"
           'Add By Sindy 98/03/23
           Printer.CurrentX = 8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
           Printer.CurrentY = 1500
           Printer.Print "備註"
           '98/03/23 End
      Case Else
   End Select
   'Modify By Sindy 2018/3/23
   Printer.Font.Size = 10
   Printer.CurrentX = 10000
   Printer.CurrentY = 900
   Printer.Print "第 " & Trim(Printer.Page) & " 頁"
   '2018/3/23 END
   Exit Sub
MsgErr:
   MsgBox "列印發生錯誤！" & vbCrLf & Err.Description, vbInformation, "櫃檯信件列印！"
End Sub

Sub PrintDetil(strT1 As String)
Dim i As Integer

On Error GoTo MsgErr

   '定位 畫格子
   Select Case strT1
'      Case "3"
'           Printer.Font.Size = 12
'           Printer.CurrentX = PLeft(1)
'           Printer.CurrentY = 2000 + (iLine * 320) + 50
'           Printer.Print strTemp(1)
'           Printer.CurrentX = PLeft(2)
'           Printer.CurrentY = 2000 + (iLine * 320) + 50
'           Printer.Print strTemp(2)
'           Printer.CurrentX = PLeft(3)
'           Printer.CurrentY = 2000 + (iLine * 320) + 50
'           Printer.Print strTemp(3)
'           Printer.CurrentX = PLeft(4)
'           Printer.CurrentY = 2000 + (iLine * 320) + 50
'           Printer.Print strTemp(4)
'           Printer.CurrentX = PLeft(5)
'           Printer.CurrentY = 2000 + (iLine * 320) + 50
'           Printer.Print strTemp(5)
'           Printer.CurrentX = PLeft(6)
'           Printer.CurrentY = 2000 + (iLine * 320) + 50
'           Printer.Print strTemp(6)
'           'Add By Sindy 98/03/23
'           Printer.CurrentX = PLeft(7)
'           Printer.CurrentY = 2000 + (iLine * 320) + 50
'           Printer.Print strTemp(7)
'           '98/03/23 End
      Case Else
           Printer.Font.Size = 12
           Printer.CurrentX = PLeft(1)
           Printer.CurrentY = 2000 + (iLine * 680) + 200
           Printer.Print strTemp(1)
           Printer.CurrentX = PLeft(2)
           Printer.CurrentY = 2000 + (iLine * 680) + 200
           Printer.Print strTemp(2)
           Printer.CurrentX = PLeft(3)
           Printer.CurrentY = 2000 + (iLine * 680) + 200
           Printer.Print strTemp(3)
           Printer.CurrentX = PLeft(4)
           Printer.CurrentY = 2000 + (iLine * 680) + 200
           Printer.Print strTemp(4)
           Printer.CurrentX = PLeft(5)
           Printer.CurrentY = 2000 + (iLine * 680) + 200
           Printer.Print strTemp(5)
           Printer.CurrentX = PLeft(6)
           Printer.CurrentY = 2000 + (iLine * 680) + 200
           Printer.Print strTemp(6)
           'Add By Sindy 98/03/23
           Printer.CurrentX = PLeft(7)
           Printer.CurrentY = 2000 + (iLine * 680) + 200
           Printer.Print strTemp(7)
           '98/03/23 End
   End Select
   Exit Sub
MsgErr:
   MsgBox "列印發生錯誤！" & vbCrLf & Err.Description, vbInformation, "櫃檯信件列印！"
End Sub

'Add By Sindy 2013/3/12
'國外信件列印
Sub PrintData3()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim NowLI08 As String, tmpLi05 As String
Dim bolHaveData As Boolean
Dim i As Integer, strCon As String
   
On Error GoTo ErrHand
      
   bolHaveData = False
   
   'Add By Sindy 2014/3/4
   m_PrintDate = InputBox("請輸入欲列印的日期：", "國外信件列印", TransDate(m_CurrKEY(0), 1))
   If m_PrintDate = "" Then '取消
      Exit Sub
   Else
      m_PrintDate = DBDATE(m_PrintDate)
   End If
   '檢查是否為日期格式
   If ChkDate(m_PrintDate) = False Then
      Exit Sub
   End If
   If ChkWorkDay(m_PrintDate) = False Then
      MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
      Exit Sub
   End If
   '2014/3/4 END
   
   cnnConnection.BeginTrans
   If Val(Trim(Text1)) > 0 Then
'      'Add By Sindy 2013/4/2
'      strSql = "SELECT * FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3'" & _
'               " and li07='" & Val(Trim(Text1)) & "'"
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <> 0 Then
'         If MsgBox("確定要重新列印第 " & Val(Trim(Text1)) & " 批的資料嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'            cnnConnection.CommitTrans
'            rsTmp.Close
'            Set rsTmp = Nothing
'            Exit Sub
'         End If
'      End If
'      rsTmp.Close
'      '2013/4/2 End
      
      '有輸入列印次數,代表為重新列印
      strSql = "delete from ForeignLetterCount where FLC01=" & m_PrintDate & " and FLC03=" & Text1
      cnnConnection.Execute strSql
   Else
      '取得最大列印次數
      strExc(0) = "SELECT max(li07) FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3' and li07>0 and li07 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text1 = Val("" & RsTemp.Fields(0)) + 1
      End If
      
      '更新列印次數
      strSql = "update letterinput set li07='" & Text1 & "' WHERE li01=" & m_PrintDate & " and li08='3' and li07 is null"
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans 'Add By Sindy 2013/4/2 以防資料丟到印表機時出現異常狀況
   
   strCon = " and li07=" & Text1
   cnnConnection.BeginTrans 'Add By Sindy 2013/4/2
   For i = 1 To 7
      Call GetSystemDept(i)
      If i = 1 Then '財務處
         MaxLine = 38
         strSql = "SELECT * FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3'" & _
                  strCon & " and (substr(li05,1,2)='支票' or substr(li05,1,4)='匯款通知')" & _
                  " order by li02"
      ElseIf i = 2 Or i = 3 Or i = 4 Or i = 5 Or i = 6 Then
         MaxLine = 38
         strSql = "select * from (" & _
                 " SELECT * FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3' and li04 is not null and instr(li04,'-')>0" & _
                 strCon & " and instr('" & strSysDept & "','<'||substr(li04,1,instr(li04,'-')-1)||'>')>0" & _
                 " Union" & _
                 " SELECT * FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3' and li04 is not null and instr(li04,'-')=0" & _
                 strCon & " and instr('" & strSysDept & "','<'||li04||'>')>0" & _
                 ") order by li02"
      ElseIf i = 7 Then 'X.未分類
         MaxLine = 12
         strSql = "SELECT * FROM letterinput WHERE li01=" & m_PrintDate & " and li08='3' and li04 is null and (li05 is null or (substr(li05,1,2)<>'支票' and substr(li05,1,4)<>'匯款通知'))" & _
                 strCon & " order by li02"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <> 0 Then
         bolHaveData = True
         Set Printer = Printers(Combo1.ListIndex)
         rsTmp.MoveFirst
         NowLI08 = CheckStr(rsTmp.Fields("li08"))
         PrintTitle3 i
         Do While Not rsTmp.EOF
            If NowLI08 <> CheckStr(rsTmp.Fields("li08")) Then
                NowLI08 = CheckStr(rsTmp.Fields("li08"))
                '加入印頁數在下方
                Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
                Printer.CurrentY = 2000 + (41 * 340) + 50
                Printer.Print Trim(Printer.Page)
                Printer.NewPage
                PrintTitle3 i
            End If
            strTemp(1) = CheckStr(rsTmp.Fields("li02"))
            strTemp(2) = CheckStr(rsTmp.Fields("li03"))
            strTemp(3) = CheckStr(rsTmp.Fields("li04"))
            strTemp(4) = CheckStr(rsTmp.Fields("li05"))
            strTemp(5) = ""
            strTemp(6) = ""
            strTemp(7) = CheckStr(rsTmp.Fields("li15"))
            
            If GetTextLength(strTemp(4)) > 24 Then
                tmpLi05 = strTemp(4)
                strTemp(4) = StrToStr(tmpLi05, 12)
                tmpLi05 = Replace(tmpLi05, strTemp(4), "", 1, 1)
                PrintDetil3 i
                iLine = iLine + 1
                If iLine >= MaxLine Then
                    '加入印頁數在下方
                    Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
                    Printer.CurrentY = 2000 + (41 * 340) + 50
                    Printer.Print Trim(Printer.Page)
                    Printer.NewPage
                    PrintTitle3 i
                End If
                strTemp(1) = ""
                strTemp(2) = ""
                strTemp(3) = ""
                strTemp(4) = ""
                strTemp(5) = ""
                strTemp(6) = ""
                strTemp(7) = ""
                Do While GetTextLength(tmpLi05) <> 0
                    strTemp(4) = StrToStr(tmpLi05, 24)
                    tmpLi05 = Replace(tmpLi05, strTemp(4), "", 1, 1)
                    PrintDetil3 i
                    iLine = iLine + 1
                    If iLine >= MaxLine Then
                        '加入印頁數在下方
                        Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
                        Printer.CurrentY = 2000 + (41 * 340) + 50
                        Printer.Print Trim(Printer.Page)
                        Printer.NewPage
                        PrintTitle3 i
                    End If
                Loop
            Else
                PrintDetil3 i
                iLine = iLine + 1
            End If
            If iLine >= MaxLine Then
                '加入印頁數在下方
                Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
                Printer.CurrentY = 2000 + (41 * 340) + 50
                Printer.Print Trim(Printer.Page)
                Printer.NewPage
                PrintTitle3 i
            End If
            rsTmp.MoveNext
         Loop
         '加入簽核欄位
         Printer.Font.Size = 14
         Printer.CurrentX = PLeft(1)
         If i = 7 Then
            Printer.CurrentY = 2000 + ((MaxLine + 1) * 680) + 50
         Else
            Printer.CurrentY = 2000 + ((MaxLine + 1) * 340) + 50
         End If
         Printer.Print "合計　" & rsTmp.RecordCount & "　筆，請簽收："
         '加入印頁數在下方
         Printer.Font.Size = 12
         If i = 7 Then
            Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
            Printer.CurrentY = 2000 + ((MaxLine + 1) * 680) + 200
            Printer.Print Trim(Printer.Page)
         Else
            Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(Trim(Printer.Page))
            Printer.CurrentY = 2000 + ((MaxLine + 2) * 340) + 50
            Printer.Print Trim(Printer.Page)
         End If
         Printer.EndDoc
         
         strSql = "insert into ForeignLetterCount(FLC01,FLC02,FLC03,FLC04) values(" & m_PrintDate & "," & CNULL(strDept) & "," & Text1 & "," & Val(Trim(Printer.Page)) & ")"
         cnnConnection.Execute strSql
      End If
      rsTmp.Close
   Next i
   Text1 = ""
   cnnConnection.CommitTrans
   
   If bolHaveData = False Then
      MsgBox "沒有資料可以列印！", , "錯誤！"
      Text1.SetFocus
   Else
      ShowPrintOk
   End If
   
   Set rsTmp = Nothing
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2013/3/12
Sub PrintTitle3(intT1 As Integer)
Dim i As Integer
   
On Error GoTo MsgErr
   
   If intT1 = 7 Then '未分類
      If Printer.Orientation <> 2 Then
         Printer.Orientation = 2 '橫印
      End If
      '畫框
      Printer.DrawWidth = 20
      Printer.Line (800, 1660)-(16000, 10840), , B
      iLine = 1
      '定位 畫格子
      Printer.DrawWidth = 1
      For i = 1 To MaxLine
          Printer.Line (800, 1660)-(16000, 10840 - (680 * i)), , B
      Next i
      PLeft(1) = 900
      PLeft(2) = 1950
      PLeft(3) = 5950
      PLeft(4) = 8050
      PLeft(7) = 9550
      Printer.Line (1900, 1660)-(1900, 10840)
      Printer.Line (5900, 2700)-(5900, 10840)
      Printer.Line (8000, 1660)-(8000, 10840)
      Printer.Line (9500, 1660)-(9500, 10840) '備註
      Printer.Line (12000, 1660)-(12000, 10840) '正確部門
      Printer.Line (14000, 1660)-(14000, 10840) '簽收人員
      Printer.Font.Size = 22
      Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("國外信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)) / 2
      Printer.CurrentY = 600
      Printer.Print "國外信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)
      Printer.Font.Size = 14
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = 1250
      Printer.Print strDeptName
      Printer.CurrentX = 10000
      Printer.CurrentY = 1250
      Printer.Print strSignPer
      'Printer.Font.Size = 14
      Printer.CurrentX = 300 + ((2900 - 800) / 2) - Printer.TextWidth("流水號") / 2
      Printer.CurrentY = 1940
      Printer.Print "流水號"
      Printer.CurrentX = 1300 + ((9000 - 2900) / 2) - Printer.TextWidth("公司名稱") / 2
      Printer.CurrentY = 1940
      Printer.Print "公司名稱"
      Printer.CurrentX = 7700 + ((11000 - 9000) / 2) - Printer.TextWidth("事由") / 2
      Printer.CurrentY = 1940
      Printer.Print "事由"
      Printer.CurrentX = 8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
      Printer.CurrentY = 1940
      Printer.Print "備註"
      Printer.CurrentX = 11500 + ((11000 - 8130) / 2) - Printer.TextWidth("正確部門") / 2
      Printer.CurrentY = 1940
      Printer.Print "正確部門"
      Printer.CurrentX = 13500 + ((11000 - 8130) / 2) - Printer.TextWidth("簽收人員") / 2
      Printer.CurrentY = 1940
      Printer.Print "簽收人員"
   ElseIf intT1 = 1 Then '支票
      If Printer.Orientation <> 1 Then
         Printer.Orientation = 1 '直印
      End If
      '畫框
      Printer.DrawWidth = 20
      Printer.Line (800, 1660)-(11000, 15260), , B
      iLine = 1
      '定位 畫格子
      Printer.DrawWidth = 1
      For i = 1 To MaxLine
          Printer.Line (800, 1660)-(11000, 15260 - (340 * i)), , B
      Next i
      PLeft(1) = 900
      PLeft(2) = 1950
      PLeft(3) = 5950
      PLeft(4) = 8050
      PLeft(7) = 9550
      Printer.Line (1900, 1660)-(1900, 15260)
      Printer.Line (5900, 1660)-(5900, 15260) '(5900, 2340)-(5900, 15260)
      'Printer.Line (8000, 1660)-(8000, 15260)
      Printer.Line (8900, 1660)-(8900, 15260) '備註
      Printer.Font.Size = 22
      Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("國外信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)) / 2
      Printer.CurrentY = 600
      Printer.Print "國外信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)
      Printer.Font.Size = 14
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = 1250
      Printer.Print strDeptName
      Printer.CurrentX = 5500 '4000
      Printer.CurrentY = 1250
      Printer.Print strSignPer
      Printer.CurrentX = 300 + ((2900 - 800) / 2) - Printer.TextWidth("流水號") / 2
      Printer.CurrentY = 1940
      Printer.Print "流水號"
      Printer.CurrentX = 3200 '1300 + ((9000 - 2900) / 2) - Printer.TextWidth("公司名稱") / 2
      Printer.CurrentY = 1940
      Printer.Print "公司名稱"
      Printer.CurrentX = 7100 '7700 + ((11000 - 9000) / 2) - Printer.TextWidth("事由") / 2
      Printer.CurrentY = 1940
      Printer.Print "事由"
      Printer.CurrentX = 9600 '8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
      Printer.CurrentY = 1940
      Printer.Print "備註"
   Else
      If Printer.Orientation <> 1 Then
         Printer.Orientation = 1 '直印
      End If
      '畫框
      Printer.DrawWidth = 20
      Printer.Line (800, 1660)-(11000, 15260), , B
      iLine = 1
      '定位 畫格子
      Printer.DrawWidth = 1
      For i = 1 To MaxLine
          Printer.Line (800, 1660)-(11000, 15260 - (340 * i)), , B
      Next i
      PLeft(1) = 900
      PLeft(2) = 1950
      PLeft(3) = 5950
      PLeft(4) = 8050
      PLeft(7) = 9550
      Printer.Line (1900, 1660)-(1900, 15260)
      Printer.Line (5900, 2340)-(5900, 15260)
      Printer.Line (8000, 1660)-(8000, 15260)
      Printer.Line (9500, 1660)-(9500, 15260) '備註
      Printer.Font.Size = 22
      Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("國外信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)) / 2
      Printer.CurrentY = 600
      Printer.Print "國外信件     中華民國  " & ChangeTStringToTDateString(Val(m_PrintDate) - 19110000)
      Printer.Font.Size = 14
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = 1250
      Printer.Print strDeptName
      Printer.CurrentX = 5500 '4000
      Printer.CurrentY = 1250
      Printer.Print strSignPer
      Printer.CurrentX = 300 + ((2900 - 800) / 2) - Printer.TextWidth("流水號") / 2
      Printer.CurrentY = 1940
      Printer.Print "流水號"
      Printer.CurrentX = 1300 + ((9000 - 2900) / 2) - Printer.TextWidth("公司名稱") / 2
      Printer.CurrentY = 1940
      Printer.Print "公司名稱"
      Printer.CurrentX = 7700 + ((11000 - 9000) / 2) - Printer.TextWidth("事由") / 2
      Printer.CurrentY = 1940
      Printer.Print "事由"
      Printer.CurrentX = 8900 + ((11000 - 8130) / 2) - Printer.TextWidth("收件人") / 2
      Printer.CurrentY = 1940
      Printer.Print "備註"
   End If
   
   Exit Sub
MsgErr:
    MsgBox "列印發生錯誤！" & vbCrLf & Err.Description, vbInformation, "櫃檯信件列印！"
End Sub

'Add By Sindy 2013/3/12
Sub PrintDetil3(intT1 As Integer)
On Error GoTo MsgErr
   
   If intT1 = 7 Then '未分類
      Printer.Font.Size = 12
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = 2000 + (iLine * 680) + 200
      Printer.Print strTemp(1)
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = 2000 + (iLine * 680) + 200
      Printer.Print strTemp(2)
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = 2000 + (iLine * 680) + 200
      Printer.Print strTemp(3)
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = 2000 + (iLine * 680) + 200
      Printer.Print strTemp(4)
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = 2000 + (iLine * 680) + 200
      Printer.Print strTemp(5)
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = 2000 + (iLine * 680) + 200
      Printer.Print strTemp(6)
      Printer.CurrentX = PLeft(7)
      Printer.CurrentY = 2000 + (iLine * 680) + 200
      Printer.Print strTemp(7)
   Else
      Printer.Font.Size = 12
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = 2000 + (iLine * 340) + 50
      Printer.Print strTemp(1)
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = 2000 + (iLine * 340) + 50
      Printer.Print strTemp(2)
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = 2000 + (iLine * 340) + 50
      Printer.Print strTemp(3)
      If intT1 = 1 Then '支票
         Printer.CurrentX = PLeft(3)
      Else
         Printer.CurrentX = PLeft(4)
      End If
      Printer.CurrentY = 2000 + (iLine * 340) + 50
      Printer.Print strTemp(4)
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = 2000 + (iLine * 340) + 50
      Printer.Print strTemp(5)
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = 2000 + (iLine * 340) + 50
      Printer.Print strTemp(6)
      If intT1 = 1 Then '支票
         Printer.CurrentX = 8950
      Else
         Printer.CurrentX = PLeft(7)
      End If
      Printer.CurrentY = 2000 + (iLine * 340) + 50
      Printer.Print strTemp(7)
   End If
   
   Exit Sub
MsgErr:
   MsgBox "列印發生錯誤！" & vbCrLf & Err.Description, vbInformation, "櫃檯信件列印！"
End Sub

'Add By Sindy 2013/3/12
'國外信件,系統別歸那一個專業部門負責
Private Sub GetSystemDept(intDept As Integer)
   strDept = "": strSysDept = "": strDeptName = "": strSignPer = ""
   Select Case intDept
   Case 1 'A.財務處
      strDept = "A"
      strDeptName = "部門：財務處"
      'modify by sonia 2022/10/25
      'strSignPer = "簽收人員：吳婧瑄經理　職代：辜苑琪"
      strSignPer = "簽收人員：溫斯閔　職代：吳婉莘"
   Case 2 'FCP.外專及未分類
      strDept = "FCP"
      strSysDept = "<FCP>,<FG>"
      strDeptName = "部門：外專"
      strSignPer = "簽收人員：顏裕洋副理　職代："
   Case 3 'P.專利處
      strDept = "P"
      strSysDept = "<P>,<PS>,<CFP>,<CPS>"
      strDeptName = "部門：專利處"
      strSignPer = "簽收人員：陳玫音主任　職代：林禧佩"
   Case 4 'T.商標處
      strDept = "T"
      strSysDept = "<T>,<TB>,<TC>,<TD>,<TF>,<TM>,<TR>,<TS>,<TT>"
      strDeptName = "部門：商標處"
      'Modified by Lydia 2021/11/10 內商林經理退休：林純貞(姓名)－＞林承慧。
      strSignPer = "簽收人員：林承慧副理　職代："
   Case 5 'FCT.外商
      strDept = "FCT"
      strSysDept = "<CFT>,<CFC>,<FCT>,<S>"
      strDeptName = "部門：外商"
      'Modified by Morgan 2021/8/10
      'strSignPer = "簽收人員：陳鳳英經理　職代：洪琬姿經理"
      strSignPer = "簽收人員：洪琬姿經理　職代：沈佳穎副理"
   Case 6 'FCL.投資法務
      strDept = "FCL"
      strSysDept = "<CFL>,<FCL>,<LIN>"
      strDeptName = "部門：投資法務"
      strSignPer = "簽收人員：王麗真主任　職代："
   Case 7 '未分類
      strDept = "X"
      strDeptName = "部門：外專"
      strSignPer = "簽收人員：顏裕洋副理　職代："
   End Select
End Sub
