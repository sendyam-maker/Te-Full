VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170002 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他所得/扣款資料"
   ClientHeight    =   5076
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8364
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5076
   ScaleWidth      =   8364
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
            Picture         =   "frm170002.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170002.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   528
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8364
      _ExtentX        =   14753
      _ExtentY        =   931
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
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   45
      TabIndex        =   12
      Top             =   660
      Width           =   8250
      _ExtentX        =   14542
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170002.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblOC03"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDsp(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDsp(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label14"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(9)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(10)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(12)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCUID"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtNote"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtOD(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtOD(3)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtOD(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtOD(4)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtOD(5)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtOD(6)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtOD(13)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtNHI10"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtNet"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtOD(14)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Combo1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtOD(16)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtSM42"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "複製資料"
      TabPicture(1)   =   "frm170002.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCopy"
      Tab(1).Control(1)=   "txtCopy(2)"
      Tab(1).Control(2)=   "txtCopy(1)"
      Tab(1).Control(3)=   "txtCopy(0)"
      Tab(1).Control(4)=   "lblCopy(3)"
      Tab(1).Control(5)=   "lblCopy(2)"
      Tab(1).Control(6)=   "lblCopy(4)"
      Tab(1).Control(7)=   "lblCopy(1)"
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(9)=   "Label8"
      Tab(1).Control(10)=   "Label7"
      Tab(1).Control(11)=   "Label6"
      Tab(1).Control(12)=   "Label5"
      Tab(1).Control(13)=   "Label4"
      Tab(1).Control(14)=   "Label2"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "多筆瀏覽"
      TabPicture(2)   =   "frm170002.frx":212C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt1(0)"
      Tab(2).Control(1)=   "txt1(1)"
      Tab(2).Control(2)=   "txt1(2)"
      Tab(2).Control(3)=   "txt1(3)"
      Tab(2).Control(4)=   "cmdok"
      Tab(2).Control(5)=   "GRD1"
      Tab(2).Control(6)=   "lblQuery"
      Tab(2).Control(7)=   "Label12(0)"
      Tab(2).Control(8)=   "Line1"
      Tab(2).Control(9)=   "Label11"
      Tab(2).Control(10)=   "Line2"
      Tab(2).ControlCount=   11
      Begin VB.TextBox txtSM42 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   3690
         MaxLength       =   8
         TabIndex        =   60
         Text            =   "8888888"
         Top             =   3420
         Width           =   915
      End
      Begin VB.TextBox txtOD 
         Height          =   270
         Index           =   16
         Left            =   6930
         MaxLength       =   3
         TabIndex        =   59
         Top             =   1290
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frm170002.frx":2148
         Left            =   5265
         List            =   "frm170002.frx":214A
         Style           =   2  '單純下拉式
         TabIndex        =   58
         Top             =   1275
         Width           =   1590
      End
      Begin VB.TextBox txtOD 
         Height          =   270
         Index           =   14
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "10410"
         Top             =   3420
         Width           =   735
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   1575
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "8888888"
         Top             =   3150
         Width           =   915
      End
      Begin VB.TextBox txtNHI10 
         Height          =   270
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "120000"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtOD 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Index           =   13
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "2000"
         Top             =   2610
         Width           =   885
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "複製"
         Height          =   375
         Left            =   -70950
         TabIndex        =   23
         Top             =   450
         Width           =   915
      End
      Begin VB.TextBox txtOD 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   6
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   5
         Text            =   "2000"
         Top             =   2340
         Width           =   885
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73890
         MaxLength       =   6
         TabIndex        =   35
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   36
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70950
         MaxLength       =   7
         TabIndex        =   37
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69960
         MaxLength       =   7
         TabIndex        =   38
         Top             =   405
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   375
         Left            =   -68580
         TabIndex        =   39
         Top             =   353
         Width           =   915
      End
      Begin VB.TextBox txtCopy 
         Height          =   285
         Index           =   2
         Left            =   -73380
         TabIndex        =   22
         Text            =   "960531"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtCopy 
         Height          =   285
         Index           =   1
         Left            =   -72150
         MaxLength       =   8
         TabIndex        =   21
         Text            =   "12345678"
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox txtCopy 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   -73380
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "12345678"
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox txtOD 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   5
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   4
         Text            =   "2000"
         Top             =   1830
         Width           =   885
      End
      Begin VB.TextBox txtOD 
         Height          =   270
         Index           =   4
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "1"
         Top             =   1290
         Width           =   345
      End
      Begin VB.TextBox txtOD 
         Height          =   270
         Index           =   2
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   1
         Text            =   "960501"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtOD 
         Height          =   270
         Index           =   3
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "999999"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtOD 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "12345678"
         Top             =   420
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170002.frx":214C
         Height          =   3255
         Left            =   -74880
         TabIndex        =   40
         Top             =   780
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5736
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "流水號　|日期　　　|員工　　　|所得/扣款代號|所得/扣款類別|所得/扣款說明|金額　　　　|所得稅金|補充保費"
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
         _Band(0).Cols   =   9
      End
      Begin MSForms.TextBox txtNote 
         Height          =   300
         Left            =   1575
         TabIndex        =   10
         Top             =   3690
         Width           =   3930
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "6932;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   360
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4020
         Width           =   6540
         VariousPropertyBits=   671105051
         Size            =   "11536;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "健保級數："
         Height          =   180
         Index           =   12
         Left            =   2760
         TabIndex        =   61
         Top             =   3465
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "健保對象："
         Height          =   180
         Index           =   11
         Left            =   4320
         TabIndex        =   57
         Top             =   1335
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "備　　註："
         Height          =   180
         Index           =   10
         Left            =   645
         TabIndex        =   56
         Top             =   3735
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "補扣/退繳年月："
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   55
         Top             =   3450
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "給付淨額："
         Height          =   180
         Index           =   9
         Left            =   645
         TabIndex        =   54
         Top             =   3195
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付時間：                     (格式：HHMMSS 例：)"
         Height          =   180
         Index           =   8
         Left            =   645
         TabIndex        =   53
         Top             =   2925
         Width           =   3630
      End
      Begin VB.Label Label14 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "補充保費："
         Height          =   180
         Left            =   645
         TabIndex        =   52
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label lblQuery 
         AutoSize        =   -1  'True
         Caption         =   "共              筆"
         Height          =   180
         Left            =   -68205
         TabIndex        =   51
         Top             =   4110
         Width           =   990
      End
      Begin VB.Label lblCopy 
         Alignment       =   1  '靠右對齊
         Caption         =   "999"
         Height          =   180
         Index           =   3
         Left            =   -73050
         TabIndex        =   50
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label lblCopy 
         Alignment       =   1  '靠右對齊
         Caption         =   "999"
         Height          =   180
         Index           =   2
         Left            =   -73050
         TabIndex        =   49
         Top             =   1770
         Width           =   660
      End
      Begin VB.Label lblCopy 
         Alignment       =   1  '靠右對齊
         Caption         =   "999"
         Height          =   180
         Index           =   4
         Left            =   -73050
         TabIndex        =   48
         Top             =   2310
         Width           =   660
      End
      Begin VB.Label lblCopy 
         Alignment       =   1  '靠右對齊
         Caption         =   "999"
         Height          =   180
         Index           =   1
         Left            =   -73050
         TabIndex        =   47
         Top             =   1500
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "所得稅金："
         Height          =   180
         Index           =   7
         Left            =   645
         TabIndex        =   46
         Top             =   2370
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   180
         Left            =   2070
         TabIndex        =   45
         Top             =   2130
         Width           =   135
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         Caption         =   "0.0"
         Height          =   180
         Index           =   4
         Left            =   1635
         TabIndex        =   44
         Top             =   2130
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "所得稅率："
         Height          =   180
         Index           =   6
         Left            =   645
         TabIndex        =   43
         Top             =   2130
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   -74820
         TabIndex        =   42
         Top             =   450
         Width           =   900
      End
      Begin VB.Line Line1 
         X1              =   -73170
         X2              =   -72570
         Y1              =   533
         Y2              =   533
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71550
         TabIndex        =   41
         Top             =   450
         Width           =   540
      End
      Begin VB.Line Line2 
         X1              =   -70320
         X2              =   -69570
         Y1              =   533
         Y2              =   533
      End
      Begin VB.Label lblDsp 
         AutoSize        =   -1  'True
         Caption         =   "A"
         Height          =   180
         Index           =   2
         Left            =   1575
         TabIndex        =   34
         Top             =   1620
         Width           =   120
      End
      Begin VB.Label lblOC03 
         AutoSize        =   -1  'True
         Caption         =   "說明文字"
         Height          =   180
         Left            =   1980
         TabIndex        =   33
         Top             =   1335
         Width           =   2130
      End
      Begin MSForms.Label lblName 
         Height          =   285
         Left            =   2400
         TabIndex        =   31
         Top             =   1080
         Width           =   1050
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1852;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "( A:加項,D:減項 )"
         Height          =   180
         Left            =   1785
         TabIndex        =   30
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label9 
         Caption         =   "　　　扣款金額：                      元"
         Height          =   180
         Left            =   -74670
         TabIndex        =   29
         Top             =   2310
         Width           =   2610
      End
      Begin VB.Label Label8 
         Caption         =   "　複製扣款筆數：                      筆"
         Height          =   180
         Left            =   -74670
         TabIndex        =   28
         Top             =   2040
         Width           =   2610
      End
      Begin VB.Label Label7 
         Caption         =   "　　　所得金額：                      元"
         Height          =   180
         Left            =   -74670
         TabIndex        =   27
         Top             =   1770
         Width           =   2610
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "共複製所得筆數：                      筆"
         Height          =   180
         Left            =   -74670
         TabIndex        =   26
         Top             =   1500
         Width           =   2610
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "複製日期："
         Height          =   180
         Left            =   -74340
         TabIndex        =   25
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -72390
         TabIndex        =   24
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "欲複製之流水號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   19
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "金　　額："
         Height          =   180
         Index           =   4
         Left            =   645
         TabIndex        =   18
         Top             =   1875
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "所得/扣款類別："
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1605
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "所得/扣款代號："
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1335
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "日　　期："
         Height          =   180
         Index           =   1
         Left            =   645
         TabIndex        =   15
         Top             =   735
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   645
         TabIndex        =   14
         Top             =   1035
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "流  水  號："
         Height          =   180
         Left            =   645
         TabIndex        =   13
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "frm170002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2008/12/19
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_OD As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean
Dim m_taxrate As String '所得稅率
'Added by Morgan 2013/1/24
Dim stNHI() As String
Dim m_arrOD() As String


Private Sub cmdCopy_Click()
   If CopyCheck = True Then
      If CopyData = True Then
         MsgBox "資料複製完成!!"
      End If
   End If
End Sub

Private Function CopyCheck() As Boolean
   Dim bCancel As Boolean
   If txtCopy(0) = "" Then
      txtCopy(0).SetFocus
      MsgBox "流水號起不可空白!"
      Exit Function
   ElseIf txtCopy(1) = "" Then
      txtCopy(1).SetFocus
      MsgBox "流水號迄不可空白!"
      Exit Function
   ElseIf txtCopy(2) = "" Then
      MsgBox "複製日期不可空白!"
      txtCopy(2).SetFocus
      Exit Function
   End If
   
   For Each oText In txtCopy
      txtCopy_Validate oText.Index, bCancel
      If bCancel = True Then
         Exit Function
      End If
   Next
   
   'Modified by Morgan 2013/2/4 +檢查欲複製的資料內不可含有01,02的所得/扣款代號的資料
   strExc(0) = "select count(*) c1,nvl(sum(decode(OC02,'A',1)),0) a1,nvl(sum(decode(OC02,'A',od05)),0) a2" & _
      ",nvl(sum(decode(OC02,'D',1)),0) d1,nvl(sum(decode(OC02,'D',od05)),0) d2,sum(decode(od04,'01',1,'02',1)) d3" & _
      " from OtherSalaryData,OtherSalaryCode" & _
      " where od01>=" & txtCopy(0) & " and od01<=" & txtCopy(1) & " and oc01(+)=od04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         If RsTemp.Fields("d3") > 0 Then
            MsgBox " 欲複製的流水號含有所得/扣款代號為 01 或 02 的資料，不可複製!", vbExclamation
            Exit Function
         Else
            lblCopy(1) = RsTemp.Fields("a1")
            lblCopy(2) = RsTemp.Fields("a2")
            lblCopy(3) = RsTemp.Fields("d1")
            lblCopy(4) = RsTemp.Fields("d2")
            If MsgBox("預計複製 " & RsTemp.Fields("c1") & " 筆資料，是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbYes Then
               CopyCheck = True
            End If
         End If
      Else
         MsgBox "無資料可複製!"
      End If
   End If
End Function

Private Function CopyData() As Boolean
   Dim stSQL As String, SNo As String
   Dim iYear As Integer
   
   iYear = Val(txtCopy(2)) \ 10000 + 1911
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   stSQL = "select nvl(max(od01)," & (iYear - 1911) & "00000) sn from OtherSalaryData where substr(od02,1,4)=" & iYear
   If RsTemp.State <> 0 Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
   SNo = Val("" & RsTemp.Fields(0))
   stSQL = "insert into OtherSalaryData (od01,od02,od03,od04,od05,od06)" & _
      " select " & SNo & "+rownum," & (Val(txtCopy(2)) + 19110000) & ",od03,od04,od05,od06" & _
      " from ( select od01,od03,od04,od05,od06 from OtherSalaryData" & _
      " where od01>=" & Val(txtCopy(0)) & " and od01<=" & Val(txtCopy(1)) & " group by od01,od03,od04,od05,od06) x"
   
   cnnConnection.Execute stSQL, intI
   
   
   stSQL = "select nvl(sum(decode(OC02,'A',1)),0) a1,nvl(sum(decode(OC02,'A',od05)),0) a2" & _
      ",nvl(sum(decode(OC02,'D',1)),0) d1,nvl(sum(decode(OC02,'D',od05)),0) d2" & _
      " from OtherSalaryData,OtherSalaryCode" & _
      " where od01>" & SNo & " and substr(od02,1,4)=" & iYear & " and oc01(+)=od04 "
   
   If RsTemp.State <> 0 Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
   lblCopy(1) = RsTemp.Fields("a1")
   lblCopy(2) = RsTemp.Fields("a2")
   lblCopy(3) = RsTemp.Fields("d1")
   lblCopy(4) = RsTemp.Fields("d2")

   cnnConnection.CommitTrans
   CopyData = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Sub cmdok_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
       If RunNick(txt1(0), txt1(1)) Then
           txt1(0).SetFocus
           Exit Sub
       End If
       If RunNick2(txt1(2), txt1(3)) Then
           txt1(2).SetFocus
           Exit Sub
       End If
       GetData
   Else
       MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
       txt1(0).SetFocus
   End If
End Sub

Sub GetData()
   Dim stCon As String
   stCon = ""
   If txt1(0) <> "" Then
       stCon = stCon & " and od03>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       stCon = stCon & " and od03<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       stCon = stCon & " and od02>='" & DBDATE(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then
       stCon = stCon & " and od02<='" & DBDATE(txt1(3)) & "' "
   End If
   strExc(0) = "SELECT od01,sqldateT(od02),st02,od04,decode(oc02,'A','加項','D','減項',oc02),oc03,od05,trunc(od06),od13 FROM othersalarydata,staff,OtherSalaryCode" & _
      " where oc01=od04(+) and st01(+)=od03 " & stCon & " order by od01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
      lblQuery = "共 " & RsTemp.RecordCount & " 筆"
      'Added by Morgan 2013/2/4
      '縮小所得/扣款代號,類別欄寬以便顯示所得稅金,補充保費
      GRD1.ColWidth(3) = 470
      GRD1.ColAlignmentFixed(3) = 6
      GRD1.ColAlignment(3) = 3
      GRD1.ColWidth(4) = 470
      GRD1.ColAlignmentFixed(4) = 6
      GRD1.ColAlignment(4) = 3
      'end 2013/2/4
   End If
End Sub

Private Sub Combo1_Click()
   If Combo1.ListIndex >= 0 Then
      txtOD(16) = Combo1.ItemData(Combo1.ListIndex)
   Else
      txtOD(16) = ""
   End If
End Sub

Private Sub Form_Activate()
   If m_bActived = False Then
      SetInputEntry
      m_bActived = True
      SSTab1.Tab = 0
   End If
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
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
   
   ClearField1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170002 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from OtherSalaryData where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_OD = .Fields.Count
      
      ReDim m_FieldList(TF_OD) As FIELDITEM
      
      For Each oText In txtOD
         idx = oText.Index
         m_FieldList(idx).fiName = "OD" & Format(idx, "00")
         'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
         'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
            m_FieldList(idx).fiType = 0
         'Else
         '   m_FieldList(idx).fiType = 1
         'End If
         'end 2017/06/29
      Next
      End With
   End If
   
   'Added by Morgan 2013/1/24
   ReDim m_arrOD(TF_OD) As String
   ReDim stNHI(TF_NHI) As String
   'end 2013/1/24
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim stKey01 As String
   Dim adoRst As New ADODB.Recordset
   
   stKey01 = txtOD(1)
   'Modified by Morgan 2013/1/29 +nhi2nd
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM OtherSalaryData,nhi2nd" & _
            " WHERE od01 = '" & stKey01 & "' and nhi01(+)=od03 and nhi02(+)=od02 and nhi03(+)=decode(od04,'01','50','02','9A') and nhi04(+)=decode(od04,'01','4','02','5')"
      Case -2
         strExc(0) = "SELECT * FROM OtherSalaryData,nhi2nd where nhi01(+)=od03 and nhi02(+)=od02 and nhi03(+)=decode(od04,'01','50','02','9A') and nhi04(+)=decode(od04,'01','4','02','5') order by 1 ASC"
      Case -1
         strExc(0) = "SELECT * FROM OtherSalaryData,nhi2nd" & _
            " WHERE od01 <'" & stKey01 & "' and nhi01(+)=od03 and nhi02(+)=od02 and nhi03(+)=decode(od04,'01','50','02','9A') and nhi04(+)=decode(od04,'01','4','02','5') order by 1 DESC"
      Case 1
         strExc(0) = "SELECT * FROM OtherSalaryData,nhi2nd" & _
            " WHERE od01 >'" & stKey01 & "' and nhi01(+)=od03 and nhi02(+)=od02 and nhi03(+)=decode(od04,'01','50','02','9A') and nhi04(+)=decode(od04,'01','4','02','5') order by 1 ASC"
      Case 2
         strExc(0) = "SELECT * FROM OtherSalaryData,nhi2nd where nhi01(+)=od03 and nhi02(+)=od02 and nhi03(+)=decode(od04,'01','50','02','9A') and nhi04(+)=decode(od04,'01','4','02','5') order by 1 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtOD(1).SetFocus
      txtOD_GotFocus 1
   End If
End Function

Private Sub GRD1_Click()
   Dim lCurRow As Long, i As Integer, j As Integer
   lCurRow = GRD1.row
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If GRD1.CellBackColor <> &HFFC0C0 Then
            GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
               GRD1.row = j
               If GRD1.CellBackColor <> QBColor(15) Then
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                  Next i
               End If
            Next j
            GRD1.row = lCurRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            GRD1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub GRD1_DblClick()
   Dim lCurRow As Long
   lCurRow = GRD1.row
   '呼叫查詢
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
            If txtOD(1).Locked = False Then
               txtOD(1).Text = GRD1.TextMatrix(lCurRow, 0)
               If TBar1.Buttons(11).Enabled = True Then
                  Call Tbar1_ButtonClick(TBar1.Buttons(11))
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 1 Then
      txtCopy(0).SetFocus
      TextInverse txtCopy(0)
   ElseIf SSTab1.Tab = 2 Then
      txt1(0).SetFocus
      TextInverse txt1(0)
   ElseIf SSTab1.Tab = 0 And PreviousTab = 2 Then
      GRD1_DblClick
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   CloseIme
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCopy_GotFocus(Index As Integer)
   TextInverse txtCopy(Index)
   CloseIme
End Sub

Private Sub txtCopy_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCopy_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Then
      If txtCopy(Index) <> "" Then
         If ChkDate(txtCopy(Index)) = False Then
            Cancel = True
         End If
      End If
   End If
End Sub

'Added by Morgan 2013/1/29
Private Sub txtNHI10_GotFocus()
   TextInverse txtNHI10
   CloseIme
End Sub

Private Sub txtNHI10_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtNHI10_Validate(Cancel As Boolean)
   If txtNHI10.Tag <> txtNHI10 Then
      SetNHI06
   End If
   txtNHI10.Tag = txtNHI10
End Sub
'end 2013/1/29

Private Sub txtOD_Change(Index As Integer)
   Select Case Index
      Case 3
         If txtOD(Index) = "" Then
            lblName.Caption = ""
         End If
         setCombo1 'Added by Morgan 2015/12/14
         
      Case 4
         If txtOD(Index) = "" Then
            lblDsp(2) = ""
            lblOC03 = "" 'Modify By Sindy 2021/12/22
         End If
         setCombo1 'Added by Morgan 2015/12/14
         
   End Select
End Sub

'Added by Morgan 2015/12/14
Private Sub setCombo1()
   If (txtOD(4) = "35" Or txtOD(4) = "36") Then
      If txtOD(3) <> "" And Combo1.Tag <> txtOD(3) Then
         
         strExc(0) = "select sr04||'   ('||decode(sr03,'1','父親','2','母親','3','配偶','4','子女','其他')||')' c01,sr02 from staff_relation where sr01='" & txtOD(3) & "' union all select st02||'   (自己)' c01,0 sr02 from staff where st01='" & txtOD(3) & "' order by sr02 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
               Combo1.AddItem RsTemp(0), 0
               Combo1.ItemData(0) = RsTemp("sr02")
               RsTemp.MoveNext
            Loop
         End If
         Combo1.Tag = txtOD(3)
      End If
      
      If txtOD(16) <> "" Then
         For intI = 0 To Combo1.ListCount - 1
            If Combo1.ItemData(intI) = Val(txtOD(16)) Then
               Combo1.ListIndex = intI
               Exit For
            End If
         Next
      End If
         
      If m_EditMode = 1 Or m_EditMode = 2 Then
         Combo1.Enabled = True
         txtSM42.Enabled = True 'Added by Morgan 2019/9/2
      Else
         Combo1.Enabled = False
         txtSM42.Enabled = False 'Added by Morgan 2019/9/2
      End If
   Else
      Combo1.Clear
      Combo1.Enabled = False
      Combo1.Tag = ""
      txtOD(16) = ""
      'Added by Morgan 2019/9/2
      txtSM42 = ""
      txtSM42.Tag = txtSM42 'Added by Morgan 2020/10/26
      txtSM42.Enabled = False
      'end 2019/9/2
   End If
End Sub

Private Sub txtOD_GotFocus(Index As Integer)
   TextInverse txtOD(Index)
   CloseIme
End Sub

Private Sub ClearField()
   For Each oText In txtOD
      oText.Text = Empty
   Next
   txtNHI10 = "" 'Added by Morgan 2013/1/29
   txtSM42 = "" 'Added by Morgan 2019/9/2
   
   'Add By Sindy 2021/12/20
   lblName.Caption = Empty
   txtNote.Text = Empty: txtNote.Tag = Empty
   lblOC03.Caption = Empty
   '2021/12/20 END
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_OD
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   
   'Added by Morgan 2013/1/24
   Erase m_arrOD
   ReDim m_arrOD(TF_OD) As String
   'end 2013/1/24
   
   m_bConfirmCheck = False
End Sub

Private Sub ClearField1()
   For Each oText In txtCopy
      oText.Text = ""
   Next
   For Each oLabel In lblCopy
      oLabel.Caption = Empty
   Next
   
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtOD
         idx = oText.Index
         m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         '日期轉民國
         If idx = 2 Then
            oText.Text = TransDate(m_FieldList(idx).fiOldData, 1)
         ElseIf idx = 14 Then
            If m_FieldList(idx).fiOldData <> "" Then
               oText.Text = m_FieldList(idx).fiOldData - 191100
            Else
               oText.Text = ""
            End If
         Else
            oText.Text = m_FieldList(idx).fiOldData
         End If
         m_arrOD(idx) = oText.Text 'Added by Morgan 2013/1/24
      Next
      
      'Add By Sindy 2021/12/20
      txtNote.Text = "" & .Fields("od15"): txtNote.Tag = "" & .Fields("od15")
      '2021/12/20 END
      
      If ClsPDGetStaffN(txtOD(3), strExc(1)) Then
         lblName.Caption = strExc(1)
      End If
      
      txtNHI10 = "" & .Fields("nhi10") 'Added by Morgan 2013/1/29
      
      setCombo1 'Added by Morgan 2015/12/14
      
      If txtOD(4) <> "" Then
         SetRefData txtOD(4)
      End If
      
      CUID(1) = "" & .Fields("od07")
      CUID(2) = "" & .Fields("od08")
      CUID(3) = "" & .Fields("od09")
      CUID(4) = "" & .Fields("od10")
      CUID(5) = "" & .Fields("od11")
      CUID(6) = "" & .Fields("od12")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtOD(1).Tag = txtOD(1)
   SetNet 'Added by Morgan 2013/2/22
   
   'Added by Morgan 2019/9/2
   If (txtOD(4) = "35" Or txtOD(4) = "36") Then
      strExc(0) = "select sm42 from salarymonth where sm01='" & txtOD(3) & "' and sm02=" & (Val(txtOD(14)) + 191100)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtSM42 = "" & RsTemp(0)
         txtSM42.Tag = txtSM42
      End If
   End If
   'end 2019/9/2
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtOD
      oText.Locked = bLocked
   Next
   txtNHI10.Locked = bLocked 'Added by Morgan 2013/1/29
   txtNote.Locked = bLocked 'Modify By Sindy 2021/12/20
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
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

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         SSTab1.Tab = 0
         m_EditMode = 1
         ClearField
         SetInputEntry
         UpdateToolbarState

      Case vbKeyF3 ' 修改
         SSTab1.Tab = 0
         'Added by Morgan 2013/2/5
         If PUB_ExistsSalaryMonth(txtOD(2)) = True Then
            MsgBox "該筆資料之月薪資已計算不可修改！", vbExclamation
            Exit Sub
         End If
         'end 2013/2/5
         m_EditMode = 2
         SetInputEntry
         UpdateToolbarState

      Case vbKeyF5 ' 刪除
         SSTab1.Tab = 0
         'Added by Morgan 2013/2/5
         If PUB_ExistsSalaryMonth(txtOD(2)) = True Then
            MsgBox "該筆資料之月薪資已計算不可刪除！", vbExclamation
            Exit Sub
         End If
         'end 2013/2/5
         
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = False Then
               Exit Sub
            End If
            UpdateToolbarState
         End If
         
      Case vbKeyF4 ' 查詢
         SSTab1.Tab = 0
         m_EditMode = 4
         ClearField
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = False Then
            Exit Sub
         End If
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyF10 ' 取消
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtOD(1) = txtOD(1).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
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
         If m_bUpdate And txtOD(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtOD(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtOD(1) <> "" Then
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

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         SetCtrlReadOnly False
         txtOD(1).Locked = True
         If Me.Visible = True Then
            txtOD(2).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
      Case 2
         SetCtrlReadOnly False
         txtOD(1).Locked = True
         txtOD(2).Locked = True
         txtNHI10.Locked = True 'Added by Morgan 2013/1/31
         If Me.Visible = True Then
            txtOD(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
      Case 4
         SetCtrlReadOnly True
         txtOD(1).Locked = False
         If Me.Visible = True Then
            txtOD(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
      Case Else
         SetCtrlReadOnly True
         If Me.Visible = True Then
            txtOD(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = True
         SSTab1.TabEnabled(2) = True
   End Select
   PUB_ChangeCaption Me, m_EditMode
   setCombo1 'Added by Morgan 2015/12/14
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
               txtOD(1).SetFocus
               txtOD_GotFocus 1
            End If
         End If
         
   End Select
End Function


Private Function TxtValidate() As Boolean
   
   Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   '查詢
   If m_EditMode = 4 Then
      If txtOD(1) = "" Then
         ShowMsg "請輸入流水號 !"
         txtOD(1).SetFocus
         txtOD_GotFocus 1
         GoTo EscPoint
      End If
   '維護
   Else
   
      For Each oText In txtOD
         If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
            idx = oText.Index
            bCancel = False
            txtOD_Validate idx, bCancel
            If bCancel = True Then
               txtOD(idx).SetFocus
               txtOD_GotFocus idx
               GoTo EscPoint
            End If
         End If
      Next
   
      If txtOD(2) = "" And txtOD(2).Locked = False Then
         ShowMsg "請輸入日期 !"
         txtOD(2).SetFocus
         txtOD_GotFocus 2
         GoTo EscPoint
      End If
      
      If txtOD(3) = "" And txtOD(3).Locked = False Then
         ShowMsg "請輸入員工代號 !"
         txtOD(3).SetFocus
         txtOD_GotFocus 3
         GoTo EscPoint
      End If
      If txtOD(4) = "" And txtOD(4).Locked = False Then
         ShowMsg "請輸入所得/扣款代號 !"
         txtOD(4).SetFocus
         txtOD_GotFocus 4
         GoTo EscPoint
      End If
      
      'Add by Morgan 2011/1/4
      If txtOD(4) = "01" And Left(txtOD(3), 1) <> "F" Then
         ShowMsg lblOC03 & "的員工編號必須是外譯編號!" 'Modify By Sindy 2021/12/22 + lblOC03
         txtOD(3).SetFocus
         txtOD_GotFocus 3
         GoTo EscPoint
      End If
      
      If txtOD(5) = "" And txtOD(5).Locked = False Then
         ShowMsg "請輸入金額 !"
         txtOD(5).SetFocus
         txtOD_GotFocus 5
         GoTo EscPoint
      End If
      
      'Added by Morgan 2013/3/21
      If m_EditMode = "1" Then
         If txtOD(4) = "02" Then
            MsgBox "此處不可新增複委託，請改至【其他各類所得資料(平日)】輸入！", vbExclamation
            GoTo EscPoint
         End If
         'Added by Morgan 2015/12/10
         If txtOD(4) = "05" Then
            MsgBox "不可再新增 05退保費！", vbExclamation
            GoTo EscPoint
         End If
         'end 2015/12/10
      End If
      'end 2013/3/21
      
      'Added by Morgan 2013/1/24
      If txtOD(4) = "01" Or txtOD(4) = "02" Then
         strExc(0) = "select * from OtherSalaryData where od02=" & DBDATE(txtOD(2)) & " and od03='" & txtOD(3) & "' and od04='" & txtOD(4) & "'"
         If txtOD(1) <> "" Then
            strExc(0) = strExc(0) & " and od01<>" & txtOD(1)
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "一員工一日不可有兩筆" & lblOC03 & "!", vbExclamation 'Modify By Sindy 2021/12/22 + lblOC03
            GoTo EscPoint
         End If
         If txtNHI10 = "" And txtNHI10.Locked = False Then
            ShowMsg "所得代碼為 01 或 02 時，請輸入給付時間以便計算補充保費 !"
            txtNHI10.SetFocus
            txtNHI10_GotFocus
            GoTo EscPoint
         End If
      End If
      'end 2013/1/24
      
      'Added by Morgan 2015/12/10
      'Modified by Morgan 2022/7/28
      'If txtOD(4) >= "31" And txtOD(4) <= "40" Then
      If txtOD(4) >= "31" And txtOD(4) <= "36" Then
      'end 2022/7/28
         If txtOD(14) = "" Then
            MsgBox "請輸入補扣/退繳年月！", vbExclamation
            txtOD(14).SetFocus
            GoTo EscPoint
         End If
         
         If txtOD(16) = "" Then
            'Modified by Morgan 2022/7/28
            'If txtOD(4) >= "37" And txtOD(4) <= "38" Then
            If txtOD(4) = "35" Or txtOD(4) = "36" Then
            'end 2022/7/28
               MsgBox "請點選健保對象！", vbExclamation
               GoTo EscPoint
            End If
         End If
         
         'Added by Morgan 2019/9/2
         If txtSM42.Enabled Then
            If txtSM42 = "" Then
               MsgBox "請輸入健保級數！", vbExclamation
               txtSM42.SetFocus
               GoTo EscPoint
            Else
               strExc(0) = "select * from SalaryInsurance where si01='H' and si02=" & Val(txtSM42)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI <> 1 Then
                  MsgBox "健保級數輸入錯誤！", vbExclamation
                  txtSM42.SetFocus
                  GoTo EscPoint
               End If
            End If
         End If
         'end 2019/9/2
      End If
      'end 2015/12/10
   End If
   
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
   Dim stCols As String, stValues As String, stSQL As String
   Dim stYear As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   '翻譯費
   If txtOD(4) = "01" Then
      SetNHI06 '此處要重算以避免輸入過程中同時有新增較早資料而沒算到
      '檢查內翻人員不可有晚於該筆資料的補充保費
      If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10), True) = False Then
         GoTo EscPoint
      End If
   End If
   'end 2013/2/6
   
   stYear = DBDATE(txtOD(2)) \ 10000
   
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtOD
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
   'Modify By Sindy 2021/12/20 + txtNote
   stCols = Mid(stCols & ",OD15", 2)
   stValues = Mid(stValues & "," & CNULL(ChgSQL(txtNote.Text)), 2)
   stSQL = "declare intMax number;begin select max(OD01)+1 into intMax from othersalarydata where substr(od02,1,4)=" & stYear & ";IF intMax IS NULL THEN intMax:=" & (stYear - 1911) & "00001; END IF;"
   stSQL = stSQL & "INSERT INTO OtherSalaryData (OD01," & stCols & ") Values (intMax," & stValues & ");end;"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   'Added by Morgan 2013/1/24
   '翻譯費,執行業務所得要扣補充保費
   If txtOD(4) = "01" Or txtOD(4) = "02" Then
      PUB_InsertNHI2nd stNHI
   End If
   'end 2013/1/24
   
   stSQL = "select max(OD01) from othersalarydata where substr(od02,1,4)=" & stYear
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      txtOD(1) = RsTemp.Fields(0)
   End If
   
   UpdateSM42 'Added by Morgan 2019/9/2
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

EscPoint:

End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   '翻譯費
   If txtOD(4) = "01" Then
      SetNHI06 'Added by Morgan 2024/11/8
      '檢查內翻人員不可有晚於該筆資料的補充保費
      If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10)) = False Then
         GoTo EscPoint
      End If
   End If
   'end 2013/2/6
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE OtherSalaryData SET "
   stSet = ""
   For Each oText In txtOD
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
   
   'Modify By Sindy 2021/12/20 + txtNote
   If txtNote.Text <> txtNote.Tag Then
      stSet = stSet & ",OD15=" & CNULL(ChgSQL(txtNote.Text))
   End If
   If bDifference = True Or txtNote.Text <> txtNote.Tag Then
   '2021/12/20 END
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where od01='" & txtOD(1) & "'; end; "
      Pub_SeekTbLog stSQL
      
      cnnConnection.Execute stSQL, intI
      
      'Added by Morgan 2013/1/24
      '翻譯費,執行業務所得要扣補充保費
      If txtOD(4) = "01" Or txtOD(4) = "02" Then
         PUB_InsertNHI2nd stNHI
      ElseIf m_FieldList(4).fiOldData = "01" Then
         strSql = "DELETE NHI2ND WHERE NHI01='" & m_FieldList(3).fiOldData & "' AND NHI02=" & DBDATE(m_FieldList(2).fiOldData) & " AND NHI03='50' AND NHI04='4'"
         cnnConnection.Execute strSql, intI
      ElseIf m_FieldList(4).fiOldData = "02" Then
         strSql = "DELETE NHI2ND WHERE NHI01='" & m_FieldList(3).fiOldData & "' AND NHI02=" & DBDATE(m_FieldList(2).fiOldData) & " AND NHI03='9A' AND NHI04='5'"
         cnnConnection.Execute strSql, intI
      End If
      'end 2013/1/24
   End If
   
   UpdateSM42 'Added by Morgan 2019/9/2
   
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

EscPoint:

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

Private Sub UpdateFieldNewData()
   For Each oText In txtOD
      idx = oText.Index
      Select Case idx
         Case 2
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case 14
            If oText.Text <> "" Then
               m_FieldList(idx).fiNewData = Val(oText.Text) + 191100
            Else
               m_FieldList(idx).fiNewData = ""
            End If
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtOD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 3, 15
         
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtOD_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 2
            If txtOD(Index) <> "" And Not txtOD(Index).Locked Then
               If ChkDate(txtOD(Index)) = False Then
                  Cancel = True
               'Added by Morgan 2013/2/5
               ElseIf txtOD(4) = "01" And Val(txtOD(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "翻譯費所得日期不可晚於系統日！", vbExclamation
                  Cancel = True
               ElseIf PUB_ExistsSalaryMonth(txtOD(2)) = True Then
                  MsgBox "該日期之月薪資已計算不可再新增！", vbExclamation
                  Cancel = True
               'end 2013/2/5
               End If
            End If
         Case 3
            If txtOD(Index) <> "" Then
               If ChkStaffID(txtOD(Index)) = True Then
                  Cancel = True
               End If
               If ClsPDGetStaffN(txtOD(Index), strExc(1)) = False Then
                  Cancel = True
               Else
                  lblName.Caption = strExc(1)
               End If
               
               If Cancel = False Then setCombo1 'Added by Morgan 2015/12/14
            End If
            
         Case 4
            If txtOD(Index) <> "" Then
               'Added by Morgan 2013/3/21
               If m_EditMode = "1" And txtOD(Index) = "02" Then
                  MsgBox "此處不可新增複委託，請改至【其他各類所得資料(平日)】輸入！", vbExclamation
                  Cancel = True
               'end 2013/3/21
               ElseIf SetRefData(txtOD(Index)) = False Then
                  Cancel = True
               End If
            End If
            
            'Added by Morgan 2013/1/29
            '翻譯費,複委託預設給付時間為系統時間
            If m_bConfirmCheck = False Then
               If txtOD(Index) = "01" Or txtOD(Index) = "02" Then
                  If txtNHI10 = "" Then
                     txtNHI10 = ServerTime
                  End If
               Else
                  txtNHI10 = ""
               End If
            End If
            'end 2013/1/29
            
            If Cancel = False Then setCombo1 'Added by Morgan 2015/12/14
         
         'Added by Morgan 2015/12/10
         Case 14
            If txtOD(Index) <> "" Then
               If ChkDate(txtOD(Index) & "01") = False Then
                  Cancel = True
               ElseIf Val(txtOD(Index)) > Val(strSrvDate(2)) \ 100 Then
                  MsgBox "補扣/退繳年月不可晚於當月！", vbExclamation
                  Cancel = True
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtOD(Index)
      
      '若是按確定的檢查時略過
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
            Case 5
               If lblDsp(2) = "A" Then
                  'Added by Morgan 2016/6/24
                  '其它薪資 代號"50" 含年終/三節/翻譯等 所得73,001 以上扣稅
                  If txtOD(4) = "01" Then
                     'modify by sonia 2018/4/17 改84,501
                     'If Val(txtOD(Index)) >= 73001 Then
                     If Val(txtOD(Index)) >= 84501 Then
                        txtOD(6) = Val(txtOD(Index)) * Val(lblDsp(4)) \ 100
                     End If
                  Else
                     'end 2016/6/24
                     'Modify by Morgan 2011/6/1 無條件捨去--辜
                     'strExc(1) = Round(Val(txtOD(Index)) * Val(lblDsp(4)) / 100, 4)
                     strExc(1) = Val(txtOD(Index)) * Val(lblDsp(4)) \ 100
                     If Val(strExc(1)) >= 2000 Then
                        txtOD(6) = strExc(1)
                     End If
                  End If 'Added by Morgan 2016/6/24
               End If
         End Select
         
         'Added by Morgan 2013/1/24
         Select Case Index
            Case 2, 3, 4, 5
            If m_arrOD(Index) <> txtOD(Index) Then
               SetNHI06
            End If
         End Select
         'end 2013/1/24
         
      End If
      
      m_arrOD(Index) = txtOD(Index) 'Added by Morgan 2013/1/24
   End If
End Sub

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   '翻譯費要檢查不可有晚於該筆資料的補充保費
   If txtOD(4) = "01" Then
      SetNHI06
      If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10)) = False Then
         GoTo EscPoint
      End If
   End If
   'end 2013/2/6
   
   '刪除
   stSQL = "delete from OtherSalaryData where od01='" & txtOD(1) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   'Added by Morgan 2013/1/24
   '刪除補充保費
   If txtOD(4) = "01" Then
      strSql = "DELETE NHI2ND WHERE NHI01='" & txtOD(3) & "' AND NHI02=" & DBDATE(txtOD(2)) & " AND NHI03='50' AND NHI04='4'"
      cnnConnection.Execute strSql, intI
   ElseIf txtOD(4) = "02" Then
      strSql = "DELETE NHI2ND WHERE NHI01='" & txtOD(3) & "' AND NHI02=" & DBDATE(txtOD(2)) & " AND NHI03='9A' AND NHI04='5'"
      cnnConnection.Execute strSql, intI
   End If
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtOD(1).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
EscPoint:

End Function

Private Function SetRefData(stCode As String) As Boolean
   strExc(0) = "select * from OtherSalaryCode where oc01='" & stCode & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      lblDsp(2) = "" & .Fields("OC02")
      lblOC03 = "" & .Fields("OC03") 'Modify By Sindy 2021/12/22
      lblDsp(4) = "" & .Fields("OC04")
      End With
      SetRefData = True
   Else
      MsgBox "代號輸入錯誤！"
   End If
End Function
'Added by Morgan 2013/1/29
'計算補充保費
Private Sub SetNHI06()
   
   If txtOD(4) = "01" Or txtOD(4) = "02" Then
      If txtOD(2) = "" Or txtOD(3) = "" Or txtOD(4) = "" Or txtOD(5) = "" Then
         txtOD(13) = ""
      Else
         stNHI(1) = txtOD(3)
         stNHI(2) = DBDATE(txtOD(2))
         If txtOD(4) = "01" Then
            stNHI(3) = "50"
            stNHI(4) = "4"
         Else
            stNHI(3) = "9A"
            stNHI(4) = "5"
         End If
         
         stNHI(5) = ""
         stNHI(6) = ""
         stNHI(7) = txtOD(5)
         stNHI(8) = ""
         stNHI(10) = Val(txtNHI10)
         stNHI(11) = GetSalaryCompany(stNHI(1), stNHI(2)) 'Added by Morgan 2013/2/26
         PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13) 'Modified by Morgan 2013/3/12 +NHI13 2014/5/1 +NHI11
         txtOD(13) = Val(stNHI(6))
      End If
   Else
      txtOD(13) = ""
   End If
   SetNet 'Added by Morgan 2013/2/22
End Sub

'Added by Morgan 2013/2/22
Private Sub SetNet()
   txtNet = Val(txtOD(5)) - Val(txtOD(6)) - Val(txtOD(13))
End Sub

Private Sub UpdateSM42()
   'Added by Morgan 2019/9/2
   '更新健保投保薪資/金額(級數)
   If txtSM42.Enabled Then
      If Val(txtSM42.Tag) <> Val(txtSM42) Then
         strSql = "update salarymonth set sm41=" & Val(txtSM42) & ",sm42=" & Val(txtSM42) & " where sm01='" & txtOD(3) & "' and sm02=" & (Val(txtOD(14)) + 191100)
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2019/9/2
End Sub

'Add By Sindy 2021/12/20
Private Sub txtNote_GotFocus()
   TextInverse txtNote
   CloseIme
End Sub
Private Sub txtNote_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'2021/12/20
