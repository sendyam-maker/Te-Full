VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090626 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人/繪圖人員外出記錄維護"
   ClientHeight    =   5508
   ClientLeft      =   1764
   ClientTop       =   1860
   ClientWidth     =   9132
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5508
   ScaleWidth      =   9132
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8385
      Top             =   330
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
            Picture         =   "frm090626.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090626.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   30
      TabIndex        =   26
      Top             =   750
      Width           =   9045
      _ExtentX        =   15939
      _ExtentY        =   8276
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090626.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "LblStarW"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDisp(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lstMailCC"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtOG(12)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtOG(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtOG(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtOG(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtOG(7)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtOG(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtOG(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtOG(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtOG(10)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtOG(11)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "mebOG(20)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "mebOG(19)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090626.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(82)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(81)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(11)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(12)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Line4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "grdList"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdQuery(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtQry(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtQry(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtQry(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtQry(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdQuery(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtST01"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtCode(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtCode(3)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtCode(2)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtCode(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   1
         Left            =   -69960
         MaxLength       =   6
         TabIndex        =   20
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   2
         Left            =   -68940
         MaxLength       =   1
         TabIndex        =   21
         Top             =   840
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   3
         Left            =   -68520
         MaxLength       =   2
         TabIndex        =   22
         Top             =   840
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   0
         Left            =   -70620
         MaxLength       =   3
         TabIndex        =   19
         Top             =   840
         Width           =   585
      End
      Begin VB.TextBox txtST01 
         Height          =   300
         Left            =   -74040
         MaxLength       =   6
         TabIndex        =   18
         Top             =   840
         Width           =   945
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "列印(&P)"
         Height          =   360
         Index           =   1
         Left            =   -67380
         TabIndex        =   24
         Top             =   390
         Width           =   912
      End
      Begin VB.TextBox txtQry 
         Height          =   300
         Index           =   4
         Left            =   -69630
         MaxLength       =   3
         TabIndex        =   17
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox txtQry 
         Height          =   300
         Index           =   3
         Left            =   -70260
         MaxLength       =   3
         TabIndex        =   16
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox txtQry 
         Height          =   300
         Index           =   2
         Left            =   -72990
         MaxLength       =   7
         TabIndex        =   15
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtQry 
         Height          =   300
         Index           =   1
         Left            =   -74040
         MaxLength       =   7
         TabIndex        =   14
         Top             =   450
         Width           =   945
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   360
         Index           =   0
         Left            =   -68340
         TabIndex        =   23
         Top             =   390
         Width           =   912
      End
      Begin MSMask.MaskEdBox mebOG 
         Height          =   270
         Index           =   19
         Left            =   4380
         TabIndex        =   8
         Top             =   1290
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   466
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebOG 
         Height          =   270
         Index           =   20
         Left            =   5670
         TabIndex        =   9
         Top             =   1290
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   466
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3300
         Left            =   -74850
         TabIndex        =   44
         Top             =   1260
         Width           =   8730
         _ExtentX        =   15409
         _ExtentY        =   5821
         _Version        =   393216
         Cols            =   10
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
         _Band(0).Cols   =   10
      End
      Begin VB.Line Line4 
         X1              =   -70050
         X2              =   -68220
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   12
         Left            =   -71430
         TabIndex        =   45
         Top             =   900
         Width           =   765
      End
      Begin MSForms.TextBox txtOG 
         Height          =   720
         Index           =   11
         Left            =   1050
         TabIndex        =   11
         Top             =   2505
         Width           =   5745
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "10134;1270"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   720
         Index           =   10
         Left            =   1050
         TabIndex        =   10
         Top             =   1650
         Width           =   5745
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "10134;1270"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   300
         Index           =   1
         Left            =   4395
         TabIndex        =   1
         Top             =   570
         Width           =   705
         VariousPropertyBits=   671105049
         Size            =   "1244;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   300
         Index           =   2
         Left            =   1050
         TabIndex        =   0
         Top             =   570
         Width           =   945
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   300
         Index           =   6
         Left            =   4395
         TabIndex        =   3
         Top             =   930
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   300
         Index           =   7
         Left            =   5025
         TabIndex        =   4
         Top             =   930
         Width           =   915
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   300
         Index           =   8
         Left            =   6015
         TabIndex        =   5
         Top             =   930
         Width           =   315
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "556;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   300
         Index           =   9
         Left            =   6435
         TabIndex        =   6
         Top             =   930
         Width           =   435
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "767;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   300
         Index           =   4
         Left            =   1050
         TabIndex        =   7
         Top             =   1290
         Width           =   945
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOG 
         Height          =   1140
         Index           =   12
         Left            =   1050
         TabIndex        =   12
         Top             =   3360
         Width           =   5745
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "10134;2011"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstMailCC 
         Height          =   2370
         Left            =   7080
         TabIndex        =   13
         Top             =   1680
         Width           =   1860
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "3281;4180"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblDisp 
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   43
         Top             =   1290
         Width           =   1305
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2302;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Top             =   930
         Width           =   1800
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3175;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Left            =   -72990
         TabIndex        =   42
         Top             =   870
         Width           =   1485
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2619;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外出人員:"
         Height          =   180
         Index           =   11
         Left            =   -74850
         TabIndex        =   41
         Top             =   900
         Width           =   765
      End
      Begin VB.Label LblStarW 
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2040
         TabIndex        =   40
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "副本："
         Height          =   540
         Index           =   10
         Left            =   6820
         TabIndex        =   39
         Top             =   1680
         Visible         =   0   'False
         Width           =   180
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5415
         X2              =   5565
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "事由："
         Height          =   180
         Index           =   7
         Left            =   150
         TabIndex        =   38
         Top             =   2535
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "地點："
         Height          =   180
         Index           =   6
         Left            =   150
         TabIndex        =   37
         Top             =   1695
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "序號："
         Height          =   180
         Index           =   1
         Left            =   3450
         TabIndex        =   36
         Top             =   630
         Width           =   540
      End
      Begin VB.Line Line3 
         X1              =   -69960
         X2              =   -69390
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外出人員部門:"
         Height          =   180
         Index           =   81
         Left            =   -71430
         TabIndex        =   35
         Top             =   450
         Width           =   1125
      End
      Begin VB.Line Line2 
         X1              =   -73260
         X2              =   -72840
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外出日期:"
         Height          =   180
         Index           =   82
         Left            =   -74850
         TabIndex        =   34
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外出日期："
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   33
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   3
         Left            =   3450
         TabIndex        =   32
         Top             =   990
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外出人員："
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   31
         Top             =   990
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   4545
         X2              =   6675
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   30
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外出時數："
         Height          =   180
         Index           =   5
         Left            =   3450
         TabIndex        =   29
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   8
         Left            =   150
         TabIndex        =   28
         Top             =   3405
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(格式 : HH:mm) "
         Height          =   180
         Index           =   9
         Left            =   6810
         TabIndex        =   27
         Top             =   1350
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9132
      _ExtentX        =   16108
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
End
Attribute VB_Name = "frm090626"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)、Label5、lblDisp(1)、txtOG(index); Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'Create by Morgan 2003/12/24
Option Explicit

'前次紀錄KEY
Dim lst_OG01 As String
'本次紀錄KEY
Dim cur_OG01 As String
'目前狀態
Dim iCurState As Integer
'使用者權限設定
Dim bolInsert As Boolean
Dim bolUpdate As Boolean
Dim bolDelete As Boolean
Dim bolSelect As Boolean
'列印控制
Dim PLeft(0 To 8) As Integer
Dim iPrint As Integer
Dim Page As Integer
Dim m_bNoTab As Boolean
'Add by Amy 2016/03/04
Dim strABS001_1 As String, strABS001_2 As String, strABS001_3 As String
Dim bolOAgent As Boolean '是否有職代或審核主管
Dim strAbs As String 'Add by Amy 2016/05/19
Dim m_blnColOrderAsc As Boolean 'Added by Lydia 2016/08/11 欄位資料由小到大排序
Dim strB0124 As String 'Add by Amy 2021/01/11 請假、出差核准後通知人員編號(多筆;區隔)
Dim m_ProState As String 'Added by Morgan 2022/9/13

'檢查查詢條件
Private Function CheckQueryData() As Boolean

   Dim bolCancel As Boolean, i As Integer
   
   If txtQry(1).Text = "" Then
        MsgBox "請輸入外出起日!!!", vbExclamation + vbOKOnly
        txtQry(1).SetFocus
        Exit Function
   End If
   If txtQry(2).Text = "" Then
        MsgBox "請輸入外出迄日!!!", vbExclamation + vbOKOnly
        txtQry(2).SetFocus
        Exit Function
   End If
   
   For i = 1 To 4
      Call txtQry_Validate(i, bolCancel)
      If bolCancel = True Then
         txtQry(i).SetFocus
         Exit Function
      End If
   Next
   CheckQueryData = True
   
End Function

'Modified by Lydia 2022/05/18 + bReset
Private Sub InitGrid(Optional ByVal bReset As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer

   arrGridHeadText = Array("", "序號", "外出日期", "外出人員", "智權人員" _
                     , "外出時間", "本所案號", "地點", "事由", "備註")

   arrGridHeadWidth = Array(300, 700, 850, 850, 800 _
                     , 1100, 1400, 3000, 3000, 3000)
   
   With grdList
      .row = 0
      .Cols = UBound(arrGridHeadText) + 1
      'Added by Lydia 2022/05/18
      If bReset = True Then
          .Clear
          .Rows = 2
          .FixedRows = 0 'Added by Lydia 2022/09/15
      End If
      'end 2022/05/18
      .row = 0 'Added by Lydia 2022/09/15
      For iCol = 0 To .Cols - 1
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignCenterCenter
      Next
      '.Rows = 1  'Remark by Lydia 2022/05/18
      .FixedRows = 1 'Added by Lydia 2022/09/15
      .row = 1 'Added by Morgan 2022/9/16
   End With
End Sub

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)

   Dim iRow As Integer, iCol As Integer
   
   'Modified by Lydia 2022/05/18
'   rsTmp.MoveFirst
'   Do While rsTmp.EOF = False
'      With grdList
'         .Rows = .Rows + 1
'         iRow = .Rows - 1
'         '.TextMatrix(iRow, 0) = iRow
'         For iCol = 1 To grdList.Cols - 1
'            .TextMatrix(iRow, iCol) = "" & rsTmp.Fields(iCol - 1).Value
'         Next iCol
'      End With
'      rsTmp.MoveNext
'   Loop
   Set grdList.Recordset = rsTmp
   Call InitGrid(False)
   'end 2022/05/18
End Sub

Private Function QueryData() As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset
   Dim strCon As String
   
On Error GoTo ErrHand

   strCon = ""
   If txtQry(1) <> "" Then
      strCon = strCon & " AND OG02>=" & 19110000 + Val(txtQry(1))
   End If
   If txtQry(2) <> "" Then
      strCon = strCon & " AND OG02<=" & 19110000 + Val(txtQry(2))
   End If
   If txtQry(3) <> "" Then
      strCon = strCon & " AND A.ST03>='" & txtQry(3) & "'"
   End If
   If txtQry(4) <> "" Then
      strCon = strCon & " AND A.ST03<='" & txtQry(4) & "'"
   End If
   'Add By Sindy 2021/2/5
   If Trim(txtST01) <> "" Then
      strCon = strCon & " AND OG03='" & txtST01 & "'"
   End If
   '2021/2/5 END
   
   'Added by Lydia 2022/05/18 本所案號
   Call InitGrid '預設清空
   If Trim(txtCode(0)) <> "" Then
       strCon = strCon & " AND OG06='" & Trim(txtCode(0)) & "'"
   End If
   If Trim(txtCode(1)) <> "" Then
       strCon = strCon & " AND OG07='" & Trim(txtCode(1)) & "'"
   End If
   If Trim(txtCode(2)) <> "" Then
       strCon = strCon & " AND OG08='" & Trim(txtCode(2)) & "'"
   End If
   If Trim(txtCode(3)) <> "" Then
       strCon = strCon & " AND OG09='" & Trim(txtCode(3)) & "'"
   End If
   'end 2022/05/18
   
   'Modify by Morgan 2010/8/17 百年蟲
   strSql = "select OG01, substrb(' '||sqldatet(OG02),-9) AS OG02, A.ST02 AS OG03, B.ST02 AS OG04, OG19||' - '||OG20 AS OG05, DECODE(OG06,NULL,NULL,OG06||'-'||OG07||'-'||OG08||'-'||OG09) AS OG06, OG10, OG11, OG12" & _
            " from outgoing, staff A, STAFF B" & _
            " where A.ST01(+)=OG03 AND B.ST01(+)=OG04" & strCon & " ORDER BY OG02, OG03"
            
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly

   If rsQuery.RecordCount > 0 Then
      QueryData = True
      Call UpdateGridList(rsQuery)
      'Modified by Lydia 2022/05/18
      'grdList.FixedRows = 1 'Added by Lydia 2022/01/11 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      'Modified by Lydia 2022/09/15 改在UpdateGridList設定
      'If grdList.Rows >= 2 Then
      '    grdList.FixedRows = 1
      'End If
      'end 2022/05/18
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
            
End Function

'報表列印
Private Sub PrintData()

   Dim ii As Integer
   
   Page = 1
   PrintTitle
   With grdList
      For ii = 1 To .Rows - 1
         '日期
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 2)
         '外出人員
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 3)
         '智權人員
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 4)
         '本所案號
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 6)
         '外出時間
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 5)
         '地　點/事　由/備　註
         Printer.CurrentX = PLeft(5)
         Printer.CurrentY = iPrint
         Printer.Print Replace(.TextMatrix(ii, 7), vbCrLf, "") & IIf(.TextMatrix(ii, 8) <> "", "/", "") & Replace(.TextMatrix(ii, 8), vbCrLf, "") & IIf(.TextMatrix(ii, 9) <> "", "/", "") & Replace(.TextMatrix(ii, 9), vbCrLf, "")
         iPrint = iPrint + 300
         If iPrint > 10000 And ii <> .Rows - 1 Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.NewPage
            Page = Page + 1
            PrintTitle
         End If
      Next ii
   End With
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.EndDoc
   
End Sub

Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = PLeft(0) + 1250
   PLeft(2) = PLeft(1) + 1250
   PLeft(3) = PLeft(2) + 1250
   PLeft(4) = PLeft(3) + 1900
   PLeft(5) = PLeft(4) + 1900
End Sub

Sub PrintTitle()
   GetPleft
   
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "外出記錄明細表"

   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "外出日期：" & Format(ChangeTStringToTDateString(Me.txtQry(1).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txtQry(2).Text)

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "外出人員部門別：" & Me.txtQry(3).Text & " " & IIf(Me.txtQry(3).Text <> "" Or Me.txtQry(4).Text <> "", "－", "") & " " & Me.txtQry(4).Text
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "外出日期"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "外出人員"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "外出時間"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "地　點/事　由/備　註"

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   
End Sub

Private Sub cmdQuery_Click(Index As Integer)
   
   If TxtValidate(1) = False Then Exit Sub
   '查詢
   If Index = 0 Then
      grdList.Rows = 1
      m_blnColOrderAsc = True 'Added by Lydia 2016/08/11 欄位資料由小到大排序
      If CheckQueryData = True Then
         Screen.MousePointer = vbHourglass
         grdList.MousePointer = flexHourglass
         If QueryData() = False Then
             MsgBox "無資料", vbOKOnly, "查詢資料"
             txtQry(1).SetFocus
         End If
         grdList.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
      End If
   '列印
   Else
      If grdList.Rows > 1 Then
         Screen.MousePointer = vbHourglass
         PrintData
         ShowPrintOk
         Screen.MousePointer = vbDefault
      Else
         ShowNoData
      End If
   End If
      
End Sub

'Add By Sindy 2014/7/31
'Modified by Lydia 2022/01/03 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Combo1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo1_LostFocus()
   If Combo1 <> "" Then
      Combo1 = Trim(Left(Combo1, 6)) & " " & GetPrjSalesNM(Trim(Left(Combo1, 6)))
   End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
Dim strEmp As String
   
   If Combo1 <> "" Then
      strEmp = GetStaffName(Trim(Left(Combo1, 6)))
      If strEmp = "" Then
         MsgBox "外出人員輸入錯誤！", vbCritical
         Combo1.SetFocus
         Cancel = True
      End If
   End If
End Sub
'2014/7/31 END

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Add By Sindy 2021/2/5 按enter鍵維持換行功能而不是存檔功能
   If KeyCode = vbKeyReturn Then
      Exit Sub
   End If
   '2021/2/5 END
   
   Select Case KeyCode
      Case vbKeyF2
      '新增
         If SSTab1.Tab = 0 And TBar1.Buttons(1).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(1))
         End If
      Case vbKeyF3
      '修改
         If SSTab1.Tab = 0 And TBar1.Buttons(2).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(2))
         End If
      Case vbKeyF5
      '刪除
         If SSTab1.Tab = 0 And TBar1.Buttons(3).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(3))
         End If
      Case vbKeyF4
      '查詢
         If SSTab1.Tab = 0 And TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
         End If
      Case vbKeyHome
      '第一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(6).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(6))
         End If

      Case vbKeyPageUp
      '上一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(7).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(7))
         End If
      Case vbKeyPageDown
      '下一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(8).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(8))
         End If
      Case vbKeyEnd
      '最後筆
         If SSTab1.Tab = 0 And TBar1.Buttons(9).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(9))
         End If
      Case vbKeyF9
      '存檔
         If SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      'Remove by Lydia 2022/01/03 去掉Enter鍵vbKeyReturn
      'Case vbKeyReturn
      ''確定
       '  If SSTab1.Tab = 1 Then
       '     Call cmdQuery_Click(0)
       '  ElseIf SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
       '     Call Tbar1_ButtonClick(TBar1.Buttons(11))
       '  End If
       'end 2022/01/03
      Case vbKeyF10
      '取消
         If SSTab1.Tab = 0 And TBar1.Buttons(12).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
        If TBar1.Buttons(14).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(14))
         End If
    End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Label5.Caption = "" 'Added by Lydia 2022/01/03
   
   Me.Show
   setAuthority
   Call SetCombo1 'Add By Sindy 2014/7/30
   Call SetMailCC 'Add by Amy 2014/12/25
   Call FormReset(0)
   'Modify by Amy 2016/05/19 原:有職代預設勾選搬至SetMailCC
   Call InitGrid
   '預設為瀏覽
   'Modified by Morgan 2019/3/27 改預設在最後一筆--柄佑
   'If doQuery(6) = True Then
   If doQuery(9) = True Then
      iCurState = 0
   Else
      iCurState = 9
   End If
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
   
   SSTab1.Tab = 0 '設定顯示在第一頁 Add By Sindy 2021/2/5
End Sub

'Add By Sindy 2014/7/30
Private Sub SetCombo1()
Dim strTemp As String, arrData As Variant, i As Integer
   Combo1.Clear
   Combo1.AddItem strUserNum & " " & strUserName
   '檢查當時是否需要為他人職代
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)

   '開放部份智權同仁的資料給彥葶操作
   If Pub_GetSpecMan("A8") = strUserNum Then
      strTemp = Pub_GetSpecMan("A7")
      arrData = Split(strTemp, ";")
      For i = 0 To UBound(arrData)
         Combo1.AddItem arrData(i) & " " & GetPrjSalesNM(CStr(arrData(i)))
      Next
   End If
   Combo1.Text = Combo1.List(0)
End Sub

'使用者權限設定
Private Sub setAuthority()
   bolInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   bolUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   bolDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   bolSelect = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   'Added by Morgan 2022/9/13 本功能沒有分個人管理，此變數僅為後面權限檢查用
   m_ProState = "1"
   'Modified by Lydia 2022/09/15 設定不彈訊息ShowMsg =False
   If CheckUse(Me.Name & "M", strExec, False) = True Then
      m_ProState = "2"
   End If
   'end 2022/9/13
End Sub
'檢查本所案號
Private Function CheckCaseNo() As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset
   
On Error GoTo ErrHnd

   CheckCaseNo = False
   
   strSql = "Select PA01 From Patent Where PA01='" & txtOG(6) & "' AND PA02='" & txtOG(7) & "' AND PA03='" & txtOG(8) & "' AND PA04='" & txtOG(9) & "'"
   strSql = strSql & " Union Select TM01 From Trademark Where TM01='" & txtOG(6) & "' AND TM02='" & txtOG(7) & "' AND TM03='" & txtOG(8) & "' AND TM04='" & txtOG(9) & "'"
   strSql = strSql & " Union Select LC01 From Lawcase Where LC01='" & txtOG(6) & "' AND LC02='" & txtOG(7) & "' AND LC03='" & txtOG(8) & "' AND LC04='" & txtOG(9) & "'"
   strSql = strSql & " Union Select HC01 From Hirecase Where HC01='" & txtOG(6) & "' AND HC02='" & txtOG(7) & "' AND HC03='" & txtOG(8) & "' AND HC04='" & txtOG(9) & "'"
   strSql = strSql & " Union Select SP01 From Servicepractice Where SP01='" & txtOG(6) & "' AND SP02='" & txtOG(7) & "' AND SP03='" & txtOG(8) & "' AND SP04='" & txtOG(9) & "'"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      CheckCaseNo = True
   End If
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   Exit Function
   
ErrHnd:

   MsgBox Err.Description
   
End Function

Private Function TxtValidate(Optional ByVal iTab As Integer = 0) As Boolean
   Dim oText As Object, bolCancel As Boolean, arrText, oMaskEdBox As MaskEdBox
   Dim bolChk As Boolean, i As Integer 'Add By Sindy 2014/7/31
   
   TxtValidate = False
   bolCancel = False
   
   Select Case iTab
      Case 0
         SSTab1.Tab = 0
         Set arrText = txtOG
         For Each oMaskEdBox In mebOG
            mebOG_Validate oMaskEdBox.Index, bolCancel
            If bolCancel = True Then
               mebOG_GotFocus oMaskEdBox.Index
               Exit For
            End If
         Next
      Case 1
         Set arrText = txtQry
   End Select
   
   If bolCancel = False Then
      For Each oText In arrText
         If oText.Locked = False Then
            txtOG_Validate oText.Index, bolCancel
            If bolCancel = True Then
               oText.SetFocus
               TextInverse oText
               'Exit For
               Exit Function
            End If
         End If
      Next
   End If
   
   'Add By Sindy 2014/7/31
   If iTab = 0 Then 'Add By Sindy 2014/8/19 +if
      bolCancel = False
      Call Combo1_Validate(bolCancel)
      If bolCancel = True Then Exit Function
      '檢查是否有增修刪權限 P10.專利處主管
      'Modified by Morgan 2022/9/13 改判斷個人權限才檢查(不要限定P10因還會設個人Ex:99050)
      'If Pub_StrUserSt03 <> "P10" And Pub_StrUserSt03 <> "M51" Then
      If m_ProState = "1" Then
      'end 2022/9/13
         bolChk = False
         For i = 0 To Combo1.ListCount - 1
            If Combo1.List(i) = Combo1.Text Then
               bolChk = True
               Exit For
            End If
         Next i
         If bolChk = False Then
            MsgBox "無權限維護該人員資料！", vbExclamation
            'Combo1.SetFocus
            If Combo1.Enabled = True Then Combo1.SetFocus 'Modify By Sindy 2014/8/19
            Exit Function
         End If
      End If
   End If
   '2014/7/31 END
   
    'Added by Lydia 2022/01/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
    
   If bolCancel = False Then TxtValidate = True
End Function

Private Sub mebOG_GotFocus(Index As Integer)
   mebOG(Index).SelStart = 0
   mebOG(Index).SelLength = Len(mebOG(Index))
End Sub

Private Sub mebOG_Validate(Index As Integer, Cancel As Boolean)

   mebOG(Index) = Replace(mebOG(Index), "_", "0")
   If Left(mebOG(Index), 2) >= 24 Then
      MsgBox "時間格式錯誤！"
      Cancel = True
   ElseIf Right(mebOG(Index), 2) >= 60 Then
      MsgBox "時間格式錯誤！"
      Cancel = True
   ElseIf (Index = 20 And mebOG(19) <> "__:__" And mebOG(19) > mebOG(20)) Then
      MsgBox "時間起迄錯誤！"
      Cancel = True
   End If
   
   If Cancel = True Then
      mebOG_GotFocus (Index)
      mebOG(Index).SetFocus
   End If
   
End Sub

Private Function CheckConfirm() As Boolean
   
   CheckConfirm = False
   
   Select Case iCurState
      '1:新增;2:修改
      Case 1, 2
      
         If TxtValidate = False Then Exit Function
         
         '外出日期
         If txtOG(2) = "" Then
            MsgBox "外出日期不可空白！", vbCritical
            txtOG(2).SetFocus
            Call txtOG_GotFocus(2)
            Exit Function
         '外出人員
         'Modify By Sindy 2014/7/31
         'ElseIf txtOG(3) = "" Then
         ElseIf Trim(Combo1.Text) = "" Then
         '2014/7/31 END
            MsgBox "外出人員不可空白！", vbCritical
            Combo1.SetFocus
            Exit Function
         '沒有打本所案號第一碼
         ElseIf txtOG(6) = "" And (txtOG(7) <> "" Or txtOG(8) <> "" Or txtOG(9) <> "") Then
               MsgBox "本所案號錯誤！", vbCritical
               txtOG(6).SetFocus
               Call txtOG_GotFocus(6)
               Exit Function
         '有打本所案號第一碼
         ElseIf txtOG(6) <> "" Then
               If CheckCaseNo() = False Then
                  MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                  txtOG(6).SetFocus
                  Call txtOG_GotFocus(6)
                  Exit Function
               End If
         End If
         
         'Add By Sindy 2015/1/16 +假單檢查
         If CheckIsAbsenceExist(Trim(Left("" & Combo1.Text, 6)), txtOG(2), mebOG(19), txtOG(2), mebOG(20), txtOG(1)) = True Then
            MsgBox "已有此外出記錄，請檢查是否已填外出記錄或假單!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
      '查詢
      Case 4
         If txtOG(1) = "" Then
            MsgBox "序號不可空白！", vbCritical
            txtOG(1).SetFocus
            Call txtOG_GotFocus(1)
            Exit Function
         End If
   End Select
   CheckConfirm = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090626 = Nothing
End Sub

Private Sub grdList_DblClick()

   Dim lRow As Long, lCurRow As Long, iCol As Integer
   
   lCurRow = grdList.row
   '呼叫查詢
   If lCurRow > 0 Then
      'If TBar1.Buttons(4).Enabled = True Then 'Removed by Morgan 2024/11/4 取消此判斷(支援紀錄也沒有),實際都沒有設查詢權限但都可瀏覽及上下筆
         Call Tbar1_ButtonClick(TBar1.Buttons(4))
         If txtOG(1).Locked = False Then
            txtOG(1).Text = grdList.TextMatrix(lCurRow, 1)
            If TBar1.Buttons(11).Enabled = True Then
               Call Tbar1_ButtonClick(TBar1.Buttons(11))
            End If
         End If
      'End If
   End If
End Sub

Private Sub grdList_Click()
      
   Dim lRow As Long, lCurRow As Long, iCol As Integer
   
   With grdList
      lCurRow = .row
      If lCurRow > 0 Then
         '還原
         For lRow = 1 To .Rows - 1
            .row = lRow: iCol = 1
            If .CellBackColor <> &H80000005 Then
               For iCol = 1 To .Cols - 1
                   .col = iCol
                   .CellBackColor = &H80000005
                   .CellForeColor = &H80000008
               Next iCol
            End If
         Next lRow
         '反白
         .row = lCurRow
         For iCol = 1 To .Cols - 1
             .col = iCol
             .CellBackColor = &H8000000D
             .CellForeColor = &H80000005
         Next iCol
      End If
      
      m_bNoTab = True
      grdList_DblClick
      m_bNoTab = False
   End With
End Sub

'Modify by Amy 2017/12/06 改Pub_SendMail發
Private Sub SendMail()
    Dim stCaseNo As String, strSubject As String, strContent As String
    Dim strTo As String, stToCC As String
    Dim ii As Integer
    
    'Removed by Morgan 2024/12/20 71011,67002皆已退休,程式已無效
    'If strUserNum = "71011" Or strUserNum = "67002" Then
    '    strTo = "94007"
    '    '71011登入且有選副本也寄
    '    If strUserNum = "71011" Then
    '        For ii = 0 To lstMailCC.ListCount - 1
    '            If lstMailCC.Selected(ii) = True Then
    '                '副本排除林總(因收件者已有)
    '                If Right(lstMailCC.List(ii), 5) <> "94007" Then
    '                    stToCC = stToCC & ";" & Right(lstMailCC.List(ii), 5)
    '                End If
    '            End If
    '        Next ii
    '        '副本欄位
    '        If stToCC <> "" Then stToCC = Mid(stToCC, 2)
    '    End If
    'Else
    'end 2024/12/20
    
        Select Case Left(GetStaffDepartment(strUserNum), 2)
            Case "P1"
                'Modified by Morgan 2024/12/20 71011已退休,刪除無效判斷及控制
                ''Added by Lydia 2023/04/24 修改王副總退休之相關控制
                'If strSrvDate(1) >= "20230511" Then
                '    strTo = "73022;99050"
                'ElseIf strSrvDate(1) >= "20230501" Then
                '    strTo = "71011;73022;99050"
                'Else
                ''end 2023/04/24
                '    strTo = "71011"
                'End If 'Added by Lydia 2023/04/24
                'Modified by Morgan 2025/2/21
                pub_PMan = Pub_GetSpecMan("專利處特定編號")
                strTo = pub_PMan & ";99050"
                'end 2025/2/21
                'end 2024/12/20
                
                '副本
                For ii = 0 To lstMailCC.ListCount - 1
                    If lstMailCC.Selected(ii) = True Then
                        '副本排除王副總(因收件者不可為空)
                        'Modified by Morgan 2024/12/20 71011已退休,刪除無效判斷及控制
                        'If Right(lstMailCC.List(ii), 5) <> "71011" Then
                        '    stToCC = stToCC & ";" & Right(lstMailCC.List(ii), 5)
                        'End If
                        stToCC = stToCC & ";" & Right(lstMailCC.List(ii), 5)
                        'end 2024/12/20
                    End If
                Next ii
                '副本欄位
                If stToCC <> "" Then stToCC = Mid(stToCC, 2)
            Case "P2"
                'strTo = "67002"  'cancel by sonia 2020/5/5
            Case Else
                Exit Sub
        End Select
        
    'End If 'Removed by Morgan 2024/12/20
   
    stCaseNo = ""
    If txtOG(6) <> "" Then
        stCaseNo = txtOG(6).Text & "-" & txtOG(7).Text & "-" & txtOG(8).Text & "-" & txtOG(9).Text
    End If
   
    If iCurState = 2 Then '修改
        strSubject = "<<外出記錄>>修改記錄通知"
    Else
        strSubject = "<<外出記錄>>新增記錄通知"
    End If
    strContent = "外出日期：" & ChangeTStringToTDateString(txtOG(2).Text) & vbCrLf & _
                 "外出人員：" & Trim(Combo1.Text) & vbCrLf & _
                 "智權人員：" & txtOG(4).Text & " " & lblDisp(1).Caption & vbCrLf & _
                 "外出時間：" & mebOG(19).Text & " - " & mebOG(20).Text & vbCrLf & _
                 "本所案號：" & stCaseNo & vbCrLf & _
                 "地　　點：" & Replace(txtOG(10).Text, vbCrLf, vbCrLf & String(4, "　")) & vbCrLf & _
                 "事　　由：" & Replace(txtOG(11).Text, vbCrLf, vbCrLf & String(4, "　")) & vbCrLf & _
                 "備　　註：" & Replace(txtOG(12).Text, vbCrLf, vbCrLf & String(4, "　")) & vbCrLf
   
   If iCurState = 2 Then '修改
        strContent = strContent & _
                     "修改資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
   Else
        strContent = strContent & _
                     "新增資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
   End If
   'Added by Lydia 2022/05/30 傳入收文號
   strExc(0) = ""
   If txtOG(6) <> "" And txtOG(7) <> "" Then
       strExc(0) = PUB_GetLastABKindCP09(txtOG(6), txtOG(7), IIf(txtOG(8) <> "", txtOG(8), "0"), IIf(txtOG(9) <> "", txtOG(9), "00"))
   End If
   If strExc(0) <> "" Then
       PUB_SendMail strUserNum, strTo, strExc(0), strSubject, strContent, , , False, , , stToCC
   Else
   'end 2022/05/30
       PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , False, , , stToCC
   End If 'Added by Lydia 2022/05/30
End Sub

Private Sub SendMail_Old()
'Dim stCaseNo As String
'Dim stToCC As String, ii As Integer 'Add by Amy 2014/12/25
'
'   If strUserNum = "71011" Or strUserNum = "67002" Then
'      'modify by sonia 2014/9/9 改68001為94007
'      frm880005.txtEmail(0).Text = "94007"
'      'Modify by Amy 2014/12/25 71011登入且有選副本也寄
'      If strUserNum = "71011" Then
'         For ii = 0 To lstMailCC.ListCount - 1
'            If lstMailCC.Selected(ii) = True Then
'                'Modify by Amy 2016/07/29 副本林總(因收件者已有)
'                If Right(lstMailCC.List(ii), 5) <> "94007" Then
'                    stToCC = stToCC & ";" & Right(lstMailCC.List(ii), 5)
'                End If
'            End If
'         Next ii
'         '副本欄位
'         If stToCC <> "" Then frm880005.txtEmail(5).Text = Mid(stToCC, 2)
'      End If
'      'end 2014/12/25
'   Else
'      Select Case Left(GetStaffDepartment(strUserNum), 2)
'         Case "P1"
'            frm880005.txtEmail(0).Text = "71011"
'            'Add by Amy 2016/03/28 +副本也寄
'            stToCC = ""
'            For ii = 0 To lstMailCC.ListCount - 1
'               If lstMailCC.Selected(ii) = True Then
'                  'Modify by Amy 2016/07/25 副本排除王副總(因收件者不可為空)
'                  If Right(lstMailCC.List(ii), 5) <> "71011" Then
'                    stToCC = stToCC & ";" & Right(lstMailCC.List(ii), 5)
'                  End If
'               End If
'            Next ii
'            '副本欄位
'            If stToCC <> "" Then frm880005.txtEmail(5).Text = Mid(stToCC, 2)
'            'end 2016/07/25
'         Case "P2"
'            frm880005.txtEmail(0).Text = "67002"
'         Case Else
'            Exit Sub
'         End Select
'   End If
'   '若使用者為北所人員, 則E-Mail後面不加@taie.com.tw
'   If PUB_GetST06(strUserNum) = "1" Then
'       '無動作
'   '若使用者非北所人員, 則E-Mail後面加@taie.com.tw
'   Else
'      frm880005.txtEmail(0).Text = frm880005.txtEmail(0).Text & "@taie.com.tw"
'   End If
'
'   stCaseNo = ""
'   If txtOG(6) <> "" Then
'      stCaseNo = txtOG(6).Text & "-" & txtOG(7).Text & "-" & txtOG(8).Text & "-" & txtOG(9).Text
'   End If
'   'Add By Sindy 2015/11/19
'   If iCurState = 2 Then '修改
'      frm880005.txtEmail(1).Text = "<<外出記錄>>修改記錄通知"
'   Else
'   '2015/11/19 END
'      frm880005.txtEmail(1).Text = "<<外出記錄>>新增記錄通知"
'   End If
'   frm880005.txtEmail(2).Text = "外出日期：" & ChangeTStringToTDateString(txtOG(2).Text) & vbCrLf & _
'                                 "外出人員：" & Trim(Combo1.Text) & vbCrLf & _
'                                 "智權人員：" & txtOG(4).Text & " " & lblDisp(1).Caption & vbCrLf & _
'                                 "外出時間：" & mebOG(19).Text & " - " & mebOG(20).Text & vbCrLf & _
'                                 "本所案號：" & stCaseNo & vbCrLf & _
'                                 "地　　點：" & Replace(txtOG(10).Text, vbCrLf, vbCrLf & String(4, "　")) & vbCrLf & _
'                                 "事　　由：" & Replace(txtOG(11).Text, vbCrLf, vbCrLf & String(4, "　")) & vbCrLf & _
'                                 "備　　註：" & Replace(txtOG(12).Text, vbCrLf, vbCrLf & String(4, "　")) & vbCrLf
'   'Add By Sindy 2015/11/19
'   If iCurState = 2 Then '修改
'      frm880005.txtEmail(2).Text = frm880005.txtEmail(2).Text & _
'                                   "修改資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
'   Else
'   '2015/11/19 END
'      frm880005.txtEmail(2).Text = frm880005.txtEmail(2).Text & _
'                                   "新增資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
'   End If
'   frm880005.Form_Activate: DoEvents
'   frm880005.cmdok_Click 0: DoEvents
End Sub
'end 2017/12/06

Private Sub SSTab1_Click(PreviousTab As Integer)

On Error Resume Next 'Added by Lydia 2022/01/03 因為改Form 2.0後，第一次lstMailCC.Visible = True會出錯
   Select Case PreviousTab
      Case 0
         If iCurState = 0 Then txtQry(1).SetFocus
      'Add by Amy 2016/03/04 專利處登入顯示副本
         If lstMailCC.Visible = True Then lstMailCC.Visible = False
      Case 1
         If Left(Pub_StrUserSt03, 2) = "P1" Or Pub_StrUserSt03 = "M51" Then
            lstMailCC.Visible = True
         End If
      'end 2016/03/04
   End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim bolChk As Boolean, i As Integer 'Add By Sindy 2014/7/31
   
   SSTab1.Tab = 0 'Add by Morgan 2011/10/19
   
   'Add By Sindy 2014/7/31
   If Button.Index = 2 Or Button.Index = 3 Then
      '檢查是否有增修刪權限 P10.專利處主管
      'Modified by Morgan 2022/9/13 改判斷個人權限才檢查(不要限定P10因還會設個人Ex:99050)
      'If Pub_StrUserSt03 <> "P10" And Pub_StrUserSt03 <> "M51" Then
      If m_ProState = "1" Then
      'end 2022/9/13
         bolChk = False
         For i = 0 To Combo1.ListCount - 1
            If Combo1.List(i) = Combo1.Text Then
               bolChk = True
               Exit For
            End If
         Next i
         If bolChk = False Then
            MsgBox "無權限維護該人員資料！", vbExclamation
            Exit Sub
         End If
      End If
   End If
   '2014/7/31 END
   
   Select Case Button.Index
      Case 1
      '新增
         iCurState = 1
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
      Case 2
      '修改
         iCurState = 2
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
      Case 3
         SSTab1.Tab = 0 'Add by Morgan 2011/10/19
      '刪除
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            If DeleteData = True Then
               If doQuery(8, False) = True Then
                  iCurState = 0
               ElseIf doQuery(9) = True Then
                  iCurState = 0
               Else
                  cur_OG01 = ""
                  iCurState = 9
               End If
            End If
         End If
      Case 4
      '查詢
         iCurState = 4
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
      Case 6
      '第一筆
         Call doQuery(6)
      Case 7
      '上一筆
         Call doQuery(7)
      Case 8
      '下一筆
         Call doQuery(8)
      Case 9
      '最後筆
         Call doQuery(9)
      Case 11
      '確定
         If CheckConfirm = False Then Exit Sub
         Select Case iCurState
            '新增
            Case 1
               If insertdata() = False Then
                  Exit Sub
               Else
                  '寄E-Mail
                  SendMail
               End If
            '查詢
            Case 4
               cur_OG01 = txtOG(1)
               
            '修改
            Case 2
               If UpdateData() = False Then
                  Exit Sub
               'Add By Sindy 2015/11/19
               Else
                  If txtOG(2).Tag <> txtOG(2).Text Or _
                     Combo1.Tag <> Combo1.Text Or _
                     txtOG(4).Tag <> txtOG(4).Text Or _
                     mebOG(19).Tag <> mebOG(19).Text Or _
                     mebOG(20).Tag <> mebOG(20).Text Or _
                     txtOG(10).Tag <> txtOG(10).Text Or _
                     txtOG(11).Tag <> txtOG(11).Text Or _
                     txtOG(12).Tag <> txtOG(12).Text Then
                     '寄E-Mail
                     SendMail
                  End If
               '2015/11/19 END
               End If
         End Select
         '重新查詢
         If doQuery(4) = True Then
            Call SetToolBar(0)
            Call SetInputs
         Else
            If iCurState = 4 Then
               txtOG(1).SetFocus
               Call txtOG_GotFocus(1)
            End If
            Exit Sub
         End If
         iCurState = 0
      Case 12
      '取消
         Select Case iCurState
            
            '1:新增
            Case 1
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               ElseIf cur_OG01 = "" Then
                  If doQuery(6) = True Then
                     iCurState = 0
                  Else
                     iCurState = 9
                  End If
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
            '2:修改
            Case 2
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
            '查詢
            Case 4
               cur_OG01 = lst_OG01
               If cur_OG01 = "" Then
                  If doQuery(6) = True Then
                     iCurState = 0
                  Else
                     iCurState = 9
                  End If
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
         End Select
      Case 14
      '結束
         If iCurState = 2 Or iCurState = 1 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               Unload Me
               Exit Sub
            End If
         Else
            Unload Me
            Exit Sub
         End If
         
   End Select
   
   If iCurState = 0 Or iCurState = 9 Then SSTab1.TabEnabled(1) = True 'Add by Morgan 2011/10/19
   
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
   lst_OG01 = cur_OG01
End Sub

'清除畫面
Private Sub FormReset(Optional ByVal iTab As Integer = 0)

   Dim oText As Object, oLabel As Object, oMaskEdBox As MaskEdBox
   
   Select Case iTab
   
      Case 0
      '頁籤0
         For Each oText In txtOG
            oText.Text = ""
         Next
         For Each oMaskEdBox In mebOG
            oMaskEdBox.Text = "00:00"
         Next
         For Each oLabel In lblDisp
            oLabel.Caption = ""
         Next
         Combo1.Text = "" 'Add By Sindy 2014/7/31
      Case 1
      '頁籤1
      
   End Select
End Sub
'工具列控制
Private Sub SetToolBar(Optional ByVal iStatus As Integer)

   Dim i As Integer
   For i = 1 To 13
      TBar1.Buttons(i).Enabled = False
   Next
   TBar1.Buttons(14).Enabled = True
   
   Select Case iStatus
   
      Case 0
      '瀏覽
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         If bolUpdate Then
            TBar1.Buttons(2).Enabled = True
         End If
         If bolDelete Then
            TBar1.Buttons(3).Enabled = True
         End If
         If bolSelect Then
            TBar1.Buttons(4).Enabled = True
         End If
         TBar1.Buttons(6).Enabled = True
         TBar1.Buttons(7).Enabled = True
         TBar1.Buttons(8).Enabled = True
         TBar1.Buttons(9).Enabled = True
         
      Case 1, 2, 4
      '1:新增  '2:修改  '4查詢
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
               
      Case 9
      '無資料
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         
   End Select
   
End Sub
'設定文字框
Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)

   Dim oText As Object, oLabel As Object, oMaskEdBox As MaskEdBox
   
   Select Case iStatus
      
      Case 0
      '瀏覽
         For Each oText In txtOG
            oText.Enabled = True
            oText.Locked = True
         Next
         For Each oMaskEdBox In mebOG
            oMaskEdBox.Enabled = False
         Next
         txtOG(2).SetFocus
         Combo1.Enabled = False 'Add By Sindy 2014/7/31
         If lstMailCC.Visible = True Then lstMailCC.Enabled = False 'Add by Amy 2014/12/25
      Case 1
      '新增
         SSTab1.Tab = 0
         For Each oText In txtOG
            oText.Text = ""
            oText.Locked = False
            oText.Enabled = True
         Next
         For Each oMaskEdBox In mebOG
            oMaskEdBox.Enabled = True
         Next
         Call FormReset(0)
         'txtOG(2).Text = strSrvDate(2) Modify By Sindy 2021/2/5 柏翰:
         Combo1.Enabled = True 'Add By Sindy 2014/7/31
         Combo1.ListIndex = 0
         If lstMailCC.Visible = True Then lstMailCC.Enabled = True 'Add by Amy 2014/12/25
'         txtOG(3).Text = strUserNum
'         lblDisp(0).Caption = strUserName
         '序號
         txtOG(1).Enabled = False
         txtOG(2).SetFocus
         Call txtOG_GotFocus(2)
      Case 2
      '修改
         SSTab1.Tab = 0
         For Each oText In txtOG
            oText.Locked = False
            oText.Enabled = True
         Next
         For Each oMaskEdBox In mebOG
            oMaskEdBox.Enabled = True
         Next
         '序號
         txtOG(1).Enabled = False
         txtOG(2).SetFocus
         Call txtOG_GotFocus(2)
         Combo1.Enabled = True 'Add By Sindy 2014/7/31
         If lstMailCC.Visible = True Then lstMailCC.Enabled = True 'Add by Amy 2014/12/25
      Case 4
      '查詢
         If m_bNoTab = False Then
            SSTab1.Tab = 0
         End If
         For Each oText In txtOG
            oText.Locked = False
            oText.Enabled = False
         Next
         For Each oMaskEdBox In mebOG
            oMaskEdBox.Enabled = False
         Next
         Call FormReset(0)
         txtOG(1).Enabled = True
         txtOG(1).SetFocus
         Combo1.Enabled = False 'Add By Sindy 2014/7/31
         If lstMailCC.Visible = True Then lstMailCC.Enabled = False  'Add by Amy 2014/12/25
      Case 9
      '無資料
         For Each oText In txtOG
            oText.Enabled = False
            oText.Locked = True
         Next
         For Each oMaskEdBox In mebOG
            oMaskEdBox.Enabled = False
         Next
         Call FormReset(0)
   End Select
   
End Sub
'讀取資料
Private Function doQuery(ByVal iAct As Integer, Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery = False
   
   Select Case iAct
      Case 4
      '查詢
         strSql = "Select OG01 From OutGoing where OG01='" & cur_OG01 & "'"
         stMessage = "查無資料！"
   
      Case 6
      '第一筆
         strSql = "Select OG01 From OutGoing ORDER BY 1 ASC"
         stMessage = "無外出紀錄！"
      Case 7
      '上一筆
         strSql = "Select OG01 From OutGoing where OG01<'" & cur_OG01 & "'" & _
            " ORDER BY 1 DESC"
         stMessage = "已是第一筆了！"

      Case 8
      '下一筆
         strSql = "Select OG01 From OutGoing where OG01>'" & cur_OG01 & "'" & _
            " ORDER BY 1 ASC"
         stMessage = "已是最後一筆了！"

      Case 9
      '最後筆
         strSql = "Select OG01 From OutGoing" & _
            " ORDER BY 1 DESC"
         stMessage = "無外出紀錄！"
        
   End Select
   
On Error GoTo ErrHand

   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
         lst_OG01 = cur_OG01
         cur_OG01 = "" & rsQuery.Fields(0).Value
         If ReQuery() = True Then doQuery = True
   ElseIf bolMsg Then
      MsgBox stMessage, vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtOG_GotFocus(Index As Integer)
   If txtOG(Index).Locked = False Then
      TextInverse txtOG(Index)
      Select Case Index
         Case 10, 11, 12
            'edit by nickc 2007/07/11 切換輸入法改用API
            'txtOG(Index).IMEMode = 1
            OpenIme
         Case Else
            'edit by nickc 2007/07/11 切換輸入法改用API
            'txtOG(Index).IMEMode = 2
            CloseIme
      End Select
   End If
End Sub

'完整資料查詢
Private Function ReQuery(Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, intI As Integer
   
On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   
   ReQuery = False
   'Add By Sindy 2015/11/19
   txtOG(2).Tag = ""
   Combo1.Tag = ""
   txtOG(4).Tag = ""
   mebOG(19).Tag = ""
   mebOG(20).Tag = ""
   txtOG(10).Tag = ""
   txtOG(11).Tag = ""
   txtOG(12).Tag = ""
   '2015/11/19 END
   
   strSql = "SELECT OG01,OG02-19110000 AS OG02,OG03,OG04,OG06,OG07,OG08,OG09,OG10,OG11,OG12, OG19,OG20" & _
            ",A.ST02 AS D01, B.ST02 AS D02" & _
         " From OutGoing, STAFF A,STAFF B Where A.ST01(+)=OG03 AND B.ST01(+)=OG04 AND OG01='" & cur_OG01 & "'"

   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      txtOG(1) = cur_OG01
      For intI = 1 To 4
         'Add By Sindy 2014/7/31
         If intI = 3 Then
            Combo1 = rsQuery.Fields("OG" & Format(intI, "00")) & " " & GetPrjSalesNM(rsQuery.Fields("OG" & Format(intI, "00")))
         Else
         '2014/7/31 END
            txtOG(intI) = "" & rsQuery.Fields("OG" & Format(intI, "00"))
         End If
      Next intI
      
      'Add By Sindy 2020/9/7 顯示星期幾
      If Val(txtOG(2)) > 0 Then
         LblStarW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtOG(2)), "####/##/##"))) & ")"
      End If
      
      For intI = 6 To 12
         txtOG(intI) = "" & rsQuery.Fields("OG" & Format(intI, "00"))
      Next intI
      
      For intI = 19 To 20
         mebOG(intI) = "" & rsQuery.Fields("OG" & Format(intI, "00"))
      Next intI
      'Modify By Sindy 2014/7/31
      For intI = 2 To 2
      '2014/7/31 END
         lblDisp(intI - 1) = "" & rsQuery.Fields("D" & Format(intI, "00"))
      Next intI
      ReQuery = True
      
      'Add By Sindy 2015/11/19
      txtOG(2).Tag = txtOG(2).Text
      Combo1.Tag = Combo1.Text
      txtOG(4).Tag = txtOG(4).Text
      mebOG(19).Tag = mebOG(19).Text
      mebOG(20).Tag = mebOG(20).Text
      txtOG(10).Tag = txtOG(10).Text
      txtOG(11).Tag = txtOG(11).Text
      txtOG(12).Tag = txtOG(12).Text
      '2015/11/19 END
   ElseIf bolMsg Then
      MsgBox "外出紀錄序號〔" & cur_OG01 & "〕已被刪除！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Screen.MousePointer = vbDefault
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
   Screen.MousePointer = vbDefault
End Function

'Modified by Lydia 2022/01/03 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtOG_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If txtOG(Index).Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
      
         Case 1, 2 ', 3, 4
         '序號,外出日期,外出人員,智權人員:只可為數字
            If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
         Case 5
         '時數:只可為數字與.
            If Not (KeyAscii = 8 Or KeyAscii = 46 Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
         Case 6
         '本所案號:只可為字母
            If Not (KeyAscii = 8 Or (KeyAscii > 64 And KeyAscii < 91)) Then
               KeyAscii = 0
            End If
         Case 7, 8, 9
         '本所案號:只可為數字
            If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
         Case 10, 11, 12
            
      End Select
   End If
End Sub

Private Sub txtOG_LostFocus(Index As Integer)
   If SSTab1.Tab = 1 Then
      Dim bolCancel As Boolean
      bolCancel = False
      Call txtOG_Validate(Index, bolCancel)
      If bolCancel = True Then
         SSTab1.Tab = 0
         txtOG(Index).SetFocus
      End If
   End If
End Sub

Private Sub txtOG_Validate(Index As Integer, Cancel As Boolean)
   If Index = 3 Then Exit Sub 'Add By Sindy 2014/7/31
   If txtOG(Index).Locked = False Then
      Select Case Index
         Case 1
            If txtOG(Index) <> "" Then
               txtOG(Index) = UCase(Right("000000000" & txtOG(Index).Text, 6))
            End If
         Case 2
         '外出日期
            LblStarW.Caption = "" 'Add By Sindy 2020/9/7
            If PUB_CheckKeyInDate(txtOG(Index)) <> 0 Then
               Cancel = True
            Else
               'Add By Sindy 2020/9/7 顯示星期幾
               If Val(txtOG(Index)) > 0 Then
                  LblStarW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtOG(Index)), "####/##/##"))) & ")"
               End If
            End If
'         Case 3
'         '外出人員
'            If txtOG(Index) <> "" Then
'               lblDisp(0) = GetStaffName(txtOG(Index))
'               If lblDisp(0) = "" Then
'                  MsgBox "外出人員輸入錯誤！", vbCritical
'                  Cancel = True
'               End If
'            End If
         Case 4
         '智權人員
            'Modify By Sindy 2015/3/19
'            If txtOG(Index) <> "" Then
'               lblDisp(1) = GetStaffName(txtOG(Index))
'               If lblDisp(1) = "" Then
'                  MsgBox "智權人員輸入錯誤！", vbCritical
'                  Cancel = True
'               End If
'            End If
            lblDisp(1).Caption = ""
            If txtOG(Index) <> "" Then
               If ByInputGetST01or02(txtOG(Index).Text, strExc(0), strExc(1)) = False Then
                  Cancel = True
                  Me.txtOG(Index).SetFocus
               End If
               Me.txtOG(Index).Text = strExc(0)
               lblDisp(1).Caption = strExc(1)
            End If
            '2015/3/19 END
         Case 5
         '時數
            If txtOG(Index) <> "" Then
               txtOG(Index) = Format(Round(txtOG(Index), 1))
               If Not (Val(txtOG(Index)) > 0 And Val(txtOG(Index)) < 1000) Then
                  MsgBox "請輸入大於 0 小於 1000 的數字！", vbCritical
                  Cancel = True
               End If
            End If
         Case 6
         '本所案號
            txtOG(Index) = Trim(txtOG(Index))
            If CheckSysKind(txtOG(6)) = False Then
               MsgBox "系統代碼輸入錯誤！", vbCritical
               Cancel = True
            End If
         Case 7
         '本所案號
            If txtOG(6) <> "" Then
               txtOG(Index) = UCase(Right("000000" & txtOG(Index).Text, 6))
            End If
         Case 8
         '本所案號
            If txtOG(6) <> "" Then
               txtOG(Index) = UCase(Right("0" & txtOG(Index).Text, 1))
            End If
         Case 9
         '本所案號
            If txtOG(6) <> "" Then
               txtOG(Index) = UCase(Right("00" & txtOG(Index).Text, 2))
            End If
         Case 10, 11, 12
         '10:地點,11:事由,12:備註
            If CheckLengthIsOK(txtOG(Index), 200) = False Then
               Cancel = True
            End If
      End Select
   End If
   If Cancel = True Then Call txtOG_GotFocus(Index)
End Sub

Private Function CheckSysKind(ByVal stSys As String, Optional ByVal bolMsg As Boolean) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   Dim i As Integer, arrSys() As String
   
On Error GoTo ErrHand

   CheckSysKind = False
   strSql = "Select SK01 from systemkind"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly

   If rsQuery.RecordCount > 0 Then
      arrSys = Split(rsQuery.GetString, Chr(13))
      For i = 0 To UBound(arrSys)
         If arrSys(i) = stSys Then
            CheckSysKind = True
            Exit For
         End If
      Next i
   Else
      If bolMsg Then
         MsgBox "無法取得系統代碼！", vbCritical
      End If
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
   
End Function

Private Function DeleteData() As Boolean
   Dim strSql As String, lngEffRec As Long
   
   strSql = "Delete OutGoing Where OG01='" & cur_OG01 & "'"
   
   DeleteData = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   DeleteData = True
   
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function UpdateData() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, OG(2 To 17) As String
   Dim rsQuery As New ADODB.Recordset, strUpdSQL As String, lngEffRec As Long
'   Dim strB1009 As String, strB1010 As String 'Add by Sindy 2013/6/27
   
   OG(2) = "OG02=" & Val(txtOG(2).Text) + 19110000
   OG(3) = "OG03='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   OG(4) = "OG04='" & txtOG(4).Text & "'"
'   'Modify by Sindy 2013/6/27
'   Call PUB_CountDayHour(txtOG(3), DBDATE(txtOG(2)), Replace(mebOG(19), ":", ""), DBDATE(txtOG(2)), Replace(mebOG(20), ":", ""), "", "", strB1009, strB1010, "", True)
'   OG(5) = "OG05=" & strB1010
'   '2013/6/27 END
   OG(5) = "OG05=1"
   OG(6) = "OG06='" & txtOG(6).Text & "'"
   OG(7) = "OG07='" & txtOG(7).Text & "'"
   OG(8) = "OG08='" & txtOG(8).Text & "'"
   OG(9) = "OG09='" & txtOG(9).Text & "'"
   OG(10) = "OG10='" & ChgSQL(txtOG(10).Text) & "'"
   OG(11) = "OG11='" & ChgSQL(txtOG(11).Text) & "'"
   OG(12) = "OG12='" & ChgSQL(txtOG(12).Text) & "'"
   OG(13) = "OG16='" & strUserNum & "'"
   OG(14) = "OG17=TO_NUMBER(TO_CHAR(SYSDATE,'YYYYMMDD'))"
   OG(15) = "OG18=TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   OG(16) = "OG19='" & mebOG(19).Text & "'"
   OG(17) = "OG20='" & mebOG(20).Text & "'"
   
   strUpdSQL = Join(OG, ",")
   
   strSql = "Update OutGoing Set " & strUpdSQL & " Where OG01='" & cur_OG01 & "'"
         
   UpdateData = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   UpdateData = True
   
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function insertdata() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, OG(1 To 17) As String
   Dim strCols As String, strValues As String, lngEffRec As Long
   Dim rsQuery As New ADODB.Recordset
   
   strCols = "OG01"
   For intI = 2 To 15
      strCols = strCols & ",OG" & Format(intI, "00")
   Next intI
   For intI = 19 To 20
      strCols = strCols & ",OG" & Format(intI, "00")
   Next intI
   
   OG(1) = "LPAD(TO_CHAR(x+1),6,'0')"
   OG(2) = Val(txtOG(2).Text) + 19110000
   OG(3) = "'" & Trim(Left("" & Combo1.Text, 6)) & "'"
   OG(4) = "'" & txtOG(4).Text & "'"
   OG(5) = "1"
   OG(6) = "'" & txtOG(6).Text & "'"
   OG(7) = "'" & txtOG(7).Text & "'"
   OG(8) = "'" & txtOG(8).Text & "'"
   OG(9) = "'" & txtOG(9).Text & "'"
   OG(10) = "'" & ChgSQL(txtOG(10).Text) & "'"
   OG(11) = "'" & ChgSQL(txtOG(11).Text) & "'"
   OG(12) = "'" & ChgSQL(txtOG(12).Text) & "'"
   OG(13) = "'" & strUserNum & "'"
   OG(14) = "TO_NUMBER(TO_CHAR(SYSDATE,'YYYYMMDD'))"
   OG(15) = "TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   OG(16) = "'" & mebOG(19).Text & "'"
   OG(17) = "'" & mebOG(20).Text & "'"
   
   strValues = Join(OG, ",")
   
   strSql = "DECLARE x NUMBER := 0;" & _
         " BEGIN " & _
         " SELECT NVL(TO_NUMBER(MAX(OG01)),0) INTO x FROM OUTGOING;" & _
         " INSERT INTO OUTGOING (" & strCols & ") VALUES(" & strValues & ");" & _
         " END;"
         
   insertdata = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec

   strSql = "SELECT MAX(OG01) FROM OUTGOING"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   
   If Not (rsQuery.BOF And rsQuery.EOF) Then
      cur_OG01 = rsQuery.Fields(0).Value
   End If
   
   cnnConnection.CommitTrans
   insertdata = True
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtQry_GotFocus(Index As Integer)
   If txtQry(Index).Locked = False Then
      TextInverse txtQry(Index)
      If txtQry(Index).Locked = False Then
         'edit by nickc 2007/07/11 切換輸入法改用API
         'txtQry(Index).IMEMode = 2
         CloseIme
      End If
   End If
End Sub

Private Sub txtQry_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtQry(Index).Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
         Case 1, 2
         '外出日期:只可為數字
            If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
         Case 3, 4
         '外出人員部門:只可為文數字
            If Not (KeyAscii = 8 Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
      End Select
   End If
End Sub

Private Sub txtQry_LostFocus(Index As Integer)
   If SSTab1.Tab = 0 Then
      Dim bolCancel As Boolean
      bolCancel = False
      Call txtQry_Validate(Index, bolCancel)
      If bolCancel = True Then
         SSTab1.Tab = 1
         txtQry(Index).SetFocus
      End If
   End If
End Sub

Private Sub txtQry_Validate(Index As Integer, Cancel As Boolean)
   If txtQry(Index).Locked = False Then
      Select Case Index
         Case 1
            '外出日期起
            If PUB_CheckKeyInDate(txtQry(Index)) <> 0 Then Cancel = True
         Case 2
            '外出日期迄
            'edit by nickc 2005/03/16
            'If PUB_CheckKeyInDate(txtQry(Index)) <> 0 Then
            If PUB_CheckKeyInDate(txtQry(Index)) = 0 Then
               'Modify by Morgan 2010/8/17 百年蟲
               'If txtQry(1) <> "" And (txtQry(2) < txtQry(1)) Then
               If txtQry(1) <> "" And Val(txtQry(2)) < Val(txtQry(1)) Then
                  MsgBox "外出日期迄日必需大於起日！", vbCritical
                  Cancel = True
               End If
            'add by nickc 2005/03/16
            Else
                Cancel = True
            End If
         Case 4
            '外出部門
            If txtQry(3) <> "" And txtQry(4) < txtQry(3) Then
               MsgBox "外出部門迄值必需大於起值！", vbCritical
               Cancel = True
            End If
      End Select
      If Cancel = True Then txtQry_GotFocus (Index)
   End If
End Sub

'Add by Amy 2014/12/25 71011登入顯示副本List
Private Sub SetMailCC()
    Dim RsQ As New ADODB.Recordset
    Dim strQuery As String
    Dim intR As Integer
    Dim strWhere(2) As String 'Modify by Amy 2016/07/18
    Dim i As Integer, stTemp As String 'Add by Amy 2016/05/19
    Dim stSignMan As String 'Add by Amy 2016/07/18 第一、二人事簽核主管
   
    lstMailCC.Clear
    bolOAgent = False
    strABS001_1 = "": strABS001_2 = "": strABS001_3 = ""
    strB0124 = "" 'Add by Amy 2021/01/11 請假、出差核准後通知人員編號(多筆;區隔)
    
    'Modify by Amy 2016/03/04 原只開放王副總71011,改開放專利處
    If Left(Pub_StrUserSt03, 2) = "P1" Or Pub_StrUserSt03 = "M51" Then
        If Pub_StrUserSt03 = "M51" Then Label1(10).Caption = "專利處外出副本："
        Label1(10).Visible = True
        lstMailCC.Visible = True
       
        'Modify by Amy 2016/03/04 先抓案件職代,沒案件職代抓人事職代,且排前面非職代改員編排
        'Modify by Amy 2021/01/11 +請假、出差核准後通知人員編號(多筆;區隔),算「職代」等級
        Call GetABS001_CaseSys(strUserNum, strABS001_1, strABS001_2, strABS001_3, strB0124)
        'Modify by Amy 2016/07/18 加勾選 第一、二人事簽核主管,顯示-職代(>雅娟)>審核主管>其他
        If strABS001_1 <> MsgText(601) Or strABS001_2 <> MsgText(601) Or strABS001_3 <> MsgText(601) Or strB0124 <> MsgText(601) Then
            bolOAgent = True
            'Modify by Amy 2016/05/19 職代離職語法有誤造成當掉
            If strABS001_1 <> MsgText(601) Then
                'strWhere(0) = "'" & Replace(strABS001_1, ",", "','") & "'"
                stTemp = "," & strABS001_1
            End If
            If strABS001_2 <> MsgText(601) Then
                'strWhere(0) = strWhere(0) & IIf(strWhere(0) = MsgText(601), "'", ",'") & Replace(strABS001_2, ",", "','") & "'"
                stTemp = stTemp & "," & strABS001_2
            End If
            If strABS001_3 <> MsgText(601) Then
                'strWhere(0) = strWhere(0) & IIf(strWhere(0) = MsgText(601), "'", ",'") & Replace(strABS001_3, ",", "','") & "'"
                stTemp = stTemp & "," & strABS001_3
            End If
            'end 2016/05/19
            If strB0124 <> MsgText(601) Then
                stTemp = stTemp & "," & Replace(strB0124, ";", ",")
            End If
            strWhere(0) = Replace(Mid(stTemp, 2), ",", "','")
            strWhere(1) = stTemp
        End If
        'end 2021/01/11
      
        stSignMan = GetABS001_2(strUserNum)
        If stSignMan <> MsgText(601) Then
            bolOAgent = True
            strWhere(2) = Replace(stSignMan, ",", "','")
            If strWhere(1) <> MsgText(601) Then
                strWhere(1) = strWhere(1) & "," & stSignMan
            Else
                strWhere(1) = "," & stSignMan
            End If
            stTemp = stTemp & "," & stSignMan
        End If
        
        strWhere(1) = Replace(Mid(strWhere(1), 2), ",", "','")
        
        'Modified by Morgan 2024/12/20 71011已退休,刪除無效判斷及控制
        'If strUserNum = "71011" Then
        '    If bolOAgent = True Then
        '        strWhere(1) = "And st01 not in ('79075'" & IIf(strWhere(1) = MsgText(601), "", ",'" & strWhere(1) & "'") & ",'" & strUserNum & "')"
        '    Else
        '        strWhere(1) = "And st01 not in ('71011','73022','79075'" & IIf(strWhere(2) = MsgText(601), "", ",'" & strWhere(2) & "'") & ",'" & strUserNum & "')"
        '    End If
        'Else
        '    strWhere(1) = "And st01 not in ('" & strUserNum & "'" & IIf(strWhere(1) = MsgText(601), "", ",'" & strWhere(1) & "'") & ")"
        'End If
        strWhere(1) = "And st01 not in ('" & strUserNum & "'" & IIf(strWhere(1) = MsgText(601), "", ",'" & strWhere(1) & "'") & ")"
        'end 2024/12/20
        
        If InStr(strWhere(2), ",") > 0 Then
            strWhere(2) = "st01 in ('" & strWhere(2) & "')"
        Else
            strWhere(2) = "st01 ='" & stSignMan & "'"
        End If
       
        'Removed by Morgan 2024/12/20 71011已退休,刪除無效判斷及控制
        'If strUserNum = "71011" Then
        '    'Modify by Amy 2016/03/04 顯示順序 案件職代->雅娟->其他人員員編
        '    'Modify by Amy 2016/07/18 顯示順序 案件職代->雅娟->審核主管->其他人員員編
        '    If bolOAgent = True Then
        '        strQuery = "Select st01,st02,'1' as Sort From Staff Where st01 in ('" & strWhere(0) & "') " & _
        '              " Union Select st01,st02,'2' as Sort From Staff Where st01='79075' " & IIf(strWhere(0) = MsgText(601), "", " And st01 not in ('" & strWhere(0) & "')") & _
        '              " Union Select st01,st02,'3' as Sort From Staff Where " & strWhere(2) & " And st01 not in ('97075'" & IIf(strWhere(0) = MsgText(601), "", ",'" & strWhere(0) & "'") & ")" & _
        '              " Union Select st01,st02,'4' as Sort From Staff Where Substr(st03,1,2)='P1' And st04='1' " & strWhere(1) & " And st01<'F' And Substr(st01,4,1)<>'9' "
        '    Else
        '        '顯示順序 73022(游登銘)、79075(郭雅娟)、審核主管、所有在職且為部門為P1開頭的人員(以姓名排)
        '        strQuery = "Select st01,st02,'1' as Sort From Staff Where st01='73022' " & _
        '              " Union Select st01,st02,'2' as Sort From Staff Where st01='79075' " & _
        '              " Union Select st01,st02,'3' as Sort From Staff Where " & strWhere(2) & " And st01 not in ('73022','79075')" & _
        '              " Union Select st01,st02,'4'||st02 as Sort From Staff Where Substr(st03,1,2)='P1' And st04='1' " & strWhere(1) & "And st01<'F' And Substr(st01,4,1)<>'9' " & _
        '              " Order by Sort "
        '    End If
        '    stTemp = stTemp & ",79075"
        'Else
        'end 2024/12/20
        
            strQuery = "Select st01,st02,'1' as Sort From Staff Where st01 in ('" & strWhere(0) & "') " & _
                " Union Select st01,st02,'2' as Sort From Staff Where " & strWhere(2) & IIf(strWhere(0) = MsgText(601), "", " And st01 not in ('" & strWhere(0) & "')") & _
                " Union Select st01,st02,'3' as Sort From Staff Where Substr(st03,1,2)='P1' And st04='1' " & strWhere(1) & " And st01<'F' And Substr(st01,4,1)<>'9' "
           
        
        'End If 'Modified by Morgan 2024/12/2
        
        If InStr(strQuery, "Order by") = 0 Then strQuery = "Select * From (" & strQuery & ") Order by Sort,st01"
        'end 2016/03/04
        'end 2016/07/18
        
        intR = 1: i = 0
        Set RsQ = ClsLawReadRstMsg(intR, strQuery)
        If intR = 1 Then
            RsQ.MoveFirst
            Do While RsQ.EOF = False
                lstMailCC.AddItem RsQ.Fields("st02") & " " & RsQ.Fields("st01")
                'Add by Amy 2016/05/19 由FormLoad搬過來修改
                If bolOAgent = True Then
                    If InStr(stTemp, "" & RsQ.Fields("st01")) > 0 Then lstMailCC.Selected(i) = True
                    i = i + 1
                End If
                RsQ.MoveNext
            Loop
        End If
    End If
End Sub

'Added by Lydia 2016/08/11 點選欄位進行排序
Private Sub grdList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   'Modified by Lydia 2022/01/11
   'Pub_MSFGrdColRow grdList, x, y, nCol, nRow
   getGrdColRow grdList, x, y, nCol, nRow
   
   If nCol < 0 Or nRow < 0 Then Exit Sub
   
   grdList.col = nCol
   grdList.row = nRow
   If Me.grdList.row < 1 Then
      If InStr("序號,外出時間", Me.grdList.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.grdList.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdList.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grdList.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
'end 2016/08/11

'Add By Sindy 2021/2/5
Private Sub txtST01_GotFocus()
   TextInverse Me.txtST01
End Sub
Private Sub txtST01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtST01_Validate(Cancel As Boolean)
   If Me.txtST01.Text = "" Then
      Me.Label5.Caption = ""
      Exit Sub
   End If
   If ByInputGetST01or02(Me.txtST01.Text, strExc(0), strExc(1)) = False Then
      Cancel = True
      Me.txtST01.SetFocus
   End If
   Me.txtST01.Text = strExc(0)
   Me.Label5.Caption = strExc(1)
   If Cancel = True Then txtST01_GotFocus
End Sub
'2021/2/5 END

'Added by Lydia 2022/05/18
Private Sub txtCode_GotFocus(Index As Integer)
    TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
