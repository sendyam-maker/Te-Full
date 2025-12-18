VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050722 
   BorderStyle     =   1  '單線固定
   Caption         =   "定稿特殊請款文字維護"
   ClientHeight    =   5720
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   8190
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   30
      TabIndex        =   11
      Top             =   630
      Width           =   8120
      _ExtentX        =   14323
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm050722.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LblLST02"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblLST01"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblLST10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textLST11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textLST10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textLST03"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textLST02"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textLST01"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt1(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(14)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm050722.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "txt1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(13)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(12)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(11)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txt1(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblFM2(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblFM2(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(5)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(6)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txt1(3)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "LblCnt"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(15)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdok"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "GRD1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   280
         Left            =   -68220
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "|#(底線)  #|"
         Top             =   3090
         Width           =   1090
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm050722.frx":0038
         Height          =   3410
         Left            =   30
         TabIndex        =   12
         Top             =   1530
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   6015
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H008080FF&
         Caption         =   "查詢"
         Height          =   345
         Left            =   5670
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(含排除的申請人)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   7
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   140
         Index           =   15
         Left            =   0
         TabIndex        =   35
         Top             =   870
         Width           =   1060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "文字中有""發文日期""時，引用時會抓取發文日期，範例：Y2049000"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   14
         Left            =   -74400
         TabIndex        =   34
         Top             =   3330
         Width           =   5140
      End
      Begin MSForms.TextBox txt1 
         Height          =   640
         Index           =   4
         Left            =   -74400
         TabIndex        =   33
         Top             =   3690
         Width           =   4190
         VariousPropertyBits=   -1466941409
         ForeColor       =   192
         Size            =   "7391;1129"
         Value           =   $"frm050722.frx":004D
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "文字下方欲加底線時，字的前後要加系統Tag，範例：|#(底線)APsto@vinge.se#|"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   8
         Left            =   -74400
         TabIndex        =   32
         Top             =   3120
         Width           =   6050
      End
      Begin VB.Label LblCnt 
         AutoSize        =   -1  'True
         Caption         =   "(共 0 筆)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   7020
         TabIndex        =   31
         Top             =   1290
         Width           =   660
      End
      Begin MSForms.TextBox textLST01 
         Height          =   300
         Left            =   -73500
         TabIndex        =   0
         Top             =   420
         Width           =   870
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1535;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLST02 
         Height          =   300
         Left            =   -73500
         TabIndex        =   1
         Top             =   750
         Width           =   870
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1535;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLST03 
         Height          =   920
         Left            =   -73500
         TabIndex        =   2
         Top             =   1080
         Width           =   5820
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "10266;1623"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLST10 
         Height          =   300
         Left            =   -73500
         TabIndex        =   3
         Top             =   2040
         Width           =   870
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1535;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textLST11 
         Height          =   300
         Left            =   -73500
         TabIndex        =   4
         Top             =   2370
         Width           =   5820
         VariousPropertyBits=   -1466941413
         MaxLength       =   50
         ScrollBars      =   2
         Size            =   "10266;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   7
         Left            =   -74040
         TabIndex        =   30
         Top             =   2430
         Width           =   540
      End
      Begin MSForms.TextBox txt1 
         Height          =   290
         Index           =   3
         Left            =   5520
         TabIndex        =   8
         Top             =   360
         Width           =   2480
         VariousPropertyBits=   680542235
         ScrollBars      =   2
         Size            =   "4374;512"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   6
         Left            =   4950
         TabIndex        =   29
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(模糊比對)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   7
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   140
         Index           =   5
         Left            =   5550
         TabIndex        =   28
         Top             =   660
         Width           =   640
      End
      Begin MSForms.Label LblLST10 
         Height          =   260
         Left            =   -72580
         TabIndex        =   27
         Top             =   2070
         Width           =   5480
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9666;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "排除的申請人："
         Height          =   180
         Index           =   4
         Left            =   -74760
         TabIndex        =   26
         Top             =   2100
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "註：無申請人或代理人時，會存 0（代理人編號或申請人編號僅能存入6碼或8碼）"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   25
         Top             =   2850
         Width           =   6410
      End
      Begin MSForms.Label Label23 
         Height          =   200
         Left            =   -74850
         TabIndex        =   24
         Top             =   4620
         Width           =   7700
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13582;353"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   260
         Index           =   1
         Left            =   2160
         TabIndex        =   23
         Top             =   690
         Width           =   2420
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "4269;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   260
         Index           =   0
         Left            =   2160
         TabIndex        =   22
         Top             =   380
         Width           =   2420
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "4269;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   290
         Index           =   1
         Left            =   1020
         TabIndex        =   6
         Top             =   680
         Width           =   1100
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1940;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   290
         Index           =   0
         Left            =   1020
         TabIndex        =   5
         Top             =   360
         Width           =   1100
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1940;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   21
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   20
         Top             =   410
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(模糊比對)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   7
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   140
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   1250
         Width           =   640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "文字內容："
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   18
         Top             =   1040
         Width           =   900
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   2
         Left            =   1020
         TabIndex        =   7
         Top             =   980
         Width           =   4190
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "7391;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "文字內容："
         Height          =   180
         Index           =   1
         Left            =   -74400
         TabIndex        =   17
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   3
         Left            =   -74220
         TabIndex        =   16
         Top             =   810
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   2
         Left            =   -74220
         TabIndex        =   15
         Top             =   480
         Width           =   690
      End
      Begin MSForms.Label LblLST01 
         Height          =   260
         Left            =   -72580
         TabIndex        =   14
         Top             =   440
         Width           =   5480
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9666;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LblLST02 
         Height          =   260
         Left            =   -72580
         TabIndex        =   13
         Top             =   780
         Width           =   5480
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9666;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm050722.frx":00AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":03C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":06E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":08C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":0BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":0EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":1214
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":1530
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":184C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":1B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050722.frx":1E84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
End
Attribute VB_Name = "frm050722"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Sindy 2025/6/2
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區1
Dim m_EditMode As Integer
'(執行各項功能的權限)
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
' 第一筆資料的本所案號
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_LST As Integer
Dim MyKind As String


Private Sub cmdok_Click()
   LblCnt.Caption = "(共 0 筆)"
   GRD1.Clear
   SetGrd
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
       GetData
   Else
       MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
       txt1(0).SetFocus
   End If
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from LetterSetText where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_LST = rsA.Fields.Count
   SetGrd
End Sub

' 按下按鍵
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub
'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ReDim m_FieldList(tf_LST) As FIELDITEM

   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   textLST01.BackColor = &H8000000F
   
   MoveFormToCenter Me

   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
   
   LblLST01.Caption = ""
   LblLST02.Caption = ""
   LblLST10.Caption = ""
   lblFM2(0).Caption = ""
   lblFM2(1).Caption = ""
'   Tbar1_ButtonClick TBar1.Buttons(4) '設定為按下查詢鍵
'   OnAction vbKeyF9 '按確定
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050722 = Nothing
End Sub

Private Sub GRD1_DblClick()
Dim tmpMouseRow
Dim i, j

   GRD1.Visible = False
   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      If GRD1.CellBackColor = &HFFC0C0 Then
         textLST01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textLST02.Text = GRD1.TextMatrix(tmpMouseRow, 2)
         Me.SSTab1.Tab = 0
         QueryRecord
      End If
   End If
   GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j

   GRD1.Visible = False
   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      If GRD1.CellBackColor <> &HFFC0C0 Then
         For j = 1 To GRD1.Rows - 1
            GRD1.row = j
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15)
            Next i
         Next j
         GRD1.row = tmpMouseRow
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
         textLST01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textLST02.Text = GRD1.TextMatrix(tmpMouseRow, 2)
         QueryRecord
      End If
   End If
   GRD1.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      cmdOK.SetFocus
      cmdOK.Default = True
   Else
      cmdOK.Default = False
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
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

   If IsNull(rsSrcTmp.Fields("LST04")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LST04")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("LST04"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LST05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LST05")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("LST05"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LST06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LST06")) = False Then
         strTemp = rsSrcTmp.Fields("LST06")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LST07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LST07")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("LST07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LST08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LST08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("LST08"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LST09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LST09")) = False Then
         strTemp = rsSrcTmp.Fields("LST09")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If

   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False

   If textLST01.Text = "0" And textLST02.Text = "0" Then
      MsgBox "代理人或申請人至少要輸入一個！", vbExclamation
      textLST01.SetFocus
      Exit Function
   End If
   
   If textLST03.Text = "" Then
      MsgBox "文字內容不可以空白！", vbExclamation
      textLST03.SetFocus
      Exit Function
   End If
   
   If textLST02.Text <> "" And textLST02.Text <> "0" And textLST10.Text <> "" Then
      MsgBox "申請人和要排除的申請人，不可同時輸入！", vbExclamation
      'textLST10.SetFocus
      Exit Function
   End If
   If textLST10.Text = "0" Then
      MsgBox "要排除的申請人，不可輸入0！", vbExclamation
      textLST10.SetFocus
      Exit Function
   End If

   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer

   For nIndex = 0 To tf_LST - 1
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

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String

   For nIndex = 0 To tf_LST - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
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
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strLST01 As String
Dim strLST02 As String
   
   AddRecord = False
   
   strLST01 = textLST01
   strLST02 = textLST02
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strLST01, strLST02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO LetterSetText ("
   For nIndex = 0 To tf_LST - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
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
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To tf_LST - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If ((strLST01 & strLST02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strLST01 & strLST02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strLST01, strLST02
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
    
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
   Dim strLST01 As String
   Dim strLST02 As String

   ModRecord = False

   strLST01 = m_CurrKEY(0)
   strLST02 = m_CurrKEY(1)

   strSql = "begin user_data.user_enabled:=1; UPDATE LetterSetText SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_LST - 1
      strTmp = Empty
      'If nIndex < 3 Or nIndex > 8 Then
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
        'End If
   Next nIndex

   strSql = strSql & " " & _
                  "WHERE LST01 = " & CNULL(strLST01) & " and LST02 = " & CNULL(strLST02) & "; end; "
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   If bDifference = True Then
      'Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans

   ShowCurrRecord strLST01, strLST02

   ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strLST01 As String
Dim strLST02 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strLST01 = m_CurrKEY(0)
   strLST02 = m_CurrKEY(1)

   strSql = "DELETE FROM LetterSetText " & _
            "WHERE LST01 = '" & strLST01 & "'  and LST02='" & strLST02 & "' "
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If (strLST01 = m_LastKEY(0) And strLST02 = m_LastKEY(1)) Or (strLST01 = m_FirstKEY(0) And strLST02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strLST01, strLST02
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False
   If IsRecordExist(textLST01, textLST02) = True Then
      m_CurrKEY(0) = textLST01
      m_CurrKEY(1) = textLST02
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse

   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textLST01 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textLST01.SetFocus
      Case 2: If Me.Visible = True Then textLST03.SetFocus
      Case 4: If Me.Visible = True Then textLST01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   IsRecordExist = False
   strSql = "SELECT * FROM LetterSetText " & _
            "WHERE LST01 = " & CNULL(strKEY01) & " and LST02 = " & CNULL(strKEY02)
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

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT * FROM LetterSetText " & _
               "WHERE LST01 = " & CNULL(m_CurrKEY(0)) & " and LST02 = " & CNULL(m_CurrKEY(1))
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("LST01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("LST01")
         If IsNull(rsTmp.Fields("LST02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("LST02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close

      strSql = "SELECT * FROM LetterSetText order by LST01 asc,LST02 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("LST01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("LST01")
         If IsNull(rsTmp.Fields("LST02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("LST02")
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
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM LetterSetText " & _
            "WHERE LST01 = '" & m_CurrKEY(0) & "' AND " & _
                  "LST02 = (SELECT MAX(LST02) FROM LetterSetText " & _
                          "WHERE LST01 = '" & m_CurrKEY(0) & "' AND " & _
                                "LST02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LST01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("LST01")
      If IsNull(rsTmp.Fields("LST02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("LST02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT * FROM LetterSetText " & _
            "WHERE LST01 = (SELECT MAX(LST01) FROM LetterSetText " & _
                           "WHERE LST01 < '" & m_CurrKEY(0) & "') AND " & _
                  "LST02 = (SELECT MAX(LST02) FROM LetterSetText " & _
                           "WHERE LST01 = (SELECT MAX(LST01) FROM LetterSetText " & _
                                          "WHERE LST01 < '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LST01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("LST01")
      If IsNull(rsTmp.Fields("LST02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("LST02")
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

   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If

   strSql = "SELECT LST01,LST02 FROM LetterSetText " & _
            "WHERE LST01 = '" & m_CurrKEY(0) & "' AND " & _
                  "LST02 = (SELECT MIN(LST02) FROM LetterSetText " & _
                          "WHERE LST01 = '" & m_CurrKEY(0) & "' AND " & _
                                "LST02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LST01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("LST01")
      If IsNull(rsTmp.Fields("LST02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("LST02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT LST01,LST02 FROM LetterSetText " & _
            "WHERE LST01 = (SELECT MIN(LST01) FROM LetterSetText " & _
                           "WHERE LST01 > '" & m_CurrKEY(0) & "') AND " & _
                  "LST02 = (SELECT MIN(LST02) FROM LetterSetText " & _
                           "WHERE LST01 = (SELECT MIN(LST01) FROM LetterSetText " & _
                                          "WHERE LST01 > '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LST01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("LST01")
      If IsNull(rsTmp.Fields("LST02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("LST02")
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
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse

   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
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
         UpdateFieldNewData
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
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
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   strSql = "SELECT LST01,LST02 FROM LetterSetText " & _
            "WHERE LST01 = (SELECT MIN(LST01) FROM LetterSetText) AND " & _
                  "LST02 = (SELECT MIN(LST02) FROM LetterSetText " & _
                           "WHERE LST01 = (SELECT MIN(LST01) FROM LetterSetText)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LST01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("LST01")
      If IsNull(rsTmp.Fields("LST02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("LST02")
   End If
   rsTmp.Close

   strSql = "SELECT LST01,LST02 FROM LetterSetText " & _
            "WHERE LST01 = (SELECT MAX(LST01) FROM LetterSetText) AND " & _
                  "LST02 = (SELECT MAX(LST02) FROM LetterSetText " & _
                           "WHERE LST01 = (SELECT MAX(LST01) FROM LetterSetText)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LST01")) = False Then: m_LastKEY(0) = rsTmp.Fields("LST01")
      If IsNull(rsTmp.Fields("LST02")) = False Then: m_LastKEY(1) = rsTmp.Fields("LST02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer

   strSql = "SELECT * FROM LetterSetText " & _
            "WHERE LST01=" & CNULL(m_CurrKEY(0)) & " and LST02=" & CNULL(m_CurrKEY(1))
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("LST01")) = False Then: textLST01 = rsTmp.Fields("LST01"): Call textLST01_Validate(False)
      If IsNull(rsTmp.Fields("LST02")) = False Then: textLST02 = rsTmp.Fields("LST02"): Call textLST02_Validate(False)
      If IsNull(rsTmp.Fields("LST03")) = False Then: textLST03 = rsTmp.Fields("LST03")
      If IsNull(rsTmp.Fields("LST10")) = False Then: textLST10 = rsTmp.Fields("LST10"): Call textLST10_Validate(False)
      If IsNull(rsTmp.Fields("LST11")) = False Then: textLST11 = rsTmp.Fields("LST11")
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If
   rsTmp.Close

EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
Dim i As Integer
Dim strNoTemp As String, strTemp As String

   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and instr(LST01," & CNULL(txt1(0)) & ")>0"
   End If
   '申請人
   If txt1(1) <> "" Then
       strSql = strSql & " and (instr(LST02," & CNULL(txt1(1)) & ")>0 or instr(LST10," & CNULL(txt1(1)) & ")>0)"
   End If
   If txt1(2) <> "" Then
       strSql = strSql & " and instr(LST03," & CNULL(txt1(2)) & ")>0"
   End If
   If txt1(3) <> "" Then
       strSql = strSql & " and instr(LST11," & CNULL(txt1(3)) & ")>0"
   End If
   Screen.MousePointer = vbHourglass
   '抓取資料
   strSql = "SELECT LST01,'',LST02,'',LST03,LST11" & _
            " FROM LetterSetText where 1=1 " & strSql & _
            " order by LST01 asc,LST02 asc "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   SetGrd
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      LblCnt.Caption = "(共 " & rsTmp.RecordCount & " 筆)"
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   If rsTmp.RecordCount > 0 Then
      For i = 1 To GRD1.Rows - 1
         GRD1.row = i
         '代理人
         If GRD1.TextMatrix(i, 0) <> "0" And GRD1.TextMatrix(i, 0) <> "" Then
            strNoTemp = ChangeCustomerL(GRD1.TextMatrix(i, 0))
            If ClsPDGetAgent(strNoTemp, strTemp) Then
               GRD1.col = 1
               GRD1.Text = strTemp
            End If
         End If
         '申請人
         If GRD1.TextMatrix(i, 2) <> "0" And GRD1.TextMatrix(i, 2) <> "" Then
            strNoTemp = ChangeCustomerL(GRD1.TextMatrix(i, 2))
            If ClsPDGetCustomer(strNoTemp, strTemp) Then
               GRD1.col = 3
               GRD1.Text = strTemp
            End If
         End If
      Next i
      SetGrd
      GRD1.col = 0
      GRD1.row = 1
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   Screen.MousePointer = vbDefault
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
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
         ' 新增
      Case 1, 2, 3, 4:
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

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTmp  As String

   CheckDataValid = False
   
   nResponse = False
   textLST01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textLST02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   textLST10_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   'textLST11_Validate nResponse
   'If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textLST01.Locked = bEnable
   textLST02.Locked = bEnable
   If bEnable Then textLST01.BackColor = &H8000000F Else textLST01.BackColor = &H80000005
   If bEnable Then textLST02.BackColor = &H8000000F Else textLST02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer

   textLST01.Locked = bEnable
   textLST02.Locked = bEnable
   textLST03.Locked = bEnable
   textLST10.Locked = bEnable
   textLST11.Locked = bEnable
   If bEnable Then textLST01.BackColor = &H8000000F Else textLST01.BackColor = &H80000005
   If bEnable Then textLST02.BackColor = &H8000000F Else textLST02.BackColor = &H80000005
   If bEnable Then textLST03.BackColor = &H8000000F Else textLST03.BackColor = &H80000005
   If bEnable Then textLST10.BackColor = &H8000000F Else textLST10.BackColor = &H80000005
   If bEnable Then textLST11.BackColor = &H8000000F Else textLST11.BackColor = &H80000005
End Sub

Private Sub ClearField()
Dim nIndex As Integer

   textLST01 = Empty
   textLST02 = Empty
   textLST03 = Empty
   textLST10 = Empty
   textLST11 = Empty
   LblLST01 = Empty
   LblLST02 = Empty
   LblLST10 = Empty
   SetGrd
   For nIndex = 0 To tf_LST - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      If Trim(textLST01) = "" Then textLST01 = "0"
      SetFieldNewData "LST01", textLST01
      If Trim(textLST02) = "" Then textLST02 = "0"
      SetFieldNewData "LST02", textLST02
   End If
   SetFieldNewData "LST03", textLST03
   SetFieldNewData "LST10", textLST10
   SetFieldNewData "LST11", textLST11
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String

   ' 初始化欄位陣列
   For nIndex = 1 To tf_LST
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "LST" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 99:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textLST01_GotFocus()
   If textLST01.Enabled = True And textLST01.Locked = False Then
      InverseTextBox textLST01
      CloseIme
   End If
End Sub

Private Sub textLST01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If textLST01.Enabled = True And textLST01.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub textLST01_LostFocus()
   If textLST01 <> "" And m_EditMode = 1 Then
      If Len(textLST01) = 6 Then
         strExc(10) = textLST01 & "00"
      ElseIf Mid(textLST01, 7) = "00" Then
         strExc(10) = Mid(textLST01, 1, 6)
      End If
      strSql = "SELECT * FROM LetterSetText " & _
               "WHERE LST01 = " & CNULL(strExc(10))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If MsgBox("代理人編號（" & strExc(10) & "）已存在，確定要繼續嗎？", _
            vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            textLST01.SetFocus
         End If
      End If
   End If
End Sub

'代理人
Private Sub textLST01_Validate(Cancel As Boolean)
   Cancel = ChkLST01(textLST01, LblLST01)
End Sub
Private Function ChkLST01(objFaNo As Object, objFaName As Object) As Boolean
Dim strNoTemp As String, strTemp As String
   
   ChkLST01 = False
   objFaName.Caption = ""
   If objFaNo <> "" Then
      If objFaNo = "0" Then Exit Function '無代理人時,存0
      If Left(objFaNo, 1) = 代理人編號 Then
         '加碼數檢查
         If Len(objFaNo) = 6 Or Len(objFaNo) = 8 Then
            strNoTemp = ChangeCustomerL(objFaNo)
            If ClsPDGetAgent(strNoTemp, strTemp) Then
               objFaName.Caption = strTemp
'               '若為6碼，統一補足為8碼。
'               If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
'                  objFaNo = Left(ChangeCustomerL(objFaNo), 8)
'               End If
            Else
               If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
                  ChkLST01 = True
                  objFaNo.SetFocus
                  InverseTextBox objFaNo
               End If
            End If
         Else
            MsgBox "代理人編號只可輸入6碼或8碼！", vbCritical
            If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
               ChkLST01 = True
               objFaNo.SetFocus
               InverseTextBox objFaNo
            End If
         End If
      Else
         MsgBox "請輸入代理人編號！", vbCritical
         If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
            ChkLST01 = True
            objFaNo.SetFocus
            InverseTextBox objFaNo
         End If
      End If
   End If
End Function

Private Sub textLST02_GotFocus()
   If textLST02.Enabled = True And textLST02.Locked = False Then
      InverseTextBox textLST02
      CloseIme
   End If
End Sub

Private Sub textLST02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If textLST02.Enabled = True And textLST02.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub textLST02_LostFocus()
   If textLST02 <> "" And m_EditMode = 1 Then
      If Len(textLST02) = 6 Then
         strExc(10) = textLST02 & "00"
      ElseIf Mid(textLST02, 7) = "00" Then
         strExc(10) = Mid(textLST02, 1, 6)
      End If
      strSql = "SELECT * FROM LetterSetText " & _
               "WHERE LST02 = " & CNULL(strExc(10))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If MsgBox("申請人編號（" & strExc(10) & "）已存在，確定要繼續嗎？", _
            vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            textLST02.SetFocus
         End If
      End If
   End If
End Sub

'申請人
Private Sub textLST02_Validate(Cancel As Boolean)
   Cancel = ChkLST02(textLST02, LblLST02)
End Sub
Private Function ChkLST02(objCuNo As Object, objCuName As Object) As Boolean
Dim strNoTemp As String, strTemp As String
   
   ChkLST02 = False
   objCuName.Caption = ""
   If objCuNo <> "" Then
      If objCuNo = "0" Then Exit Function '無申請人時,存0
      If Left(objCuNo, 1) = 客戶編號 Then
         '加碼數檢查
         If Len(objCuNo) = 6 Or Len(objCuNo) = 8 Then
            strNoTemp = ChangeCustomerL(objCuNo)
            If ClsPDGetCustomer(strNoTemp, strTemp) Then
               objCuName.Caption = strTemp
'               '若為6碼，統一補足為8碼。
'               If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
'                  objCuNo = Left(ChangeCustomerL(objCuNo), 8)
'               End If
            Else
               If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
                  ChkLST02 = True
                  objCuNo.SetFocus
                  InverseTextBox objCuNo
               End If
            End If
         Else
            MsgBox "客戶編號只可輸入6碼或8碼！", vbCritical
            If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
               ChkLST02 = True
               objCuNo.SetFocus
               InverseTextBox objCuNo
            End If
         End If
      Else
         MsgBox "請輸入客戶編號！", vbCritical
         If m_EditMode <> 0 Or Me.SSTab1.Tab = 1 Then
            ChkLST02 = True
            objCuNo.SetFocus
            InverseTextBox objCuNo
         End If
      End If
   End If
End Function

Private Sub textLST10_GotFocus()
   If textLST10.Enabled = True And textLST10.Locked = False Then
      InverseTextBox textLST10
      CloseIme
   End If
End Sub

Private Sub textLST10_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If textLST10.Enabled = True And textLST10.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

'排除的申請人
Private Sub textLST10_Validate(Cancel As Boolean)
   Cancel = ChkLST02(textLST10, LblLST10)
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

   '                        0         1             2         3             4           5
   arrGridHeadText = Array("代理人", "代理人名稱", "申請人", "申請人名稱", "文字內容", "備註")
   arrGridHeadWidth = Array(800, 1200, 800, 1200, 3000, 800)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 0, 1
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then
      If Index = 0 Or Index = 1 Then
         lblFM2(Index).Caption = ""
      End If
      Exit Sub
   End If
   Select Case Index
      Case 0
         Cancel = ChkLST01(txt1(Index), lblFM2(Index))
      Case 1
         Cancel = ChkLST02(txt1(Index), lblFM2(Index))
      Case Else
   End Select
End Sub
