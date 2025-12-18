VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880007 
   BorderStyle     =   1  '單線固定
   Caption         =   "變更事項"
   ClientHeight    =   5760
   ClientLeft      =   165
   ClientTop       =   990
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   7728
      TabIndex        =   62
      Top             =   50
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   6792
      TabIndex        =   61
      Top             =   50
      Width           =   912
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   120
      TabIndex        =   26
      Top             =   540
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9022
      _Version        =   393216
      TabsPerRow      =   10
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm880007.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPetition(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblPetition(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblPetition(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPetition(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblPetition(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label22"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblTrademarkKind"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCaseField(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCaseField(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCaseField(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCaseField(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCaseField(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCaseField(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCaseField(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCaseField(8)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCaseField(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtCaseField(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCaseField(12)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCaseField(11)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtCaseField(13)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkChoose(7)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkChoose(11)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkChoose(10)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "chkChoose(9)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "chkChoose(8)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "chkChoose(6)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "chkChoose(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "chkChoose(2)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chkChoose(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "chkChoose(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "chkChoose(4)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "chkChoose(3)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtCaseField(7)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm880007.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkChoose(12)"
      Tab(1).Control(1)=   "chkChoose(13)"
      Tab(1).Control(2)=   "chkChoose(14)"
      Tab(1).Control(3)=   "chkChoose(16)"
      Tab(1).Control(4)=   "txtCaseField(42)"
      Tab(1).Control(5)=   "txtCaseField(41)"
      Tab(1).Control(6)=   "txtCaseField(25)"
      Tab(1).Control(7)=   "txtCaseField(24)"
      Tab(1).Control(8)=   "txtCaseField(22)"
      Tab(1).Control(9)=   "txtCaseField(23)"
      Tab(1).Control(10)=   "txtCaseField(21)"
      Tab(1).Control(11)=   "txtCaseField(20)"
      Tab(1).Control(12)=   "txtCaseField(19)"
      Tab(1).Control(13)=   "txtCaseField(14)"
      Tab(1).Control(14)=   "txtCaseField(15)"
      Tab(1).Control(15)=   "txtCaseField(16)"
      Tab(1).Control(16)=   "txtCaseField(17)"
      Tab(1).Control(17)=   "txtCaseField(18)"
      Tab(1).Control(18)=   "Label15"
      Tab(1).Control(19)=   "Label10"
      Tab(1).Control(20)=   "Label9"
      Tab(1).Control(21)=   "Label14"
      Tab(1).Control(22)=   "Label12"
      Tab(1).Control(23)=   "Label5"
      Tab(1).Control(24)=   "Label8"
      Tab(1).Control(25)=   "Label11"
      Tab(1).Control(26)=   "Label13"
      Tab(1).Control(27)=   "Label16"
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "第三頁"
      TabPicture(2)   =   "frm880007.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkChoose(15)"
      Tab(2).Control(1)=   "txtCaseField(26)"
      Tab(2).Control(2)=   "txtCaseField(27)"
      Tab(2).Control(3)=   "txtCaseField(28)"
      Tab(2).Control(4)=   "txtCaseField(29)"
      Tab(2).Control(5)=   "txtCaseField(30)"
      Tab(2).Control(6)=   "txtCaseField(31)"
      Tab(2).Control(7)=   "txtCaseField(32)"
      Tab(2).Control(8)=   "txtCaseField(33)"
      Tab(2).Control(9)=   "txtCaseField(34)"
      Tab(2).Control(10)=   "txtCaseField(35)"
      Tab(2).Control(11)=   "txtCaseField(36)"
      Tab(2).Control(12)=   "txtCaseField(37)"
      Tab(2).Control(13)=   "txtCaseField(38)"
      Tab(2).Control(14)=   "txtCaseField(39)"
      Tab(2).Control(15)=   "txtCaseField(40)"
      Tab(2).Control(16)=   "Label26"
      Tab(2).Control(17)=   "Label25"
      Tab(2).Control(18)=   "Label24"
      Tab(2).Control(19)=   "Label23"
      Tab(2).Control(20)=   "Label7"
      Tab(2).Control(21)=   "Label6(0)"
      Tab(2).Control(22)=   "Label17"
      Tab(2).Control(23)=   "Label6(1)"
      Tab(2).Control(24)=   "Label18"
      Tab(2).Control(25)=   "Label6(3)"
      Tab(2).Control(26)=   "Label19"
      Tab(2).Control(27)=   "Label6(4)"
      Tab(2).Control(28)=   "Label20"
      Tab(2).Control(29)=   "Label6(5)"
      Tab(2).ControlCount=   30
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   7
         Left            =   1680
         TabIndex        =   14
         Top             =   2460
         Width           =   375
         VariousPropertyBits=   671107097
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   12
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "申請人中譯文1："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   13
         Left            =   -74880
         TabIndex        =   33
         Top             =   1860
         Width           =   1455
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2566;450"
         Value           =   "0"
         Caption         =   "代表人1(中)："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   14
         Left            =   -74880
         TabIndex        =   43
         Top             =   4260
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1508;450"
         Value           =   "0"
         Caption         =   "其他："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         CausesValidation=   0   'False
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;450"
         Value           =   "0"
         Caption         =   "代表人印鑑"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   1335
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2355;450"
         Value           =   "0"
         Caption         =   "代理人"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1125
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1984;450"
         Value           =   "0"
         Caption         =   "申請日："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   1155
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2037;450"
         Value           =   "0"
         Caption         =   "申請人1："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2355;450"
         Value           =   "0"
         Caption         =   "申請人印鑑"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1455
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2566;450"
         Value           =   "0"
         Caption         =   "正商標號數："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   2460
         Width           =   1635
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2884;450"
         Value           =   "0"
         Caption         =   "專利/商標種類："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;450"
         Value           =   "0"
         Caption         =   "案件名稱(中)："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   3660
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;450"
         Value           =   "0"
         Caption         =   "減縮商品："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;450"
         Value           =   "0"
         Caption         =   "商品類別："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   24
         Top             =   4500
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;450"
         Value           =   "0"
         Caption         =   "商品組群："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         CausesValidation=   0   'False
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   15
         Top             =   2460
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1508;450"
         Value           =   "0"
         Caption         =   "圖樣"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   16
         Left            =   -74880
         TabIndex        =   40
         Top             =   3660
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "代表人中譯文1："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkChoose 
         Height          =   255
         Index           =   15
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   1635
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2884;450"
         Value           =   "0"
         Caption         =   "申請地址1(中)："
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   42
         Left            =   -73200
         TabIndex        =   42
         Top             =   3960
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   26
         Left            =   -73260
         TabIndex        =   46
         Top             =   360
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   41
         Left            =   -73200
         TabIndex        =   41
         Top             =   3660
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   495
         Index           =   13
         Left            =   1680
         TabIndex        =   25
         Top             =   4500
         Width           =   6975
         VariousPropertyBits=   -1467987939
         ScrollBars      =   2
         Size            =   "12303;873"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   11
         Left            =   1680
         TabIndex        =   21
         Top             =   3660
         Width           =   6975
         VariousPropertyBits=   671107097
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   495
         Index           =   12
         Left            =   1680
         TabIndex        =   23
         Top             =   3960
         Width           =   6975
         VariousPropertyBits=   -1467987939
         ScrollBars      =   2
         Size            =   "12303;873"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   9
         Left            =   1680
         TabIndex        =   18
         Top             =   3060
         Width           =   6975
         VariousPropertyBits=   671107097
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   10
         Left            =   1680
         TabIndex        =   19
         Top             =   3360
         Width           =   6975
         VariousPropertyBits=   671107097
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   8
         Left            =   1680
         TabIndex        =   17
         Top             =   2760
         Width           =   6975
         VariousPropertyBits=   671107097
         Size            =   "12303;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   6
         Left            =   1680
         TabIndex        =   12
         Top             =   2160
         Width           =   2835
         VariousPropertyBits=   671107097
         MaxLength       =   20
         Size            =   "5001;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   0
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   8
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   735
         Index           =   25
         Left            =   -73200
         TabIndex        =   44
         Top             =   4260
         Width           =   6855
         VariousPropertyBits=   -1467987943
         ScrollBars      =   2
         Size            =   "12091;1296"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   660
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   3
         Left            =   1260
         TabIndex        =   5
         Top             =   960
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   2
         Left            =   5460
         TabIndex        =   4
         Top             =   660
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   4
         Left            =   5460
         TabIndex        =   6
         Top             =   960
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   5
         Left            =   1260
         TabIndex        =   7
         Top             =   1260
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   27
         Left            =   -73260
         TabIndex        =   47
         Top             =   660
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   28
         Left            =   -73260
         TabIndex        =   48
         Top             =   960
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   29
         Left            =   -73260
         TabIndex        =   49
         Top             =   1260
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   30
         Left            =   -73260
         TabIndex        =   50
         Top             =   1560
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   31
         Left            =   -73260
         TabIndex        =   51
         Top             =   1860
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   32
         Left            =   -73260
         TabIndex        =   52
         Top             =   2160
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   33
         Left            =   -73260
         TabIndex        =   53
         Top             =   2460
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   34
         Left            =   -73260
         TabIndex        =   54
         Top             =   2760
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   35
         Left            =   -73260
         TabIndex        =   55
         Top             =   3060
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   36
         Left            =   -73260
         TabIndex        =   56
         Top             =   3360
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   37
         Left            =   -73260
         TabIndex        =   57
         Top             =   3660
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   38
         Left            =   -73260
         TabIndex        =   58
         Top             =   3960
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   39
         Left            =   -73260
         TabIndex        =   59
         Top             =   4260
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   40
         Left            =   -73260
         TabIndex        =   60
         Top             =   4560
         Width           =   6945
         VariousPropertyBits=   671107097
         Size            =   "12250;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   24
         Left            =   -73200
         TabIndex        =   39
         Top             =   3360
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   22
         Left            =   -73200
         TabIndex        =   37
         Top             =   2760
         Width           =   6855
         VariousPropertyBits=   671107097
         MaxLength       =   50
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   23
         Left            =   -73200
         TabIndex        =   38
         Top             =   3060
         Width           =   6855
         VariousPropertyBits=   671107097
         MaxLength       =   80
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   21
         Left            =   -73200
         TabIndex        =   36
         Top             =   2460
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   20
         Left            =   -73200
         TabIndex        =   35
         Top             =   2160
         Width           =   6855
         VariousPropertyBits=   671107097
         MaxLength       =   80
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   19
         Left            =   -73200
         TabIndex        =   34
         Top             =   1860
         Width           =   6855
         VariousPropertyBits=   671107097
         MaxLength       =   50
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   14
         Left            =   -73200
         TabIndex        =   28
         Top             =   360
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   15
         Left            =   -73200
         TabIndex        =   29
         Top             =   660
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   16
         Left            =   -73200
         TabIndex        =   30
         Top             =   960
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   17
         Left            =   -73200
         TabIndex        =   31
         Top             =   1260
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   18
         Left            =   -73200
         TabIndex        =   32
         Top             =   1560
         Width           =   6855
         VariousPropertyBits=   671107097
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblTrademarkKind 
         Caption         =   "lbl"
         Height          =   255
         Left            =   2160
         TabIndex        =   97
         Top             =   2505
         Width           =   2415
      End
      Begin VB.Label Label26 
         Caption         =   "申請地址5(中)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   96
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "申請地址4(中)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   95
         Top             =   3060
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "申請地址3(中)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   94
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "申請地址2(中)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   93
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "案件名稱(日)："
         Height          =   255
         Left            =   390
         TabIndex        =   92
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "案件名稱(英)："
         Height          =   255
         Left            =   390
         TabIndex        =   91
         Top             =   3060
         Width           =   1215
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   1
         Left            =   6660
         TabIndex        =   88
         Top             =   660
         Width           =   1995
         VariousPropertyBits=   27
         Size            =   "3519;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPetition 
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   87
         Top             =   720
         Width           =   1995
         VariousPropertyBits=   27
         Size            =   "3519;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   3
         Left            =   6660
         TabIndex        =   84
         Top             =   960
         Width           =   1935
         VariousPropertyBits=   27
         Size            =   "3413;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   83
         Top             =   1020
         Width           =   1995
         VariousPropertyBits=   27
         Size            =   "3519;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPetition 
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   82
         Top             =   1260
         Width           =   2055
         VariousPropertyBits=   27
         Size            =   "3625;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "申請地址1(日)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   81
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "申請地址1(英)："
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   80
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "申請地址2(日)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   79
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "申請地址2(英)："
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   78
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "申請地址3(日)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   77
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "申請地址3(英)："
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   76
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "申請地址4(日)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   75
         Top             =   3660
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "申請地址4(英)："
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   74
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "申請地址5(日)："
         Height          =   255
         Left            =   -74640
         TabIndex        =   73
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "申請地址5(英)："
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   72
         Top             =   4260
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "代表人2(中)："
         Height          =   255
         Left            =   -74610
         TabIndex        =   71
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "代表人1(日)："
         Height          =   255
         Left            =   -74610
         TabIndex        =   70
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "代表人1(英)："
         Height          =   255
         Left            =   -74610
         TabIndex        =   69
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "代表人2(日)："
         Height          =   255
         Left            =   -74610
         TabIndex        =   68
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "代表人2(英)："
         Height          =   255
         Left            =   -74610
         TabIndex        =   67
         Top             =   3060
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "申請人中譯文2："
         Height          =   255
         Left            =   -74610
         TabIndex        =   66
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "申請人中譯文3："
         Height          =   255
         Left            =   -74610
         TabIndex        =   65
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "申請人中譯文4："
         Height          =   255
         Left            =   -74610
         TabIndex        =   64
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "申請人中譯文5："
         Height          =   255
         Left            =   -74610
         TabIndex        =   63
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "申請人3："
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   90
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "申請人5："
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   89
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "申請人4："
         Height          =   255
         Left            =   4560
         TabIndex        =   86
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "申請人2："
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   85
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "代表人中譯文2："
         Height          =   255
         Left            =   -74610
         TabIndex        =   98
         Top             =   3960
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frm880007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/5 改成Form2.0 (chkChoose,txtCaseField)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'strChange存變更事項   bolIsChange是否有輸入變更事項   intFunction 0:其他  1:內商著作權   2:專利
Public strChange As String, bolIsChange As Boolean, intFunction As Integer
Public strCP09 As String

Private Sub chkChoose_Click(Index As Integer)

Dim i As Integer

Select Case Index
             Case 0
                        txtCaseField(0).Enabled = chkChoose(Index).Value
                        If txtCaseField(0).Enabled Then
                           txtCaseField(0).SetFocus
                        End If
             Case 1
                        For i = 1 To intFunction
                               txtCaseField(i).Enabled = chkChoose(Index).Value
                        Next
                        If txtCaseField(1).Enabled Then
                           txtCaseField(1).SetFocus
                        End If
             Case 5, 6
                        txtCaseField(Index + 1).Enabled = chkChoose(Index).Value
                        If txtCaseField(Index + 1).Enabled Then
                           txtCaseField(Index + 1).SetFocus
                        End If
             Case 8
                        For i = 8 To 10
                                  txtCaseField(i).Enabled = chkChoose(Index).Value
                        Next
                        If txtCaseField(8).Enabled Then
                           txtCaseField(8).SetFocus
                        End If
             Case 9, 10, 11
                        txtCaseField(Index + 2).Enabled = chkChoose(Index).Value
                        If txtCaseField(Index + 2).Enabled Then
                           txtCaseField(Index + 2).SetFocus
                        End If
             Case 12
                        For i = 14 To 13 + intFunction
                               txtCaseField(i).Enabled = chkChoose(Index).Value
                        Next
                        If txtCaseField(14).Enabled Then
                           txtCaseField(14).SetFocus
                        End If
             Case 13
                        For i = 19 To 24
                                  txtCaseField(i).Enabled = chkChoose(Index).Value
                        Next
                        If txtCaseField(19).Enabled Then
                           txtCaseField(19).SetFocus
                        End If
             Case 14
                        txtCaseField(Index + 11).Enabled = chkChoose(Index).Value
                        If txtCaseField(Index + 11).Enabled Then
                           txtCaseField(Index + 11).SetFocus
                        End If
             Case 15
                        For i = 26 To 25 + intFunction * 3
                               txtCaseField(i).Enabled = chkChoose(Index).Value
                        Next
                        If txtCaseField(26).Enabled Then
                           txtCaseField(26).SetFocus
                        End If
             Case 16
                        For i = 41 To 42
                               txtCaseField(i).Enabled = chkChoose(Index).Value
                        Next
                        If txtCaseField(41).Enabled Then
                           txtCaseField(41).SetFocus
                        End If
End Select
End Sub
Private Function CheckKeyInOkay(ByRef intIndex As Integer) As Boolean
Dim i As Integer
CheckKeyInOkay = True
Select Case intIndex
             Case 0, 1, 8
                        If txtCaseField(intIndex).Text = "" Then
                           CheckKeyInOkay = False
                           txtCaseField(intIndex).SetFocus
                        Else
                           If intIndex = 1 Then
                              For i = 1 To 5
                                 If CheckKeyIn(i) = -1 Then
                                    CheckKeyInOkay = False
                                    TextInverse txtCaseField(i)
                                    Exit Function
                                 End If
                              Next
                           End If
                        End If
                        
             Case 5, 6
                        If txtCaseField(intIndex + 1).Text = "" Then
                           CheckKeyInOkay = False
                           txtCaseField(intIndex + 1).SetFocus
                        End If
             Case 9, 10, 11, 12
                        If txtCaseField(intIndex + 2).Text = "" Then
                           CheckKeyInOkay = False
                           txtCaseField(intIndex + 2).SetFocus
                        End If
             Case 13
                        If txtCaseField(intIndex + 6).Text = "" Then
                           CheckKeyInOkay = False
                           txtCaseField(intIndex + 6).SetFocus
                        End If
             Case 14, 15
                        If txtCaseField(intIndex + 11).Text = "" Then
                           CheckKeyInOkay = False
                           txtCaseField(intIndex + 11).SetFocus
                        End If
             Case 16
                        If txtCaseField(intIndex + 25).Text = "" Then
                           CheckKeyInOkay = False
                           txtCaseField(intIndex + 25).SetFocus
                        End If
End Select
If CheckKeyInOkay = False Then
   ShowMsg MsgText(9203)
End If
End Function
Private Sub KeyInOkay()
Dim ce(2 To 65) As String, i As Integer, j As Integer
       
For i = 0 To 16
       If chkChoose(i).Value Then
          Select Case i
                       Case 0
                                 ce(2) = txtCaseField(i)
                       Case 1
                                  For j = 0 To 4
                                         ce(4 + j) = txtCaseField(1 + j)
                                  Next
                       Case 2
                                 ce(51) = "1"
                       Case 3
                                 ce(53) = "1"
                       Case 4
                                 ce(55) = "1"
                       Case 5
                                 ce(57) = txtCaseField(6)
                       Case 6
                                 ce(39) = txtCaseField(7)
                       Case 7
                                 ce(59) = "1"
                       Case 8
                                  For j = 0 To 2
                                         ce(41 + j) = txtCaseField(8 + j)
                                  Next
                       Case 9
                                 ce(45) = txtCaseField(11)
                       Case 10
                                 ce(47) = txtCaseField(12)
                       Case 11
                                 ce(49) = txtCaseField(13)
                       Case 12
                                  For j = 0 To 4
                                         ce(17 + j) = txtCaseField(14 + j)
                                  Next
                       Case 13
                                  For j = 0 To 5
                                         ce(10 + j) = txtCaseField(19 + j)
                                  Next
                       Case 14
                                 ce(61) = txtCaseField(25)
                       Case 15
                                  For j = 0 To 14
                                         ce(23 + j) = txtCaseField(26 + j)
                                  Next
                       Case 16
                                  For j = 0 To 1
                                         ce(63 + j) = txtCaseField(41 + j)
                                  Next
           End Select
       End If
Next
strChange = ""
For i = 2 To 65
       'Modify by Morgan 2011/6/10 變更事項改用 chr(29) 分隔,因為地址欄內會有逗號(Ex.CFP-023675)
       'strChange = strChange + "," + ce(i)
       strChange = strChange + Chr(29) + ce(i)
Next
End Sub
Private Sub Analysis()
Dim varTemp As Variant, i As Integer, j As Integer

'Modify by Morgan 2011/6/10 變更事項改用 chr(29) 分隔,因為地址欄內會有逗號(Ex.CFP-023675)
'varTemp = Split(strChange, ",")
varTemp = Split(strChange, Chr(29))

If varTemp(1) <> "" Then
   txtCaseField(0) = varTemp(1)
   chkChoose(0).Value = 1
End If

For j = 0 To 4
       If varTemp(3 + j) <> "" Then
          txtCaseField(1 + j) = varTemp(3 + j)
          CheckKeyIn j + 1
          chkChoose(1).Value = 1
       End If
Next
                       
If varTemp(50) = "1" Then
   chkChoose(2).Value = 1
End If
                       
If varTemp(52) = "1" Then
   chkChoose(3).Value = 1
End If
                       
If varTemp(54) = "1" Then
   chkChoose(4).Value = 1
End If
                       
If varTemp(56) <> "" Then
   txtCaseField(6) = varTemp(56)
   chkChoose(5).Value = 1
End If
                       
If varTemp(38) <> "" Then
   txtCaseField(7) = varTemp(38)
   chkChoose(6).Value = 1
End If

If varTemp(58) = "1" Then
   chkChoose(7).Value = 1
End If
                       
For j = 0 To 2
       If varTemp(40 + j) <> "" Then
          txtCaseField(8 + j) = varTemp(40 + j)
          chkChoose(8).Value = 1
       End If
Next
                       
If varTemp(44) <> "" Then
   txtCaseField(11) = varTemp(44)
   chkChoose(9).Value = 1
End If

If varTemp(46) <> "" Then
   txtCaseField(12) = varTemp(46)
   chkChoose(10).Value = 1
End If
                       
If varTemp(48) <> "" Then
   txtCaseField(13) = varTemp(48)
   chkChoose(11).Value = 1
End If
                       
For j = 0 To 4
       If varTemp(16 + j) <> "" Then
          txtCaseField(14 + j) = varTemp(16 + j)
          chkChoose(12).Value = 1
       End If
Next
                       
For j = 0 To 5
       If varTemp(9 + j) <> "" Then
          txtCaseField(19 + j) = varTemp(9 + j)
          chkChoose(13).Value = 1
       End If
Next
                       
If varTemp(60) <> "" Then
   txtCaseField(25) = varTemp(60)
   chkChoose(14).Value = 1
End If
                       
For j = 0 To 14
       If varTemp(22 + j) <> "" Then
          txtCaseField(26 + j) = varTemp(22 + j)
          chkChoose(15).Value = 1
       End If
Next

For j = 0 To 1
       If varTemp(62 + j) <> "" Then
          txtCaseField(41 + j) = varTemp(62 + j)
          chkChoose(16).Value = 1
       End If
Next
End Sub
Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer

bolIsChange = False

   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Sub
   End If
   

If Index = 0 Then
   For i = 0 To 15
          If chkChoose(i).Value Then
             bolIsChange = True
             If CheckKeyInOkay(i) = False Then
                Exit For
             End If
          End If
   Next
   If i = 16 Then
      'Add by Morgan 2011/8/2 只有CFP發文用到
      If chkChoose(1).Value = 1 Then
         If field(1) = "CFP" Then
            '若有勾申請人時需檢查輸入的申請人編號不可與原來的完全相同(九碼)
            If ChangeCustomerL(txtCaseField(1)) = ChangeCustomerL(field(26)) And _
                ChangeCustomerL(txtCaseField(2)) = ChangeCustomerL(field(27)) And _
                 ChangeCustomerL(txtCaseField(3)) = ChangeCustomerL(field(28)) And _
                  ChangeCustomerL(txtCaseField(4)) = ChangeCustomerL(field(29)) And _
                   ChangeCustomerL(txtCaseField(5)) = ChangeCustomerL(field(30)) Then
               MsgBox "新申請人編號與目前相同 !", vbCritical
               Exit Sub
            End If
         End If
      End If
      'end 2011/8/2
      KeyInOkay
      Unload Me
   End If
Else
   Unload Me
End If
End Sub
Private Sub Form_Activate()
   ReadAllData
   If strChange <> "" Then Analysis
End Sub
Private Sub ReadAllData()
   Dim rt As Boolean, i As Integer, strTemp As String, strTemp1 As String, j As Integer

On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   ReDim cp(TF_CP) As String
   'Modify by Morgan 2009/7/24 發文畫面不再共用,收文號改用傳的
   'cp(9) = frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.row, 5)
   cp(9) = strCP09
   'end 2009/7/24
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
      txtCaseField(8) = field(5)
      txtCaseField(9) = field(6)
      txtCaseField(10) = field(7)
            
      Select Case intCaseKind
         Case 專利
            
            txtCaseField(0) = field(10)
            txtCaseField(1) = field(26)
            txtCaseField(2) = field(27)
            txtCaseField(3) = field(28)
            txtCaseField(4) = field(29)
            txtCaseField(5) = field(30)
            txtCaseField(7) = field(8)
            txtCaseField(19) = field(79)
            txtCaseField(20) = field(80)
            txtCaseField(21) = field(81)
            txtCaseField(22) = field(82)
            txtCaseField(23) = field(83)
            txtCaseField(24) = field(84)
            
         Case 商標
            txtCaseField(0) = field(11)
            txtCaseField(1) = field(23)
            txtCaseField(6) = field(27)
            txtCaseField(7) = field(8)
            txtCaseField(12) = field(9)
            txtCaseField(13) = field(32)
            txtCaseField(19) = field(47)
            txtCaseField(20) = field(48)
            txtCaseField(21) = field(49)
            txtCaseField(22) = field(50)
            txtCaseField(23) = field(51)
            txtCaseField(24) = field(52)
            
         Case Else
            txtCaseField(0) = field(10)
            txtCaseField(1) = field(8)
            txtCaseField(2) = field(58)
            txtCaseField(3) = field(59)
      End Select
   
      For i = 0 To intFunction
         CheckKeyIn i + 1
      Next
      CheckKeyIn 7
   Else
      Unload Me
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   ErrorMsg
End Sub
Private Sub Form_Load()
'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
txtCaseField(19).MaxLength = Pub_MaxCEL10
txtCaseField(20).MaxLength = Pub_MaxCEL11
txtCaseField(22).MaxLength = Pub_MaxCEL10
txtCaseField(23).MaxLength = Pub_MaxCEL11
'end 2016/09/10

MoveFormToCenter Me
'intFunction 0:其他  1:內商著作權   2:專利
'並轉為 1,3,5
If intFunction = 2 Then
   intFunction = 5
   chkChoose(5).Enabled = False
   chkChoose(7).Enabled = False
   chkChoose(9).Enabled = False
   chkChoose(10).Enabled = False
   chkChoose(11).Enabled = False
ElseIf intFunction = 1 Then
   intFunction = 3
   chkChoose(5).Enabled = False
   chkChoose(6).Enabled = False
   chkChoose(7).Enabled = False
   chkChoose(9).Enabled = False
   chkChoose(10).Enabled = False
   chkChoose(11).Enabled = False
Else
   intFunction = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
'   Set frm880007 = Nothing
End Sub

Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 1, 2, 3, 4, 5
                        lblPetition(Index - 1).Caption = ""
                        txtCaseField(13 + Index) = ""
                        txtCaseField((Index - 1) * 3 + 26) = ""
                        txtCaseField((Index - 1) * 3 + 27) = ""
                        txtCaseField((Index - 1) * 3 + 28) = ""
             Case 7
                        lblTrademarkKind = ""
End Select
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 2, 3, 4, 5
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
End If
If Cancel Then TextInverse txtCaseField(Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strTemp2 As String, strTemp3 As String, strCusTemp As String

CheckKeyIn = -1
Select Case intIndex
             Case 0
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsDate(txtCaseField(intIndex).Text) Then
                              CheckKeyIn = 1
                           End If
                        End If
             Case 1, 2, 3, 4, 5
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           strCusTemp = txtCaseField(intIndex)
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCustomer(strCusTemp, strTemp) Then
                           If ClsPDGetCustomer(strCusTemp, strTemp) Then
                              txtCaseField(intIndex) = strCusTemp
                              lblPetition(intIndex - 1).Caption = strTemp
                              'edit by nickc 2007/02/02 不用 dll 了
                              'If objPublicData.GetCustomerNameAndAddress(strCusTemp, strTemp, strTemp1, strTemp2, strTemp3) Then
                              If ClsPDGetCustomerNameAndAddress(strCusTemp, strTemp, strTemp1, strTemp2, strTemp3) Then
                                 txtCaseField(13 + intIndex) = strTemp
                                 txtCaseField((intIndex - 1) * 3 + 26) = strTemp1
                                 txtCaseField((intIndex - 1) * 3 + 27) = strTemp2
                                 txtCaseField((intIndex - 1) * 3 + 28) = strTemp3
                              End If
                              CheckKeyIn = 1
                           End If
                        End If
             Case 7
                        lblTrademarkKind = ""
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetPatentTrademarkKind(專利, txtCaseField(intIndex).Text, strTemp, False, 台灣國家代號) Then
                           If ClsPDGetPatentTrademarkKind(專利, txtCaseField(intIndex).Text, strTemp, False, 台灣國家代號) Then
                              lblTrademarkKind = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
End Sub
