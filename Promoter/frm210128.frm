VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210128 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內潛在客戶資料維護"
   ClientHeight    =   5736
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.ComboBox cboStatus 
      Height          =   300
      ItemData        =   "frm210128.frx":0000
      Left            =   6705
      List            =   "frm210128.frx":000A
      TabIndex        =   13
      Text            =   "cboStatus"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtSameCnt 
      Height          =   270
      Left            =   -960
      MaxLength       =   6
      TabIndex        =   39
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1260
      TabIndex        =   8
      Top             =   2700
      Width           =   1875
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7710
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
            Picture         =   "frm210128.frx":002A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":0346
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":0662
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":083E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":0B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":0E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":1192
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":14AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":17CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":1AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210128.frx":1E02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox textCUID 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   690
      Width           =   5670
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8952
      _ExtentX        =   15790
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
   Begin MSForms.Label lblPOC16 
      Height          =   255
      Left            =   2340
      TabIndex        =   50
      Top             =   3330
      Width           =   3000
      BackColor       =   16777215
      VariousPropertyBits=   27
      Size            =   "5292;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC16N 
      Height          =   288
      Left            =   6384
      TabIndex        =   15
      Top             =   3264
      Visible         =   0   'False
      Width           =   3000
      VariousPropertyBits=   679493659
      MaxLength       =   79
      Size            =   "5292;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "為關係企業"
      Height          =   180
      Index           =   8
      Left            =   5400
      TabIndex        =   49
      Top             =   3360
      Width           =   900
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   264
      Index           =   17
      Left            =   390
      TabIndex        =   47
      Top             =   4770
      Visible         =   0   'False
      Width           =   825
      VariousPropertyBits=   679493659
      MaxLength       =   12
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   28
      Left            =   7665
      TabIndex        =   10
      Top             =   2700
      Width           =   330
      VariousPropertyBits=   679493659
      MaxLength       =   1
      Size            =   "582;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   27
      Left            =   1260
      TabIndex        =   7
      Top             =   2385
      Width           =   7335
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12938;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   24
      Left            =   1260
      TabIndex        =   4
      Top             =   1530
      Width           =   5415
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "9551;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   25
      Left            =   1260
      TabIndex        =   5
      Top             =   1800
      Width           =   5415
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "9551;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   26
      Left            =   1260
      TabIndex        =   6
      Top             =   2100
      Width           =   5415
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "9551;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   23
      Left            =   1260
      TabIndex        =   3
      Top             =   1260
      Width           =   5415
      VariousPropertyBits=   679493659
      MaxLength       =   30
      Size            =   "9551;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   12
      Left            =   1260
      TabIndex        =   11
      Top             =   3000
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   7
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   13
      Left            =   3255
      TabIndex        =   12
      Top             =   3000
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   6
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   264
      Index           =   14
      Left            =   -690
      TabIndex        =   38
      Top             =   4770
      Visible         =   0   'False
      Width           =   1035
      VariousPropertyBits=   679493659
      MaxLength       =   12
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   16
      Left            =   1260
      TabIndex        =   14
      Top             =   3300
      Width           =   1035
      VariousPropertyBits=   679493659
      MaxLength       =   9
      Size            =   "1826;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   585
      Index           =   10
      Left            =   1260
      TabIndex        =   21
      Top             =   4500
      Width           =   7665
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "13520;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   5
      Left            =   1260
      TabIndex        =   16
      Top             =   3600
      Width           =   3000
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "5292;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   6
      Left            =   5235
      TabIndex        =   17
      Top             =   3600
      Width           =   3000
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "5292;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   7
      Left            =   1260
      TabIndex        =   18
      Top             =   3900
      Width           =   3000
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "5292;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   9
      Left            =   5235
      TabIndex        =   20
      Top             =   4200
      Width           =   3675
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "6482;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   8
      Left            =   1260
      TabIndex        =   19
      Top             =   4200
      Width           =   3000
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "5292;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   11
      Left            =   4575
      TabIndex        =   9
      Top             =   2700
      Width           =   330
      VariousPropertyBits=   679493659
      MaxLength       =   1
      Size            =   "582;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   585
      Index           =   15
      Left            =   1260
      TabIndex        =   22
      Top             =   5100
      Width           =   7665
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13520;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   1
      Top             =   675
      Width           =   255
      VariousPropertyBits=   679493659
      MaxLength       =   1
      Size            =   "450;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   0
      Top             =   675
      Width           =   1095
      VariousPropertyBits=   679493659
      MaxLength       =   8
      Size            =   "1926;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   264
      Index           =   4
      Left            =   -300
      TabIndex        =   37
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   285
      Index           =   3
      Left            =   1260
      TabIndex        =   2
      Top             =   990
      Width           =   7335
      VariousPropertyBits=   679493659
      MaxLength       =   79
      Size            =   "12938;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "ＰＳ：非研發處或專利處程序之潛在客戶, 三年內無往來記錄者, 系統會自動刪除．"
      ForeColor       =   &H000000FF&
      Height          =   720
      Index           =   7
      Left            =   6840
      TabIndex        =   48
      Top             =   1440
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否寄發專利雙週報：      （N:不寄）"
      Height          =   180
      Index           =   4
      Left            =   5850
      TabIndex        =   46
      Top             =   2760
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Index           =   19
      Left            =   150
      TabIndex        =   26
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名稱（英）："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   44
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名稱（日）："
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   43
      Top             =   2460
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   3
      Left            =   2340
      TabIndex        =   42
      Top             =   3030
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "狀　　態："
      Height          =   180
      Index           =   21
      Left            =   5760
      TabIndex        =   41
      Top             =   3030
      Width           =   915
   End
   Begin MSForms.Label lbl1 
      Height          =   240
      Index           =   2
      Left            =   4170
      TabIndex        =   40
      Top             =   3030
      Width           =   1185
      VariousPropertyBits=   27
      Caption         =   "lbl1"
      Size            =   "2090;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國　　籍："
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   28
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "與"
      Height          =   180
      Index           =   107
      Left            =   996
      TabIndex        =   36
      Top             =   3360
      Width           =   180
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "地　　址："
      Height          =   180
      Index           =   28
      Left            =   150
      TabIndex        =   35
      Top             =   4560
      Width           =   1080
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "電　話1："
      Height          =   180
      Index           =   9
      Left            =   150
      TabIndex        =   34
      Top             =   3660
      Width           =   1080
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "傳　真1："
      Height          =   180
      Index           =   10
      Left            =   150
      TabIndex        =   33
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "E-MAIL："
      Height          =   180
      Index           =   11
      Left            =   4305
      TabIndex        =   32
      Top             =   4230
      Width           =   900
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "電　話2："
      Height          =   180
      Index           =   12
      Left            =   4305
      TabIndex        =   31
      Top             =   3660
      Width           =   900
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "行動電話："
      Height          =   180
      Index           =   15
      Left            =   150
      TabIndex        =   30
      Top             =   4230
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否寄電子報：      （N:不寄）"
      Height          =   180
      Index           =   17
      Left            =   3315
      TabIndex        =   29
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "開發日期："
      Height          =   180
      Index           =   13
      Left            =   150
      TabIndex        =   27
      Top             =   3030
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "編　　號："
      Height          =   210
      Index           =   0
      Left            =   150
      TabIndex        =   24
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名稱（中）："
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   45
      Top             =   1020
      Width           =   1080
   End
End
Attribute VB_Name = "frm210128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/06/15 杜燕文協理要求改成用Label顯示關聯企業名稱，原TextBox隱藏
'Memo by Amy 2021/12/14 Form2.0已修改 lbl1(2)/txtPoc();2023/06/14 更換txtPOC16N
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
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

Dim TF_POC As Integer
Dim strTmp As String
Dim oText 'Modfy by Amy 2021/12/14 原:As TextBox
Dim oLabel 'Modfy by Amy 2021/12/14 原:As LABEL
Dim idx As Integer
Dim rsContact As ADODB.Recordset
Dim rsContactOld As ADODB.Recordset
Dim rsContactSim As ADODB.Recordset
Dim m_bSaveCheck As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim oldPOC11 As String, oldPOC28 As String 'Add by Amy 2014/03/07 記錄 是否寄電子報值及專利雙週報值

'Modify by Amy 2021/12/14 從Form_KeyDown搬過來
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim bCancel As Boolean 'Add by Amy 2017/09/01
    
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
         'Add by Amy 2017/09/01
         If m_EditMode = 1 Or m_EditMode = 2 Then
            Call txtPOC_Validate(3, bCancel)
            If bCancel = True Then KeyCode = 0
         End If
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
         'Add by Amy 2015/01/15 +備註可輸入換行
         If KeyCode = vbKeyReturn And UCase(Me.ActiveControl.Name) = UCase("txtPOC") Then
            If Me.ActiveControl.Index = 15 Then Exit Sub
         End If
         'end 2015/01/15
         KeyCode = 0
         'Mark by Amy 2021/12/14 不使用Enter執行確定,否則有錯誤訊息按Enter會再彈一次
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyInsert
   End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   '取得使用者執行各項功能的權限
   m_bInsert = True 'IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = True 'IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = True 'IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   '狀態的下拉選單
   Me.cboStatus.Clear
   Me.cboStatus.AddItem ""
   Me.cboStatus.AddItem "刪址"
   Me.cboStatus.AddItem "倒閉"
   Me.cboStatus.AddItem "遷移不明"
   Me.cboStatus.AddItem "解散"
   Me.cboStatus.AddItem "廢止"
   Me.cboStatus.AddItem "撤銷"
   Me.cboStatus.AddItem "停業"
   Me.cboStatus.AddItem "往生"
   Me.cboStatus.AddItem "業務自行處理"
   Me.cboStatus.AddItem "國內同業" 'Add by Amy 2021/11/29
   'Add by Amy 2015/03/23 電腦中心及研發處 +"開拓不寄"
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "D01" Then
        Me.cboStatus.AddItem "開拓不寄"
        'Add by Amy 2021/11/29
        If Pub_StrUserSt03 = "M51" Then Me.cboStatus.AddItem "設為對造"
   End If
   txtPOC(14).Visible = False
   
   '國籍的下拉選單
   Me.Combo1.Clear
   'Modify By Sindy 2012/10/31 可以輸入1011(美國)
   'StrSQLa = "Select NA01, NA03 From Nation Where Length(NA01)=3 And NA01<='9999' And NA01<>'000' Order By NA01 "
   StrSQLa = "Select NA01, NA03 From Nation Where NA01<='9999' And NA01<>'000' Order By NA01 "
   '2012/10/31 End
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   While Not rsA.EOF
      Me.Combo1.AddItem "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
      rsA.MoveNext
   Wend
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   txtPOC(4).Visible = False
   
   txtSameCnt.Visible = False
   
   textCUID.BackColor = &H8000000F
   InitialField
   ShowRecord 99
   m_EditMode = 0
   SetInputEntry
   UpdateToolbarState
End Sub

Private Sub Form_Initialize()
   strExc(0) = "select * from PotCustomer1 where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_POC = RsTemp.Fields.Count
   
   ReDim m_FieldList(TF_POC) As FIELDITEM
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210128 = Nothing
End Sub

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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   
   strExc(0) = "SELECT * FROM PotCustomer1 " & _
            "WHERE POC01 = '" & strKEY01 & "' AND POC02 = '" & strKEY02 & "' "
                  
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      IsRecordExist = True
   End If
End Function

Private Sub txtPOC_Change(Index As Integer)
Dim strTempName As String
   
   Select Case Index
      Case 4 '國籍
'         If Left(txtPOC(Index).Tag, 3) <> Left(txtPOC(Index), 3) Then
'            lbl1(1).Caption = ""
'            If Len(txtPOC(Index)) >= 3 Then
'               If ClsPDGetNation(Left(txtPOC(Index), 3), strTmp) = True Then
'                  lbl1(1).Caption = strTmp
'               End If
'            End If
'         End If
'         txtPOC(Index).Tag = txtPOC(Index)
      Case 13 '智權人員
         'Mark by Amy 2016/09/05
'        If Len(txtPOC(Index)) = 5 Then
'            If ClsPDGetStaff(txtPOC(Index), strTempName) = True Then
'               LBL1(2).Caption = strTempName
'            End If
'         Else
'            LBL1(2).Caption = ""
'         End If
   End Select
End Sub

Public Sub txtPOC_GotFocus(Index As Integer)
   Select Case Index
      Case 3, 10, 15, 27
         OpenIme
      Case Else
         CloseIme
   End Select
   
   '國籍第4碼檢查錯誤時
   If m_bSaveCheck = True Then
      txtPOC(Index).SelStart = 3
      txtPOC(Index).SelLength = 1
      m_bSaveCheck = False
      Exit Sub
   End If
   
   TextInverse txtPOC(Index)
End Sub

'Modify by Amy 2021/12/14 原:Integer
Private Sub txtPOC_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 2, 13, 16
         KeyAscii = UpperCase(KeyAscii)
      'Added by Lydia 2019/08/19 Email輸入字元檢查(與客戶檔frm140401一致)
      Case 9
          PUB_EMailFilter (KeyAscii)
      'end 2019/08/19
      Case 10
         'Modify by Amy 2021/12/14 +txtPOC(Index)
         KeyAscii = ChangeZIP(KeyAscii, txtPOC(Index))
      Case 11
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      'Add By Sindy 2011/1/14
      Case 28
         KeyAscii = UpperCase(KeyAscii)
         'Modified by Morgan 2012/1/2 改放N
         'If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Public Sub txtPOC_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String
Dim iLen As Integer
   
   If txtPOC(Index).Locked = True Then Exit Sub
   Select Case Index
      Case 1 '編號
         If Not IsEmptyText(txtPOC(1)) Then
            If Mid(txtPOC(1), 1, 1) <> "R" Then
               Cancel = True
               MsgBox "客戶編號必須為R開頭", vbCritical + vbOKOnly, "檢核資料"
               txtPOC(1).Text = ""
               txtPOC_GotFocus Index
               Exit Sub
            End If
            
            If Len(txtPOC(1)) < 6 Then
               Cancel = True
               MsgBox "客戶編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
               txtPOC_GotFocus Index
               Exit Sub
            End If
            
            txtPOC(1) = Left(txtPOC(1) & "00", 8)
            txtPOC(2) = Left(txtPOC(2) & "0", 1)
            If m_EditMode = 1 Then
               If IsRecordExist(txtPOC(1), txtPOC(2)) = True Then
                  Cancel = True
                  MsgBox "該筆客戶已存在! ", vbCritical + vbOKOnly, "檢核資料"
                  txtPOC_GotFocus Index
                  Exit Sub
               End If
               If IsOverAutoNumber("R", Empty, Mid(txtPOC(1), 2, 5)) = True Then
                  Cancel = True
                  MsgBox "客戶代碼超過自動編號! ", vbCritical + vbOKOnly, "檢核資料"
                  txtPOC_GotFocus Index
                  Exit Sub
               End If
            End If
          End If
           
      Case 3, 23, 24, 25, 26, 27 '名稱
         If m_EditMode = 1 Or m_EditMode = 2 Then
            If txtPOC(Index) <> "" Then
                'Add by Amy 2017/09/01 (股)公司改為股份有限公司,?判斷
                If Index = 3 Then
                    If InStr(txtPOC(Index), "(股)公司") > 0 Or InStr(txtPOC(Index), "（股）公司") > 0 _
                      Or InStr(txtPOC(Index), "(股)有限公司") > 0 Or InStr(txtPOC(Index), "（股）有限公司") Then
                        MsgBox Left(Label1(6), Len(Label1(6)) - 1) & "(股)公司，請改為全名「股份有限公司」！", vbExclamation
                        Cancel = True
                        txtPOC(Index).SetFocus
                        txtPOC_GotFocus (Index)
                    End If
                    If InStr(txtPOC(Index), "?") > 0 Then
                        MsgBox Left(Label1(6), Len(Label1(6)) - 1) & " 有「?」請確認！", vbExclamation
                        Cancel = True
                        txtPOC(Index).SetFocus
                        txtPOC_GotFocus (Index)
                    End If
                End If
'               If txtSameCnt = "Y" Then
'                  Me.Show
'                  Me.Combo1.SetFocus
'                  txtSameCnt = ""
'               ElseIf txtSameCnt = "N" Then
'                  Me.Show
'                  Me.txtPOC(Index).SetFocus
'                  Me.txtPOC_GotFocus Index
'                  Cancel = True
'                  txtSameCnt = ""
'                  Exit Sub
'               Else
'                  Me.Enabled = False
'                  Screen.MousePointer = vbHourglass
'                  frm210128_1.G_strText = Trim(txtPOC(Index).Text)
'                  frm210128_1.G_intIndex = Index
'                  frm210128_1.Show
'                  frm210128_1.StrMenu
'                  Screen.MousePointer = vbDefault
'                  Me.Enabled = True
'                  Me.Hide
'                  If txtSameCnt.Text = "" Or txtSameCnt.Text = 0 Then
'                     txtSameCnt.Text = ""
'                     Me.Show
'                  End If
'               End If
               'Modify By Sindy 2014/2/27
               txtSameCnt.Text = ""
               frm210128_1.G_strText = Trim(txtPOC(Index).Text)
               frm210128_1.G_intIndex = Index
               strCustNo = ""
               If txtPOC(1) <> "" Then
                  strCustNo = Left(Trim(txtPOC(1)) & "00000", 8) & Left(Trim(txtPOC(2)) & "0", 1)
               End If
               If frm210128_1.StrMenu(strCustNo) = True Then
                  Me.Hide
                  frm210128_1.Show vbModal
                  If txtSameCnt = "Y" Then
                     Me.Show
                     Me.Combo1.SetFocus
                     txtSameCnt = ""
                  ElseIf txtSameCnt = "N" Then
                     Me.Show
                     Me.txtPOC(Index).SetFocus
                     Me.txtPOC_GotFocus Index
                     Cancel = True
                     txtSameCnt = ""
                     Exit Sub
                  End If
               Else
                  Unload frm210128_1
                  Me.Show
               End If
               '2014/2/27 END
            End If
         End If
         
      Case 4 '國籍
         If txtPOC(4) <> "" Then
            If txtPOC(4) = 台灣國家代號 Then
               Cancel = True
               ShowMsg MsgText(9153)
'            Else
'               If lbl1(1).Caption = "" Then
'                  Cancel = True
'               End If
            End If
         End If
         
      '2012/12/3 ADD BY SONIA
      'Add by Amy 2017/09/01 ?判斷
      Case 9 'mail
        If m_EditMode = 1 Or m_EditMode = 2 Then
            'Modified by Lydia 2019/08/19 Email輸入字元檢查(與客戶檔frm140401一致)
            'If InStr(txtPOC(Index), "?") > 0 Then
            '    MsgBox Left(Label63(11), Len(Label63(11)) - 1) & " 有「?」請確認！", vbExclamation
            If PUB_CheckMail(txtPOC(Index)) = False Then
            'end 2019/08/19
               Cancel = True
               txtPOC(Index).SetFocus
               txtPOC_GotFocus (Index)
            End If
        End If
      Case 10 '地址
         'Add by Amy 2017/09/01 ?判斷
         If m_EditMode = 1 Or m_EditMode = 2 Then
            If InStr(txtPOC(Index), "?") > 0 Then
               MsgBox Left(Label41(28), Len(Label41(28)) - 1) & " 有「?」請確認！", vbExclamation
               Cancel = True
               txtPOC(Index).SetFocus
               txtPOC_GotFocus (Index)
            End If
         End If
         If txtPOC(10) <> "" Then
            If CheckTaiwanAddr(txtPOC(10), Combo1, Label41(28)) = False Then
               txtPOC(Index).SetFocus
               txtPOC_GotFocus Index
               Cancel = True
            End If
         End If
      '2012/12/3 END
      
      Case 12 '開發日期
         If txtPOC(Index) <> "" Then
            If ChkDate(txtPOC(Index)) = False Then
               txtPOC(Index).SetFocus
               txtPOC_GotFocus Index
               Cancel = True
            End If
         End If
         
      Case 13 '智權人員
         If txtPOC(Index).Visible = True Then
            'If txtPOC(Index) <> "" And LBL1(2) = "" Then
               If Len(txtPOC(Index)) = 5 Then
                  'Modify by Amy 2016/09/05
                  LBL1(2).Caption = ""
                  If m_EditMode = 1 Then
                    If ClsPDGetStaff(txtPOC(Index), strTempName) = True Then
                       LBL1(2) = strTempName
                    Else
                        Cancel = True
                    End If
'                    If LBL1(2) = "" Then
'                       MsgBox "員工編號輸入錯誤！", vbExclamation
'                       Cancel = True
'                    End If
                  Else
                        strExc(0) = PUB_GetStaffNameDept(txtPOC(Index), strTempName, strExc(0), True, IIf(m_EditMode = 2, True, False))
                        LBL1(2).Caption = strTempName
                  End If
                  'end 2016/09/05
               End If
            'End If
         End If
         
      Case 9 'E-Mail
         If txtPOC(Index) <> "" Then
            If InStr(1, txtPOC(Index), "@") = 0 Then
                MsgBox "Mail 必需要有 @ 符號！"
                Cancel = True
            ElseIf InStr(1, txtPOC(Index), ",") > 0 Or InStr(1, txtPOC(Index), "[") > 0 Or InStr(1, txtPOC(Index), "]") > 0 Or InStr(1, txtPOC(Index), "!") > 0 Or InStr(1, txtPOC(Index), "(") > 1 Or InStr(1, txtPOC(Index), ")") > 0 Or InStr(1, txtPOC(Index), "=") > 0 Or InStr(1, txtPOC(Index), "\") > 0 Or InStr(1, txtPOC(Index), "/") > 0 Or InStr(1, txtPOC(Index), "<") > 0 Or InStr(1, txtPOC(Index), ">") > 0 Or InStr(1, txtPOC(Index), "~") > 0 Or InStr(1, txtPOC(Index), "$") > 0 Or InStr(1, txtPOC(Index), "%") > 0 Or InStr(1, txtPOC(Index), "^") > 0 Or InStr(1, txtPOC(Index), "&") > 0 Or InStr(1, txtPOC(Index), "*") > 0 Then
                MsgBox "Mail 不允許有下列符號！" & vbCrLf & ",、[、]、!、(、)、=、\、/、<、>、~、$、%、^、&、* "
                Cancel = True
            End If
         End If
         
      Case 16 '關係企業
         If txtPOC(Index) <> "" Then
            If Len(txtPOC(Index)) > 5 Then
               txtPOC(Index) = Left(txtPOC(Index) & "000", 9)
               If GetCustData(txtPOC(Index)) = False Then
                  If m_EditMode = "1" Or m_EditMode = "2" Then
                     Cancel = True
                     txtPOC_GotFocus Index
                  End If
               End If
            Else
               Cancel = True
               MsgBox "關係企業編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
               txtPOC_GotFocus Index
            End If
         End If
   End Select
   
   If Cancel = False Then
      '欄位長度檢查
      Select Case Index
         '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
         Case 3, 10, 15, 27
            iLen = txtPOC(Index).MaxLength - 1
         Case Else
            iLen = txtPOC(Index).MaxLength
      End Select
      
      If Not CheckLengthIsOK(txtPOC(Index), iLen) Then
         Cancel = True
      End If
   End If
   CloseIme
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean 'Add by Amy 2017/09/01
   Dim strTp As String 'Add by Amy 2023/07/12
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         '開發日期預設當天
         txtPOC(12) = ChangeWStringToTString(strSrvDate(1))
         
         'Removed by Morgan 2012/1/2 改放N,都預設要寄
         ''Add By Sindy 2011/1/14 由98012陳品薇(內專程序主管)輸入的資料, 預設為要寄專利雙週報
         'If PUB_GetST05(strUserNum) = "73" Then
         '   txtPOC(28) = "Y"
         'End If
         
      Case vbKeyF3 ' 修改
         m_EditMode = 2
         If CheckModifyLimit(txtPOC(13).Text, True, txtPOC(17).Text) = False Then Exit Sub
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyF5 ' 刪除
         If Pub_StrUserSt03 <> "M51" Then
            MsgBox "無刪除權限 !!!", vbInformation
            Exit Sub
         End If
         'Add by Amy 2023/07/12 若存在XYS02介紹來源編號,則不可刪
         'Modify by Amy 2024/11/29 考慮多筆,改訊息至共用
         If txtPOC(2) = "0" And Pub_GetXYSource(2, txtPOC(1), , , , Me.Name, strTp) = True Then
            MsgBox strTp, vbOKOnly, "注意"
            Exit Sub
         End If
         
         'If CheckModifyLimit(txtPOC(13).Text,True) = False Then Exit Sub
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
         'Add By Sindy 2009/04/20
         If m_EditMode = 4 Then
            txtPOC(1).Tag = txtPOC(1)
            txtPOC(2).Tag = txtPOC(2)
         'Add by Amy 2017/09/01 取代(股)
         ElseIf m_EditMode = 1 Or m_EditMode = 2 Then
            Call txtPOC_Validate(3, bCancel)
            If bCancel = True Then Exit Sub
         End If
         'Modify By Amy 2015/01/15 +不過濾的文字框.name
         PUB_FilterFormText Me, "txtPOC(15)"
         
         'If txtSameCnt <> "" Then Exit Sub
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetInputEntry
      Case vbKeyF10 ' 取消
         'If txtSameCnt <> "" Then Exit Sub
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  txtPOC(1) = txtPOC(1).Tag
                  txtPOC(2) = txtPOC(2).Tag
                  m_EditMode = 0
                  SetInputEntry
                  ShowRecord
                  UpdateToolbarState
               End If
            Case Else
               txtPOC(1) = txtPOC(1).Tag
               txtPOC(2) = txtPOC(2).Tag
               m_EditMode = 0
               SetInputEntry
               'Add By Sindy 2023/9/6
               If txtPOC(1).Tag <> "" And txtPOC(2).Tag <> "" Then
               '2023/9/6 END
                  ShowRecord
               End If
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

Private Sub ClearField()
   For Each oText In txtPOC
      oText.Text = Empty
   Next
   For Each oLabel In LBL1
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_POC
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   'Add By Sindy 2009/06/24
   txtPOC16N = ""
   '2009/06/24 End
   'Modify By Sindy 2012/10/4 打字室預設 智權人員=A0038賴晏翎
   'Modify By Sindy 2013/8/30 打字室預設 智權人員=79075郭雅娟
   If Pub_StrUserSt03 = "M13" Then
      txtPOC(13).Text = "79075"
   Else
   '2012/10/4 End
      txtPOC(13).Text = strUserNum
   End If
   textCUID = "": txtPOC(17) = ""
   cboStatus = Empty
   Combo1 = Empty
   lblPOC16 = "" 'Added by Lydia 2023/06/15
   'Add By Sindy 2023/9/6
   txtPOC(1).Tag = ""
   txtPOC(2).Tag = ""
   '2023/9/6 END
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtPOC
      oText.Locked = bLocked
   Next
   cboStatus.Locked = bLocked
   Combo1.Locked = bLocked
   'Add By Sindy 2009/06/24
   txtPOC16N.Enabled = Not bLocked
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
         If m_bUpdate And txtPOC(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtPOC(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         'Modfiy by Amy 2015/06/10 +登入者資料量少按 第/前/後/最後一筆會跑很久,故鎖住
'         If m_bQuery And txtPOC(1) <> "" Then
'            TBar1.Buttons(6).Enabled = True
'            TBar1.Buttons(7).Enabled = True
'            TBar1.Buttons(8).Enabled = True
'            TBar1.Buttons(9).Enabled = True
'         Else
'            TBar1.Buttons(6).Enabled = False
'            TBar1.Buttons(7).Enabled = False
'            TBar1.Buttons(8).Enabled = False
'            TBar1.Buttons(9).Enabled = False
'         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
         
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         'Modfiy by Amy 2015/06/10 +登入者資料量少按 第/前/後/最後一筆會跑很久,故鎖住
'         TBar1.Buttons(6).Enabled = False
'         TBar1.Buttons(7).Enabled = False
'         TBar1.Buttons(8).Enabled = False
'         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
   'Modfiy by Amy 2015/06/10 +登入者資料量少按 第/前/後/最後一筆會跑很久,故鎖住
   TBar1.Buttons(6).Enabled = False
   TBar1.Buttons(7).Enabled = False
   TBar1.Buttons(8).Enabled = False
   TBar1.Buttons(9).Enabled = False
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 1
            txtPOC(1).Locked = False
            txtPOC(2).Locked = False
            txtPOC(3).SetFocus
         Case 2
            txtPOC(1).Locked = True
            txtPOC(2).Locked = True
            Combo1.SetFocus
         Case 4
            txtPOC(1).Locked = False
            txtPOC(2).Locked = False
            txtPOC(1).SetFocus
         Case Else
            txtPOC(1).Locked = True
            txtPOC(2).Locked = True
            txtPOC(1).SetFocus
      End Select
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean, ii As Integer, jj As Integer
   Dim iRtn As Integer 'Add by Amy 2021/11/29

   For Each oText In txtPOC
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         Cancel = False
         If idx <> 3 Then
            txtPOC_Validate idx, Cancel
         End If
         If Cancel = True Then
            txtPOC(idx).SetFocus
            txtPOC_GotFocus idx
            Exit Function
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtPOC(1) = "" Then
         ShowMsg "請輸入欲查詢之客戶編號 !"
         txtPOC(1).SetFocus
         txtPOC_GotFocus 1
         Exit Function
      End If
   '維護
   Else
      'Add by Amy 2021/12/14檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, True, True) = False Then
         Exit Function
      End If
      
      If txtPOC(3) = "" And txtPOC(23) = "" And txtPOC(27) = "" Then
         ShowMsg "客戶中、英、日文名稱不可同時為空白 !"
         txtPOC(3).SetFocus
         Exit Function
      End If
      'Add by Amy 2015/03/23 +避免研發處輸入後直接按Enter存檔,未檢查狀態
      Cancel = False
      cboStatus_Validate Cancel
      If Cancel = True Then
          cboStatus.SetFocus
          Exit Function
      End If
      'end 2015/03/23
      'Add by Amy 2021/11/29 國內同業判斷
      If Left(Trim(Combo1), 3) < "010" And InStr(txtPOC(3), "事務所") > 0 And cboStatus <> "國內同業" Then
        iRtn = MsgBox("國籍在台灣之事務所，請確認是否為國內同業？" & vbCrLf & _
                                "是:為國內同業　否:非國內同業", vbYesNoCancel + vbDefaultButton3)
        '取消
        If iRtn = 2 Then
            Exit Function
        '是
        ElseIf iRtn = 6 Then
            cboStatus = "國內同業"
        End If 'iRtn
      End If
      If cboStatus = "國內同業" Then
        '電子報要設定不寄
        If txtPOC(11) <> "N" Then
            ShowMsg "此為國內同業, 不可寄電子報 ！"
            txtPOC(11).SetFocus
            txtPOC_GotFocus (11)
            Exit Function
        End If
        If txtPOC(28) <> "N" Then
            ShowMsg "此為國內同業, 不可寄專利雙週報 ！"
            txtPOC(28).SetFocus
            txtPOC_GotFocus (28)
            Exit Function
        End If
        '不可寄mail
        If txtPOC(9) <> MsgText(601) Then
            ShowMsg "此為國內同業,不可輸入E-Mail以免誤發電子郵件, 如有需要請加註於備註欄 ！"
            txtPOC(9).SetFocus
            txtPOC_GotFocus (9)
            Exit Function
        End If
      End If
      'end 2021/11/29
      
      '有輸入客戶狀態時，不寄雜誌、電子報
      If txtPOC(14) <> "" Then
         txtPOC(11) = "N"
      End If
      
      '檢查英文名稱第一碼
      m_bSaveCheck = True
      'Modify By Sindy 2012/10/31
      'If Trim(txtPOC(4)) <> pub_NationByName(txtPOC(23) & txtPOC(24) & txtPOC(25) & txtPOC(26), Trim(txtPOC(4)), True, "客戶") Then
      If Trim(Left(Combo1.Text, 4)) <> pub_NationByName(txtPOC(23) & txtPOC(24) & txtPOC(25) & txtPOC(26), Trim(Left(Combo1.Text, 4)), True, "客戶") Then
      '2012/10/31 End
'         If Me.ActiveControl = txtPOC(4) Then
'            txtPOC_GotFocus 4
'         Else
'            txtPOC(4).SetFocus
'         End If
         Combo1.SetFocus 'Added by Lydia 2016/08/10
         Exit Function
      End If
      m_bSaveCheck = False
      
      
      
      'Modify by Amy 2014/07/03 開放電腦中心權限 因開拓轉檔轉入可能無國籍資料,但要能修改 故不鎖
      'modify by sonia 2019/4/12邱素蓮調職改成莊敏惠73017
      'modify by sonia 2019/5/15再改為開放'北所業務助理人員'權限
      'If strUserNum <> "73017" And Pub_StrUserSt03 <> "M51" Then '開放邱素蓮權限
      If InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) = 0 And Pub_StrUserSt03 <> "M51" Then
      'end 2019/5/15
         If Combo1.Text = "" Then
            ShowMsg "客戶國籍不可為空白 !"
            Call Combo1_GotFocus
            Exit Function
         End If
         If Trim(txtPOC(10).Text) = "" And Trim(txtPOC(9).Text) = "" Then
          ShowMsg "地址和E-Mail至少輸入一項 !"
          txtPOC(9).SetFocus
          Exit Function
        End If
        If txtPOC(12).Text = "" Then
          ShowMsg "開發日期不可空白 !"
          txtPOC(12).SetFocus
          Exit Function
        End If
      End If
      If txtPOC(13).Text = "" Then
        ShowMsg "智權人員不可空白!"
        txtPOC(13).SetFocus
        txtPOC_GotFocus 13
        Exit Function
      'Add by  Amy 2016/09/05
      ElseIf LBL1(2) = "" Then
        ShowMsg "智權人員有誤請確認!"
        txtPOC(13).SetFocus
        txtPOC_GotFocus 13
        Exit Function
      End If
      
      If m_EditMode = 1 And txtPOC(1) <> "" Then
         strExc(0) = "select count(*) from potcustomer1 where POC01='" & Left(txtPOC(1) & "000", 8) & "' and POC02='" & Left(txtPOC(2) & "0", 1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp(0) > 0 Then
               ShowMsg "客戶編號重覆，請重新輸入 !"
               txtPOC(1).SetFocus
               txtPOC_GotFocus 1
               Exit Function
            End If
         End If
      End If
      
      'Add By Sindy 2014/7/10
      '新增存檔時,若智權人員為81040的資料,則'是否寄電子報'及'是否寄專利雙週報'都不可設為要寄
      If m_EditMode = 1 Then '新增
         If txtPOC(13) = "81040" Then '閻副所長
            If txtPOC(11) = "" Or txtPOC(28) = "" Then
               ShowMsg "若要寄發則先存檔再修改！"
               If txtPOC(11) = "" Then
                  txtPOC(11).SetFocus
                  txtPOC_GotFocus 11
               ElseIf txtPOC(28) = "" Then
                  txtPOC(28).SetFocus
                  txtPOC_GotFocus 28
               End If
               Exit Function
            End If
         End If
      End If
      '2014/7/10 END
   End If
   
   TxtValidate = True
End Function

Private Sub UpdateFieldNewData()
   txtPOC(14).Text = cboStatus.Text
   'Modify By Sindy 2012/10/31
   'txtPOC(4).Text = Left(Trim(Combo1.Text), 3)
   txtPOC(4).Text = Left(Trim(Combo1.Text), 4)
   '2012/10/31 End
   For Each oText In txtPOC
      idx = oText.Index
      Select Case idx
         Case 12 '開發日期
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            'Modify By Sindy 2013/11/15 國籍Trim掉
            If idx = 4 Then
               m_FieldList(idx).fiNewData = Trim(oText.Text)
            Else
               m_FieldList(idx).fiNewData = oText.Text
            End If
      End Select
   Next
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
   Dim stSQL As String, stCols As String, stValues As String
   Dim intR As Integer 'Added by Lydia 2024/01/05
   
   'Move by Lydia 2024/01/05 從cnnConnection.BeginTrans下面搬上來
   If txtPOC(1) = "" Then
JumpToReNo: 'Added by Lydia 2024/01/05
      If ClsPDGetAutoNumber("R", strTmp, True, False) Then
         strTmp = "R" + Right(strTmp, 5) & "00"
         'Added by Lydia 2024/01/05 防止編號重覆
         stSQL = "select pcu01 from potcustomer where pcu01='" & strTmp & "' " & _
                 "union all select poc01 from potcustomer1 where poc01='" & strTmp & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            If intR > 3 Then
               MsgBox "存檔失敗!", vbCritical, "流水號給號"
               Exit Function
            End If
            GoTo JumpToReNo
         End If
         txtPOC(1) = strTmp
         'end 2024/01/05
         m_FieldList(1).fiNewData = strTmp
         m_FieldList(2).fiNewData = "0"
      End If
   End If
   'end --- Move by Lydia 2024/01/05 從cnnConnection.BeginTrans下面搬上來
   

On Error GoTo ErrHand
   cnnConnection.BeginTrans

   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtPOC
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
   stSQL = "INSERT INTO PotCustomer1 (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   AddRecord = True
   
   txtPOC(1).Tag = m_FieldList(1).fiNewData
   txtPOC(2).Tag = m_FieldList(2).fiNewData
   
'CANCEL BY SONIA 2013/6/14 不知要做什麼?
'   'Add By Sindy 98/04/14 發mail通知83002該筆潛在客戶狀態
'   If Trim(cboStatus.Text) <> "" And Not IsNull(cboStatus.Text) Then
'      PUB_SendMail strUserNum, "83002", "", "國內潛在客戶狀態通知！", "潛在客戶編號：" + txtPOC(1) + txtPOC(2) & vbCrLf & _
'         "狀態：" + cboStatus.Text
'   End If
'   '98/03/02 End
'2013/6/14 END
   
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
   
   '刪除潛在客戶資料
   stSQL = "delete from PotCustomer1 where POC01='" & txtPOC(1) & "' and POC02='" & txtPOC(2) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   ClearField
   txtPOC(1).Tag = ""
   txtPOC(2).Tag = ""
   
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   'Add by Amy 2014/03/07
   Dim strMail(5) As String
   Dim CountName As Integer
   'end 2014/03/07
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE PotCustomer1 SET "
   stSet = ""
   For Each oText In txtPOC
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
      stSQL = stSQL & stSet & " where POC01='" & txtPOC(1) & "' and POC02='" & txtPOC(2) & "'; end; "
      Pub_SeekTbLog stSQL
      
      cnnConnection.Execute stSQL, intI
   End If
   
   cnnConnection.CommitTrans
   ModRecord = True
   
'CANCEL BY SONIA 2013/6/14 不知要做什麼?
'   'Add By Sindy 98/04/14 發mail通知83002該筆潛在客戶狀態
'   If Trim(cboStatus.Text) <> "" And Not IsNull(cboStatus.Text) Then
'      PUB_SendMail strUserNum, "83002", "", "國內潛在客戶狀態通知！", "潛在客戶編號：" + txtPOC(1) + txtPOC(2) & vbCrLf & _
'         "狀態：" + cboStatus.Text
'   End If
'   '98/03/02 End
'2013/6/14 END
   
   
   'Add by Amy 2014/03/07 當研發處修改是否寄電子報/專利雙週報為N時發mail通知智權人員
   Erase strMail: CountName = 0
   If m_EditMode = 2 And Pub_StrUserSt03 = "D01" Then
         '組客戶名稱
         For intI = 23 To 26
            If Trim(txtPOC(intI)) <> MsgText(601) Then
                strMail(0) = strMail(0) & txtPOC(intI) & " "
            Else
                Exit For
            End If
         Next intI
         If strMail(0) <> MsgText(601) Then strMail(0) = strMail(0) & "(英文)" & vbCrLf: CountName = CountName + 1
         If txtPOC(3) <> MsgText(601) Then strMail(0) = txtPOC(3) & "(中文)" & vbCrLf & strMail(0): CountName = CountName + 1
         If txtPOC(27) <> MsgText(601) Then strMail(0) = strMail(0) & txtPOC(27) & "(日文)" & vbCrLf: CountName = CountName + 1
         
         '組mail 內容
         strMail(1) = "客戶編號：" & txtPOC(1) & " " & txtPOC(2) & vbCrLf & _
                            "客戶名稱：" & Replace(strMail(0), vbCrLf, vbCrLf & String(5, "　")) & vbCrLf & _
                            "E-mail  ：" & txtPOC(9) & vbCrLf & vbCrLf & _
                            "研發處已將客戶資料的"
                                      
         strMail(2) = " 欄改為 N," & vbCrLf & _
                            "請詳查原因後, 自行修正客戶資料內容, 謝謝 !"
        If oldPOC11 = MsgText(601) And txtPOC(11) = "N" Then strMail(3) = "'電子報'"
        If oldPOC28 = MsgText(601) And txtPOC(28) = "N" Then strMail(4) = "'專利雙週報'"
        
        If txtPOC(13) = "001-1" Then
            '若智權人為業務助理發給邱素蓮
            'modify by sonia 2019/4/12邱素蓮調職改成莊敏惠
            'modify by sonia 2019/5/15再改智權委辦區ip_transfer
            strMail(5) = "ip_transfer"
        Else
            strMail(5) = txtPOC(13)
        End If
           
        If strMail(3) <> MsgText(601) And strMail(4) <> MsgText(601) Then
            strMail(1) = strMail(1) & strMail(3) & "及" & strMail(4) & strMail(2)
            PUB_SendMail strUserNum, strMail(5), "", "客戶E-mail 信箱於此次寄發" & strMail(3) & "及" & strMail(4) & "時遭退回通知 !", strMail(1)
        ElseIf strMail(3) <> MsgText(601) Or strMail(4) <> MsgText(601) Then
            strMail(1) = strMail(1) & IIf(strMail(3) = "", strMail(4), strMail(3)) & strMail(2)
            PUB_SendMail strUserNum, strMail(5), "", "客戶E-mail 信箱於此次寄發" & IIf(strMail(3) = "", strMail(4), strMail(3)) & "時遭退回通知 !", strMail(1)
        End If
           
   End If
   'end 201/003/07
   
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

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
         'Modify by Morgan 2008/1/24 +刪除檢查
         If PUB_POCDelCheck(txtPOC(1), txtPOC(2)) = True Then
            If DelRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 2
            End If
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtPOC(1).SetFocus
               txtPOC_GotFocus 1
            End If
         End If
         
   End Select
End Function

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim adoRst As New ADODB.Recordset
   
Top:
'   Select Case p_iWay
'      Case 0
'         strExc(0) = "SELECT * FROM PotCustomer1" & _
'            " WHERE POC01 = '" & Left(txtPOC(1).Tag & "000", 8) & "' AND POC02 = '" & Left(txtPOC(2).Tag & "0", 1) & "'"
'      Case -2
'         strExc(0) = "SELECT * FROM PotCustomer1 order by POC01 ASC,POC02 ASC"
'      Case -1
'         strExc(0) = "SELECT * FROM PotCustomer1" & _
'            " WHERE POC01||POC02 <'" & Left(txtPOC(1).Tag & "000", 8) & Left(txtPOC(2).Tag & "0", 1) & "' order by POC01 DESC,POC02 DESC"
'      Case 1
'         strExc(0) = "SELECT * FROM PotCustomer1" & _
'            " WHERE POC01||POC02 >'" & Left(txtPOC(1).Tag & "000", 8) & Left(txtPOC(2).Tag & "0", 1) & "' order by POC01 ASC,POC02 ASC"
'      Case 2
'         strExc(0) = "SELECT * FROM PotCustomer1 order by POC01 DESC,POC02 DESC"
'      Case 99
'         '2009/6/9 modify by sonia
'         'strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC13='" & strUserNum & "' order by POC01 ASC,POC02 ASC"
'         If Pub_StrUserSt03 = "M51" Then
'            strExc(0) = "SELECT * FROM PotCustomer1 order by POC01 ASC,POC02 ASC"
'         Else
'            strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC13='" & strUserNum & "' order by POC01 ASC,POC02 ASC"
'         End If
'         '2009/6/9 end
'   End Select
   'Modify By Sindy 2012/10/4
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM PotCustomer1" & _
            " WHERE POC01 = '" & Left(txtPOC(1).Tag & "000", 8) & "' AND POC02 = '" & Left(txtPOC(2).Tag & "0", 1) & "'"
      Case -2
         strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC01||POC02 in (SELECT min(POC01||POC02) FROM PotCustomer1)"
      Case -1
         strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC01||POC02 in (" & _
            "SELECT max(POC01||POC02) FROM PotCustomer1" & _
            " WHERE POC01||POC02 <'" & Left(txtPOC(1).Tag & "000", 8) & Left(txtPOC(2).Tag & "0", 1) & "')"
      Case 1
         strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC01||POC02 in (" & _
            "SELECT min(POC01||POC02) FROM PotCustomer1" & _
            " WHERE POC01||POC02 >'" & Left(txtPOC(1).Tag & "000", 8) & Left(txtPOC(2).Tag & "0", 1) & "')"
      Case 2
         strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC01||POC02 in (SELECT max(POC01||POC02) FROM PotCustomer1)"
      Case 99
         '2009/6/9 modify by sonia
         'strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC13='" & strUserNum & "' order by POC01 ASC,POC02 ASC"
         'Modify by Amy 2015/06/10 +等級 F1,F2 及部門別M71的人員可以修改所有資料
         If Pub_StrUserSt03 = "M51" Or PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "F2" Or Pub_StrUserSt03 = "M71" Then
            strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC01||POC02 in (SELECT min(POC01||POC02) FROM PotCustomer1)"
         Else
            'Modify By Sindy 2012/10/4 建檔人員為strUserNum也可查詢
            'strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC01||POC02 in (SELECT min(POC01||POC02) FROM PotCustomer1 WHERE POC13='" & strUserNum & "')"
            strExc(0) = "SELECT * FROM PotCustomer1 WHERE POC01||POC02 in (SELECT min(POC01||POC02) FROM PotCustomer1 WHERE POC13='" & strUserNum & "' or POC17='" & strUserNum & "')"
         End If
         '2009/6/9 end
   End Select
   '2012/10/4 End
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If CheckModifyLimit(adoRst.Fields("POC13"), False, adoRst.Fields("POC17")) = False Then
         txtPOC(1).Tag = adoRst.Fields("POC01")
         txtPOC(2).Tag = adoRst.Fields("POC02")
         Set adoRst = Nothing
         '0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
         If p_iWay = -2 Then
            p_iWay = 1
            GoTo Top
         ElseIf p_iWay = 2 Then
            p_iWay = -1
            GoTo Top
         ElseIf p_iWay = 0 Then
            MsgBox "您沒有此筆潛在客戶維護權限 !!!", vbInformation
'            p_iWay = 1
'            GoTo Top
            'Add By Sindy 2023/9/6
            txtPOC(1).Tag = ""
            txtPOC(2).Tag = ""
            Exit Function
            '2023/9/6 END
         Else
            If p_iWay = 99 Then 'Modify By Sindy 2022/1/12 ex:R18583 / 82021.楊世安
               GoTo gotoExit
            Else
               GoTo Top
            End If
         End If
         'GoTo Top
      End If
      ClearField
      If Not IsNull(adoRst.Fields("POC16")) Then
         Call GetCustData(adoRst.Fields("POC16"))
      End If
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = 0 Then
         'Add By Sindy 2023/9/6
         txtPOC(1).Tag = ""
         txtPOC(2).Tag = ""
         '2023/9/6 END
         MsgBox "查無資料！", vbInformation
      ElseIf p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
         '2008/12/10 ADD BY SONIA
         Set adoRst = Nothing
         txtPOC(1).Tag = txtPOC(1)
         txtPOC(2).Tag = txtPOC(2)
         p_iWay = 0
         'GoTo Top
         '2008/12/10 END
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
         '2008/12/10 ADD BY SONIA
         Set adoRst = Nothing
         txtPOC(1).Tag = txtPOC(1)
         txtPOC(2).Tag = txtPOC(2)
         p_iWay = 0
         'GoTo Top
         '2008/12/10 END
      Else
         ClearField
         MsgBox "查無資料！", vbInformation
      End If
   End If
   
gotoExit:
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtPOC(1).SetFocus
      txtPOC_GotFocus 1
   End If
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtPOC
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
            
            If idx = 12 Then '開發日期
               oText.Text = ChangeWStringToTString(m_FieldList(idx).fiOldData)
            ElseIf idx = 14 Then '狀態
               If IsNull(m_FieldList(idx).fiOldData) = False Then: cboStatus = m_FieldList(idx).fiOldData
               oText.Text = m_FieldList(idx).fiOldData
            ElseIf idx = 4 Then '國籍
               If IsNull(m_FieldList(idx).fiOldData) = False Then: Combo1 = m_FieldList(idx).fiOldData
               Call Combo1_Validate(False)
               'Modify By Sindy 2013/11/15 +trim
               'oText.Text = m_FieldList(idx).fiOldData
               oText.Text = Trim(m_FieldList(idx).fiOldData)
               '2013/11/15 END
            Else
               'Add by Amy 2014/03/07
               If idx = 11 Then oldPOC11 = m_FieldList(idx).fiOldData
               If idx = 28 Then oldPOC28 = m_FieldList(idx).fiOldData
               'end 2014/03/07
               oText.Text = m_FieldList(idx).fiOldData
            End If
         Next
         CUID(1) = "" & .Fields("POC17"): txtPOC(17) = "" & .Fields("POC17")
         CUID(2) = "" & .Fields("POC18")
         CUID(3) = "" & .Fields("POC19")
         CUID(4) = "" & .Fields("POC20")
         CUID(5) = "" & .Fields("POC21")
         CUID(6) = "" & .Fields("POC22")
         
         '智權人員姓名
         If IsNull(txtPOC(13)) = False Then
            If IsEmptyText(txtPOC(13)) = False Then
               'Modify by Amy 2016/09/05 顯示離職員工
               LBL1(2).Caption = GetStaffName(txtPOC(13), True)
            End If
         End If
         
      End If
   End With
   UpdateCUID CUID, textCUID
   txtPOC(1).Tag = txtPOC(1)
   txtPOC(2).Tag = txtPOC(2)
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   For Each oText In txtPOC
      idx = oText.Index
      m_FieldList(idx).fiName = "POC" & Format(idx, "00")
   Next
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As TextBox)
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

Private Function GetCustData(p_stCust As String) As Boolean
   Dim aiOrder(1 To 3) As Integer
   Select Case Left(p_stCust, 1)
      Case "X"
         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
      Case Else
         MsgBox "關係企業必須為 X 開頭", vbCritical + vbOKOnly, "檢核資料"
         Exit Function
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   txtPOC16N.Text = ""
   If intI = 1 Then
      For intI = 1 To 3
         If Not IsNull(RsTemp(intI)) Then
            txtPOC16N.Text = RsTemp(intI)
            lblPOC16 = "" & RsTemp(intI) 'Added by Lydia 2023/06/15
            Exit For
         End If
      Next
      GetCustData = True
   Else
      MsgBox "關係企業輸入錯誤！"
   End If
End Function

'檢查維護權限
Private Function CheckModifyLimit(strChkID As String, bType As Boolean, strCreateID As String) As Boolean
Dim strUserNumST05 As String
   
   CheckModifyLimit = True
   
   If Trim(strChkID) = "" Then Exit Function
   
   strUserNumST05 = PUB_GetST05(strUserNum)
   
   '2009/5/14 add by sonia 開放M51權限
   '2009/7/14 MODIFY BY SONIA 開放75033夏慧珠的權限
   'Modify By Sindy 2011/1/20 開放雅娟和品薇可以互相維護其建立的資料
   'Modify by Amy 2015/06/10 +等級 F1,F2 及部門別M71的人員可以修改所有資料
   'Modify By Sindy 2015/12/3 Mark:(PUB_GetST05(strChkID) = strUserNumST05 And strUserNumST05 = "73") Or
   'Modify By Sindy 2023/7/31 夏慧珠退休改先開放給杜協理
   If Pub_StrUserSt03 = "M51" Or _
      strUserNum = "74018" Or _
      PUB_GetST05(strUserNum) = "F1" Or _
      PUB_GetST05(strUserNum) = "F2" Or _
      Pub_StrUserSt03 = "M71" Then
      Exit Function
   End If
   '2009/5/14 end
   
   'add by sonia 2019/5/15 開放'北所業務助理人員'可維護001-1資料權限(2014/07/03漏改)
   If InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 And strChkID = "001-1" Then
      Exit Function
   End If
   'end 2019/5/15
   
   'LoginUser須為智權人員或其案件主管, 方可維護此筆資料
   'Modify By Sindy 2022/1/12 建檔人員可以查詢及維護 ex:R18583 / 82021.楊世安
   If strUserNum = Trim(strChkID) Or strUserNum = Trim(strCreateID) Then
      Exit Function
   Else
      'Modify By Sindy 2012/10/4 +STAFF B
      'modify by sonia 2017/8/10 A0909->A0908
      strExc(0) = "SELECT A0908,B.ST03,A.ST03 FROM STAFF A,ACC090,STAFF B " & _
                         "WHERE A.ST03=A0901(+) and A.ST01= '" & Trim(strChkID) & "' and B.ST01= '" & Trim(strCreateID) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If strUserNum = RsTemp(0) Then Exit Function '開發人員的主管也可以修改
         'Add By Sindy 2012/10/4 若建檔人員為打字室,則LoginUser同部門者,資料可修改
         If Trim(RsTemp(1)) = "M13" And Pub_StrUserSt03 = Trim(RsTemp(1)) Then Exit Function
         '2012/10/4 End
         'Add By Sindy 2015/12/3 建檔人員為P12程序, 互相可修改
         '                    或 開發人員為P12程序, 則P12人員可互相修改
         If Trim(RsTemp(1)) = "P12" And Pub_StrUserSt03 = Trim(RsTemp(1)) Then Exit Function
         If Trim(RsTemp(2)) = "P12" And Pub_StrUserSt03 = Trim(RsTemp(2)) Then Exit Function
         '2015/12/3 END
      End If
   End If
   
   CheckModifyLimit = False
   If bType = True Then
      MsgBox "無修改權限 !!!", vbInformation
   End If
End Function

Private Sub cboStatus_GotFocus()
   OpenIme
End Sub

Private Sub cboStatus_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   Select Case cboStatus
      Case "", "刪址", "倒閉", "遷移不明", "解散", "廢止", "撤銷", "停業", "往生", "業務自行處理", "不再使用"
      'Add by Amy 2015/03/23 電腦中心及研發處 +"開拓不寄"
      Case "開拓不寄"
         If Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "D01" Then
            ShowMsg "客戶狀態錯誤, 請以下拉方式點選 !"
            Cancel = True
         End If
      Case Else
         'Add by Amy 2015/03/23 +if 電腦中心可以自行輸入非上述狀態
         If Pub_StrUserSt03 <> "M51" Then
            ShowMsg "客戶狀態錯誤, 請以下拉方式點選 !"
            Cancel = True
         End If
   End Select
End Sub

Private Sub Combo1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox Combo1
End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
If Combo1.Text <> "" Then
    Dim MyArr As Variant
    Dim MyArr2 As Variant
    Dim Myi As Integer
    MyArr = Split(Combo1, " ")
    For Myi = 0 To Combo1.ListCount - 1
        MyArr2 = Split(Combo1.List(Myi), " ")
        If MyArr(0) = MyArr2(0) Then
            Combo1.Text = Combo1.List(Myi)
            Exit Sub
        End If
    Next Myi
'    If m_EditMode <> 0 Then
        MsgBox "國籍代號輸入錯誤!!!", vbExclamation + vbOKOnly
        Call Combo1_GotFocus
        Cancel = True
        Exit Sub
'    End If
End If
End Sub

Private Sub txtPOC16N_GotFocus()
'Memo by Lydia 2023/06/15 杜燕文協理要求改成用Label顯示關聯企業名稱，原TextBox隱藏
   OpenIme
   TextInverse txtPOC16N
End Sub

Private Sub txtPOC16N_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub txtPOC16N_Validate(Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If txtPOC(16).Text = "" And txtPOC16N.Text <> "" Then
'         If txtSameCnt = "Y" Then
'            Me.Show
'            Me.txtPOC(5).SetFocus
'            txtSameCnt = ""
'         ElseIf txtSameCnt = "E" Then
'            Me.Show
'            Me.txtPOC16N.SetFocus
'            Cancel = True
'            txtSameCnt = ""
'            Exit Sub
'         Else
'            Me.Enabled = False
'            Screen.MousePointer = vbHourglass
'            frm210128_2.Show
'            'Modify By Sindy 2009/06/23
'            'frm210128_2.StrMenu
'            Call frm210128_2.StrMenu("1", Me.txtPOC16N.Text)
'            '2009/06/23 End
'            Screen.MousePointer = vbDefault
'            Me.Enabled = True
'            Me.Hide
'            If txtSameCnt.Text = 0 Then txtSameCnt.Text = ""
'            If Trim(txtPOC(16).Text) <> "" And Not IsNull(txtPOC(16)) Then
'               Me.Show
'               Me.txtPOC(5).SetFocus
'            ElseIf txtSameCnt.Text = "" Then
'               Me.Show
'               Me.txtPOC16N.SetFocus
'               Cancel = True
'               Exit Sub
'            End If
'         End If
         'Modify By Sindy 2014/2/27
         txtSameCnt = ""
         If frm210128_2.StrMenu("1", Me.txtPOC16N.Text) = True Then
            Me.Hide
            frm210128_2.Show vbModal
            If txtSameCnt = "Y" Then
               Me.Show
               Me.txtPOC(5).SetFocus
               txtSameCnt = ""
            ElseIf txtSameCnt = "E" Then
               Me.Show
               Me.txtPOC16N.SetFocus
               Cancel = True
               txtSameCnt = ""
               Exit Sub
            End If
         Else
            Unload frm210128_2
            Me.Show
            If Trim(txtPOC(16).Text) <> "" And Not IsNull(txtPOC(16)) Then
               Call txtPOC_GotFocus(5)
            ElseIf txtSameCnt.Text = "" Then
               Call txtPOC16N_GotFocus
               Cancel = True
               Exit Sub
            End If
         End If
         '2014/2/27 END
      End If
   End If
End Sub
