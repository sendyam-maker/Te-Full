VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100123_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標延展結案說明輸入"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9390
   Begin VB.CommandButton CImg 
      BackColor       =   &H00C0C0FF&
      Height          =   320
      Index           =   0
      Left            =   120
      Picture         =   "frm100123_1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   70
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   9
      Left            =   4560
      Picture         =   "frm100123_1.frx":0073
      Style           =   1  '圖片外觀
      TabIndex        =   69
      Top             =   4970
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   8
      Left            =   4560
      Picture         =   "frm100123_1.frx":00E6
      Style           =   1  '圖片外觀
      TabIndex        =   68
      Top             =   4490
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   7
      Left            =   4560
      Picture         =   "frm100123_1.frx":0159
      Style           =   1  '圖片外觀
      TabIndex        =   67
      Top             =   4010
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   6
      Left            =   4560
      Picture         =   "frm100123_1.frx":01CC
      Style           =   1  '圖片外觀
      TabIndex        =   66
      Top             =   3530
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   5
      Left            =   4560
      Picture         =   "frm100123_1.frx":023F
      Style           =   1  '圖片外觀
      TabIndex        =   65
      Top             =   3050
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   4
      Left            =   4560
      Picture         =   "frm100123_1.frx":02B2
      Style           =   1  '圖片外觀
      TabIndex        =   64
      Top             =   2570
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   3
      Left            =   4560
      Picture         =   "frm100123_1.frx":0325
      Style           =   1  '圖片外觀
      TabIndex        =   63
      Top             =   2090
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   2
      Left            =   4560
      Picture         =   "frm100123_1.frx":0398
      Style           =   1  '圖片外觀
      TabIndex        =   62
      Top             =   1610
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   1
      Left            =   4560
      Picture         =   "frm100123_1.frx":040B
      Style           =   1  '圖片外觀
      TabIndex        =   61
      Top             =   1130
      Width           =   255
   End
   Begin VB.CommandButton CmdImg 
      BackColor       =   &H00C0C0FF&
      Height          =   350
      Index           =   0
      Left            =   4560
      Picture         =   "frm100123_1.frx":047E
      Style           =   1  '圖片外觀
      TabIndex        =   60
      Top             =   650
      Width           =   255
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "回前畫面"
      Default         =   -1  'True
      Height          =   400
      Left            =   8400
      TabIndex        =   11
      Top             =   50
      Width           =   980
   End
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00C0C0FF&
      Caption         =   "回覆單匯入"
      Height          =   400
      Left            =   5560
      Style           =   1  '圖片外觀
      TabIndex        =   58
      Top             =   50
      Width           =   1065
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "複製第一筆結案備註"
      Height          =   400
      Left            =   6640
      Style           =   1  '圖片外觀
      TabIndex        =   57
      Top             =   50
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   400
      Left            =   7750
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   50
      Width           =   600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   9
      Left            =   4845
      TabIndex        =   8
      Top             =   4890
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   8
      Left            =   4845
      TabIndex        =   7
      Top             =   4410
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   7
      Left            =   4845
      TabIndex        =   6
      Top             =   3930
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   6
      Left            =   4845
      TabIndex        =   5
      Top             =   3450
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   5
      Left            =   4845
      TabIndex        =   4
      Top             =   2970
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   4
      Left            =   4845
      TabIndex        =   32
      Top             =   2490
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   3
      Left            =   4845
      TabIndex        =   3
      Top             =   2010
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   2
      Left            =   4845
      TabIndex        =   2
      Top             =   1530
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   1
      Left            =   4845
      TabIndex        =   1
      Top             =   1050
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Index           =   0
      Left            =   4845
      TabIndex        =   0
      Top             =   570
      Width           =   4500
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "7937;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   315
      Left            =   2610
      TabIndex        =   71
      Top             =   30
      Width           =   1065
      VariousPropertyBits=   746604571
      Size            =   "1879;556"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   4560
      Picture         =   "frm100123_1.frx":04F1
      Top             =   270
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "：已有回覆單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   200
      Index           =   10
      Left            =   350
      TabIndex        =   59
      Top             =   80
      Width           =   1995
   End
   Begin VB.Label LCaseN 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   56
      Top             =   4890
      Width           =   1200
   End
   Begin MSForms.Label Lbl9 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   55
      Top             =   4890
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl9 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   54
      Top             =   4890
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl9 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   53
      Top             =   4890
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl8 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   52
      Top             =   4410
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl8 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   51
      Top             =   4410
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl8 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   50
      Top             =   4410
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   49
      Top             =   4410
      Width           =   1200
   End
   Begin MSForms.Label Lbl7 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   48
      Top             =   3930
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl7 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   47
      Top             =   3930
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl7 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   46
      Top             =   3930
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   45
      Top             =   3930
      Width           =   1200
   End
   Begin MSForms.Label Lbl6 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   44
      Top             =   3450
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl6 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   43
      Top             =   3450
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl6 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   42
      Top             =   3450
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   41
      Top             =   3450
      Width           =   1200
   End
   Begin MSForms.Label Lbl5 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   40
      Top             =   2970
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl5 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   39
      Top             =   2970
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl5 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   38
      Top             =   2970
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   37
      Top             =   2970
      Width           =   1200
   End
   Begin MSForms.Label Lbl4 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   36
      Top             =   2490
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl4 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   35
      Top             =   2490
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl4 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   34
      Top             =   2490
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   33
      Top             =   2490
      Width           =   1200
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   31
      Top             =   2010
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   30
      Top             =   2010
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   29
      Top             =   2010
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   2010
      Width           =   1200
   End
   Begin MSForms.Label Lbl2 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   27
      Top             =   1530
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl2 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   26
      Top             =   1530
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl2 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   25
      Top             =   1530
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   1530
      Width           =   1200
   End
   Begin MSForms.Label Lbl1 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   23
      Top             =   1050
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl1 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   22
      Top             =   1050
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl1 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   21
      Top             =   1050
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1050
      Width           =   1200
   End
   Begin MSForms.Label Lbl0 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   19
      Top             =   600
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2328;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl0 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   18
      Top             =   600
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl0 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   17
      Top             =   600
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1270;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LCaseN 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   345
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   1410
      TabIndex        =   14
      Top             =   345
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "法定期限"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   2280
      TabIndex        =   13
      Top             =   345
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      Top             =   345
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   5085
      TabIndex        =   9
      Top             =   345
      Width           =   390
   End
End
Attribute VB_Name = "frm100123_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/15 改成Form2.0 ; Text1(index)、Lbl0(index)、Lbl1(index)、Lbl2(index)、Lbl3(index)、Lbl4(index)、Lbl5(index)、Lbl6(index)、Lbl7(index)、Lbl8(index)、Lbl9(index)
'Create by Amy 2018/05/24
Option Explicit

Public m_strSaveFiles As String '新增附件
Dim i As Integer
Dim oLbl As Label, oTxt As TextBox, oCmdImg As CommandButton
Dim m_PrevForm As Form '前一畫面
Dim intShow As Integer
Dim strAllCaseNo As String
Dim strNP22() As String, strCP09() As String
Dim strErr As String '錯誤的本所號及訊息
Dim bolPreForm As Boolean '回前畫面

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Public Sub StrMenu(ByRef strData() As String, intEnd As Integer)
    Dim j As Integer
    Dim strTmp
    ReDim strNP22(intEnd)
    ReDim strCP09(intEnd)
    
    intShow = intEnd
    For i = LBound(strData) To intShow
        If strData(i) <> MsgText(601) Then
            strTmp = Split(strData(i), ",")
            LCaseN(i).Caption = strTmp(LBound(strTmp)) '本所案號
            strAllCaseNo = strAllCaseNo & "," & LCaseN(i).Caption
            For Each oLbl In Me.Controls("Lbl" & i)
                oLbl.Caption = strTmp(oLbl.Index)
            Next
            strCP09(i) = strTmp(UBound(strTmp) - 1) '總收文號
            strNP22(i) = strTmp(UBound(strTmp)) '下一程序序號
            Text1(i).Visible = True
        End If
    Next i
    strAllCaseNo = Mid(strAllCaseNo, 2)
End Sub

Private Sub cmdCopy_Click()
    If Text1(0) = MsgText(601) Then Exit Sub
    If Text1(1).Visible = False Then Exit Sub
    
    For i = 1 To intShow
        If Text1(i) = MsgText(601) Then
            Text1(i) = Text1(0)
        End If
    Next i
End Sub

Private Sub cmdExit_Click()
    bolPreForm = True
    Unload Me
End Sub

Private Sub cmdFile_Click()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
        
    Call frm090801_8.SetParent(Me)
    frm090801_8.m_strSaveFiles = strAllCaseNo
    frm090801_8.lblCaseNo.Visible = False
    frm090801_8.Show vbModal
    
    '顯示夾檔圖示
    If ChkTempTB = True Then
        For Each oLbl In LCaseN
            strCP01 = SystemNumber(oLbl.Caption, 1)
            strCP02 = SystemNumber(oLbl.Caption, 2)
            strCP03 = SystemNumber(oLbl.Caption, 3)
            strCP04 = SystemNumber(oLbl.Caption, 4)
            
            strQ = "Select R001||'-'||R002||'-'||R003||'-'||R004 as CaseNo,R005 as sPath,R006 as FileN From R090801_8 " & _
                  "Where ID='" & strUserNum & "' And R007='" & UCase(Me.Name) & "' And R005 is not null " & _
                  "And R001='" & strCP01 & "' And R002='" & strCP02 & "' And R003='" & strCP03 & "' And R004='" & strCP04 & "' "
            If RsQ.State <> 0 Then RsQ.Close
            RsQ.CursorLocation = adUseClient
            RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
            If RsQ.RecordCount > 0 Then
                CmdImg(oLbl.Index).Visible = True
            End If
        Next
    End If
    
End Sub

Private Sub cmdOK_Click()
    Dim strMsg As String
    
    'Added by Lydia 2022/02/15 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Sub
    End If
    'end 2022/02/15
    If ChkTempTB = False Then
        If MsgBox("回覆單未匯入，要匯入回覆單？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            Call cmdFile_Click
            Exit Sub
        End If
    End If

    Call SaveForm
    If strErr <> MsgText(601) Then strErr = Mid(strErr, 2)
    Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    
    For i = 0 To 9
        For Each oLbl In LCaseN
            oLbl.Caption = ""
        Next
        For Each oLbl In Me.Controls("Lbl" & i)
            oLbl.Caption = ""
        Next
        For Each oTxt In Text1
            oTxt = ""
            oTxt.Visible = False
        Next
        For Each oCmdImg In CmdImg
            oCmdImg.Visible = False
        Next
    Next i
    '刪除結案單暫存檔
    strExc(0) = "Delete From R090801_8 Where ID='" & strUserNum & "' And R007='" & UCase(Me.Name) & "' "
    cnnConnection.Execute strExc(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm100123_1 = Nothing
    m_PrevForm.Show
    Call m_PrevForm.SetGrdColor(strErr, bolPreForm)
    Set m_PrevForm = Nothing
End Sub

Private Sub SaveForm()
    Dim RsQ As New ADODB.Recordset, RsF As New ADODB.Recordset
    Dim strQ As String, strF As String, strIns As String
    Dim j As Integer
    Dim strFileName As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
    Dim bolHasFile As Boolean

On Error GoTo CheckingErr
    
    strErr = ""

    For i = 0 To intShow
        strQ = "Select * From T102inform Where ti01=to_number(to_char(sysdate, 'YYYYMMDD')) and ti02='" & strCP09(i) & "' and ti04=" & strNP22(i)
        If RsQ.State <> 0 Then RsQ.Close
        RsQ.CursorLocation = adUseClient
        RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
        If RsQ.RecordCount = 0 Then
            bolHasFile = False: strFileName = ""
            strCP01 = SystemNumber(LCaseN(i), 1)
            strCP02 = SystemNumber(LCaseN(i), 2)
            strCP03 = SystemNumber(LCaseN(i), 3)
            strCP04 = SystemNumber(LCaseN(i), 4)
            
            '存回覆單
             strF = "Select R001||'-'||R002||'-'||R003||'-'||R004 as CaseNo,R005 as sPath,R006 as FileN From R090801_8 " & _
                    "Where ID='" & strUserNum & "' And R007='" & UCase(Me.Name) & "' And R005 is not null " & _
                    "And R001='" & strCP01 & "' And R002='" & strCP02 & "' And R003='" & strCP03 & "' And R004='" & strCP04 & "' "
            If RsF.State <> 0 Then RsF.Close
            RsF.CursorLocation = adUseClient
            RsF.Open strF, cnnConnection, adOpenStatic, adLockReadOnly
            If RsF.RecordCount > 0 Then
                RsF.MoveFirst
                 Do While RsF.EOF = False
                    strFileName = strFileName & "&" & RsF.Fields("sPath")
                    If PUB_ChkIsReplyFile(strCP01 & strCP02 & strCP03 & strCP04, , , , strNP22(j)) = True Then
                        strErr = strErr & "," & LCaseN(i) & "@@檔案已存系統中！"
                        GoTo NextRec
                    Else
                         bolHasFile = True
                    End If
                    RsF.MoveNext
                 Loop
            End If
            
            cnnConnection.BeginTrans
            If bolHasFile = True Then
                strFileName = Mid(strFileName, 2)
                If PUB_UpdReplyFile(strFileName, "", strCP01, strCP02, strCP03, strCP04, , strNP22(i)) = False Then
                    strErr = strErr & "," & LCaseN(i) & "@@檔案上傳有誤！"
                    cnnConnection.RollbackTrans
                    GoTo NextRec
                End If
            End If
          
            '存結案記錄 : 1.沒有要存回覆單 或是 2.要存回覆單且有對應到電子檔
            strIns = "Insert into t102inform (ti01,ti02,ti03,ti04,ti05) Values (to_number(to_char(sysdate, 'YYYYMMDD')),'" & strCP09(i) & "','" & strUserNum & "'," & strNP22(i) & "," & CNULL(ChgSQL(Text1(i))) & ") "
            cnnConnection.Execute strIns
            '已結案清除案號記錄
            If bolHasFile = True Then Call PUB_DelPCOrgFile(strFileName) '刪除原檔
            cnnConnection.CommitTrans
        End If
NextRec:
    Next i
  
    Exit Sub
    
CheckingErr:
    cnnConnection.RollbackTrans
    strErr = strErr & "," & LCaseN(i) & "@@存檔錯誤"
    'MsgBox (Err.Description)
    GoTo NextRec
End Sub

'確認暫存檔是否有資料
Private Function ChkTempTB() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    ChkTempTB = False
    strQ = "Select R001||'-'||R002||'-'||R003||'-'||R004,R005,R006 From R090801_8 " & _
              "Where ID='" & strUserNum & "' And R007='" & UCase(Me.Name) & "' " & _
              "Order by R001,R002,R003,R004"
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        ChkTempTB = True
    End If
    RsQ.Close
End Function

Private Sub Text1_GotFocus(Index As Integer)
    OpenIme
    TextInverse Text1(Index)
    Text1(Index).SetFocus
End Sub
