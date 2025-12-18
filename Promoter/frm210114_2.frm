VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210114_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任契約書-CFP及P非台灣案"
   ClientHeight    =   8100
   ClientLeft      =   636
   ClientTop       =   1548
   ClientWidth     =   10860
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10860
   Begin VB.CheckBox ChkSeal 
      Caption         =   "用印"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   5220
      TabIndex        =   116
      Top             =   7770
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "空白列印"
      Height          =   330
      Index           =   5
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   115
      Top             =   30
      Width           =   920
   End
   Begin VB.CheckBox ChkDou 
      Caption         =   "多發明人 多申請人"
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   114
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton CmdFind2 
      Caption         =   "搜尋發明人(&I)"
      Height          =   330
      Left            =   8640
      TabIndex        =   113
      Top             =   480
      Width           =   1365
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋申請人(&Q)"
      Height          =   330
      Left            =   8640
      TabIndex        =   112
      Top             =   840
      Width           =   1365
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210114_2.frx":0000
      Left            =   7110
      List            =   "frm210114_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   58
      Top             =   7755
      Width           =   2475
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   59
      Text            =   "2"
      Top             =   60
      Width           =   270
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印        份"
      Height          =   330
      Index           =   0
      Left            =   7416
      TabIndex        =   64
      Top             =   30
      Width           =   1100
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8520
      TabIndex        =   65
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "清空資料"
      Height          =   330
      Index           =   2
      Left            =   6492
      TabIndex        =   63
      Top             =   30
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   0
      TabIndex        =   109
      Top             =   0
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   660
         Style           =   2  '單純下拉式
         TabIndex        =   60
         Top             =   30
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   110
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存文檔"
      Height          =   330
      Index           =   3
      Left            =   4644
      TabIndex        =   61
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "讀取文檔"
      Height          =   330
      Index           =   4
      Left            =   5568
      TabIndex        =   62
      Top             =   30
      Width           =   920
   End
   Begin VB.OptionButton opt1 
      Caption         =   "會稿"
      Height          =   210
      Index           =   0
      Left            =   1035
      TabIndex        =   47
      Top             =   6015
      Width           =   1080
   End
   Begin VB.OptionButton opt1 
      Caption         =   "不會稿"
      Height          =   210
      Index           =   1
      Left            =   2175
      TabIndex        =   48
      Top             =   6015
      Width           =   1080
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9330
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   55
      Left            =   480
      TabIndex        =   45
      Top             =   5670
      Width           =   1035
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1826;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   54
      Left            =   3930
      TabIndex        =   13
      Top             =   4110
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   53
      Left            =   2580
      TabIndex        =   12
      Top             =   4110
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   12
      Left            =   5280
      TabIndex        =   14
      Top             =   4110
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   2745
      TabIndex        =   11
      Top             =   3540
      Width           =   1035
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1826;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   42
      Left            =   6630
      TabIndex        =   44
      Top             =   5385
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "3519;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   41
      Left            =   5280
      TabIndex        =   43
      Top             =   5385
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   19
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   37
      Left            =   30
      TabIndex        =   39
      Top             =   5385
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   38
      Left            =   1305
      TabIndex        =   40
      Top             =   5385
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   39
      Left            =   2580
      TabIndex        =   41
      Top             =   5385
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   40
      Left            =   3930
      TabIndex        =   42
      Top             =   5385
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   36
      Left            =   6630
      TabIndex        =   38
      Top             =   5130
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "3519;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   35
      Left            =   5280
      TabIndex        =   37
      Top             =   5130
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   31
      Left            =   30
      TabIndex        =   33
      Top             =   5130
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   32
      Left            =   1305
      TabIndex        =   34
      Top             =   5130
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   33
      Left            =   2580
      TabIndex        =   35
      Top             =   5130
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   34
      Left            =   3930
      TabIndex        =   36
      Top             =   5130
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   30
      Left            =   6630
      TabIndex        =   32
      Top             =   4875
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "3519;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   29
      Left            =   5280
      TabIndex        =   31
      Top             =   4875
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   25
      Left            =   30
      TabIndex        =   27
      Top             =   4875
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   26
      Left            =   1305
      TabIndex        =   28
      Top             =   4875
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   27
      Left            =   2580
      TabIndex        =   29
      Top             =   4875
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   28
      Left            =   3930
      TabIndex        =   30
      Top             =   4875
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   24
      Left            =   6630
      TabIndex        =   26
      Top             =   4620
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "3519;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   23
      Left            =   5280
      TabIndex        =   25
      Top             =   4620
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   19
      Left            =   30
      TabIndex        =   21
      Top             =   4620
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   20
      Left            =   1305
      TabIndex        =   22
      Top             =   4620
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   21
      Left            =   2580
      TabIndex        =   23
      Top             =   4620
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   22
      Left            =   3930
      TabIndex        =   24
      Top             =   4620
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   18
      Left            =   6630
      TabIndex        =   20
      Top             =   4365
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "3519;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   17
      Left            =   5280
      TabIndex        =   19
      Top             =   4365
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1590
      TabIndex        =   8
      Top             =   2925
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1590
      TabIndex        =   6
      Top             =   2295
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1590
      TabIndex        =   3
      Top             =   1350
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1590
      TabIndex        =   4
      Top             =   1665
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1155
      TabIndex        =   0
      Top             =   405
      Width           =   7365
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12991;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1590
      TabIndex        =   1
      Top             =   720
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1590
      TabIndex        =   2
      Top             =   1035
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1590
      TabIndex        =   5
      Top             =   1980
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1590
      TabIndex        =   7
      Top             =   2610
      Width           =   6930
      VariousPropertyBits=   671105051
      MaxLength       =   74
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   13
      Left            =   30
      TabIndex        =   15
      Top             =   4365
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   14
      Left            =   1305
      TabIndex        =   16
      Top             =   4365
      Width           =   1290
      VariousPropertyBits=   671105051
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "2275;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   15
      Left            =   2580
      TabIndex        =   17
      Top             =   4365
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   270
      Index           =   16
      Left            =   3930
      TabIndex        =   18
      Top             =   4365
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2408;476"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   43
      Left            =   1680
      TabIndex        =   46
      Top             =   5670
      Width           =   1485
      VariousPropertyBits=   671105051
      MaxLength       =   12
      Size            =   "2619;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   44
      Left            =   1440
      TabIndex        =   49
      Top             =   6255
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   56
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   45
      Left            =   1440
      TabIndex        =   50
      Top             =   6555
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   56
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   46
      Left            =   1440
      TabIndex        =   51
      Top             =   6855
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   56
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   47
      Left            =   1440
      TabIndex        =   52
      Top             =   7155
      Width           =   2280
      VariousPropertyBits=   671105051
      MaxLength       =   26
      Size            =   "4022;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   48
      Left            =   4440
      TabIndex        =   53
      Top             =   7155
      Width           =   2445
      VariousPropertyBits=   671105051
      MaxLength       =   22
      Size            =   "4313;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   49
      Left            =   1440
      TabIndex        =   54
      Top             =   7455
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   56
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   50
      Left            =   1440
      TabIndex        =   55
      Top             =   7755
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1244;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   51
      Left            =   2550
      TabIndex        =   56
      Top             =   7770
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   52
      Left            =   3750
      TabIndex        =   57
      Top             =   7770
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   2025
      TabIndex        =   9
      Top             =   3240
      Width           =   6495
      VariousPropertyBits=   671105051
      MaxLength       =   78
      Size            =   "11456;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   1890
      TabIndex        =   10
      Top             =   3540
      Width           =   645
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1138;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊為必填欄位"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   3300
      TabIndex        =   122
      Top             =   6030
      Width           =   1080
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   6135
      TabIndex        =   121
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   1170
      TabIndex        =   120
      Top             =   7515
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   1200
      TabIndex        =   119
      Top             =   6300
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   118
      Top             =   6030
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   117
      Top             =   2205
      Width           =   180
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      Caption         =   "受任人："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6360
      TabIndex        =   111
      Top             =   7800
      Width           =   720
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      Caption         =   "八、其他"
      Height          =   180
      Left            =   8865
      TabIndex        =   108
      Top             =   3390
      Width           =   720
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "七、專利調查"
      Height          =   180
      Left            =   8865
      TabIndex        =   107
      Top             =   3210
      Width           =   1080
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "六、專利調卷"
      Height          =   180
      Left            =   8865
      TabIndex        =   106
      Top             =   3030
      Width           =   1080
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      Caption         =   "五、讓渡或授權"
      Height          =   180
      Left            =   8865
      TabIndex        =   105
      Top             =   2850
      Width           =   1260
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "四、救濟程序"
      Height          =   180
      Left            =   8865
      TabIndex        =   104
      Top             =   2670
      Width           =   1080
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      Caption         =   "三、領證及繳納年費程序"
      Height          =   180
      Left            =   8865
      TabIndex        =   103
      Top             =   2490
      Width           =   1980
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "二、中間處裡程序"
      Height          =   180
      Left            =   8865
      TabIndex        =   102
      Top             =   2310
      Width           =   1440
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "一、提出申請程序"
      Height          =   180
      Left            =   8865
      TabIndex        =   101
      Top             =   2130
      Width           =   1440
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "第二條 委辦範圍"
      Height          =   180
      Left            =   8550
      TabIndex        =   100
      Top             =   1890
      Width           =   1305
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6750
      TabIndex        =   99
      Top             =   5700
      Width           =   900
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      Caption         =   "指定聯絡人及通訊地址："
      Height          =   180
      Left            =   60
      TabIndex        =   98
      Top             =   3300
      Width           =   1980
   End
   Begin VB.Label Label34 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "專利種類"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1605
      TabIndex        =   96
      Top             =   4050
      Width           =   735
   End
   Begin VB.Label Label35 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "備　　　註"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7155
      TabIndex        =   97
      Top             =   4035
      Width           =   915
   End
   Begin VB.Label Label33 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   6630
      TabIndex        =   95
      Top             =   3855
      Width           =   1995
   End
   Begin VB.Label Label32 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      Caption         =   "金　　　　　　　　　　　額"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2580
      TabIndex        =   94
      Top             =   3855
      Width           =   4065
   End
   Begin VB.Label Label30 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "國　　別"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   93
      Top             =   4035
      Width           =   735
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1260
      TabIndex        =   92
      Top             =   2940
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1260
      TabIndex        =   91
      Top             =   2670
      Width           =   300
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1260
      TabIndex        =   90
      Top             =   2310
      Width           =   300
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1260
      TabIndex        =   89
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1260
      TabIndex        =   88
      Top             =   1680
      Width           =   300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1260
      TabIndex        =   87
      Top             =   1410
      Width           =   300
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "發明人地址："
      Height          =   180
      Left            =   60
      TabIndex        =   86
      Top             =   1530
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1260
      TabIndex        =   85
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1260
      TabIndex        =   84
      Top             =   810
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利之名稱："
      Height          =   180
      Left            =   60
      TabIndex        =   83
      Top             =   450
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請人姓名："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   82
      Top             =   2205
      Width           =   1080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請人地址："
      Height          =   180
      Left            =   60
      TabIndex        =   81
      Top             =   2850
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "一、乙方受委辦前條第　　　　款　　　　　　程序之費用（包括國外代理人費用），約定如下："
      Height          =   180
      Left            =   45
      TabIndex        =   80
      Top             =   3615
      Width           =   7740
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "合計"
      Height          =   180
      Left            =   60
      TabIndex        =   79
      Top             =   5700
      Width           =   360
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "元整，於本契約簽定同時由甲方一次付清。"
      Height          =   180
      Left            =   3270
      TabIndex        =   78
      Top             =   5700
      Width           =   3420
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "甲方：委任人"
      Height          =   180
      Left            =   75
      TabIndex        =   77
      Top             =   6300
      Width           =   1080
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代表人："
      Height          =   180
      Left            =   585
      TabIndex        =   76
      Top             =   6615
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Left            =   585
      TabIndex        =   75
      Top             =   6915
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "電　話："
      Height          =   180
      Left            =   585
      TabIndex        =   74
      Top             =   7215
      Width           =   720
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "傳　真："
      Height          =   180
      Left            =   3750
      TabIndex        =   73
      Top             =   7215
      Width           =   720
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "乙方：經手人"
      Height          =   180
      Left            =   30
      TabIndex        =   72
      Top             =   7515
      Width           =   1080
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "中　華　民　國　　　　　年　  　　　　月　　　　　　日"
      Height          =   180
      Left            =   60
      TabIndex        =   71
      Top             =   7800
      Width           =   4770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "發明人姓名："
      Height          =   180
      Left            =   60
      TabIndex        =   70
      Top             =   975
      Width           =   1080
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   30
      TabIndex        =   69
      Top             =   3855
      Width           =   1290
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      Caption         =   "申請費"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2580
      TabIndex        =   68
      Top             =   4110
      Width           =   1365
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   1305
      TabIndex        =   67
      Top             =   3855
      Width           =   1290
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      Caption         =   "實審費"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3930
      TabIndex        =   66
      Top             =   4110
      Width           =   1365
   End
End
Attribute VB_Name = "frm210114_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/18 改成Form2.0 ; txt1(index)、Printer改成Word列印
'Memo by Lydia 2019/07/01 表單名稱:案件委任契約書=>委任契約書
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iCount As Integer
'Add By Sindy 2010/4/30
Public m_strCustCode As String
Public m_blnOneRec As Boolean
'2010/4/30 End
Dim strNowCustNo As String 'Add by Amy 2016/08/19 客戶編號
Dim iPrintC As Integer 'Added by Lydia 2017/03/28 目前列印第幾份
Dim bolAddSeal As Boolean 'Added by Lydia 2017/03/28 是否用印
Dim d_Left As Double, d_Top As Double 'Added by Lydia 2017/04/25 印表機實際列印的左邊界、右邊界
Dim strPrinter As String 'Added by Lydia 2017/04/28
Dim strDetail As String 'Move by Lydia 2017/05/16 記錄內容(從StrMenu移出來)
Dim strCompSeal As String 'Added by Lydia 2020/03/25 記錄"公司名稱|用印編號",用,區隔
'Added by Lydia 2022/01/18  加入圖片用(Word)
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1
'end 2022/01/18
Dim m_TempPDF As String 'Added by Lydia 2022/01/19
Dim m_TempFN As String 'Added by Lydia 2022/01/22

'Add By Sindy 2010/4/29
Private Sub cmdFind_Click()
   Dim strCmpName As String, strMsg As String 'Add by Amy 2016/08/19
   
   If Me.txt1(5).Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(5).SetFocus
      Exit Sub
   End If
   frm090801_1.m_Type = 0  'add by Lydia 2014/9/22
   If ChkDou.Value = 1 Then
     frm090801_1.m_DouChk = True '可複選
   Else
     frm090801_1.m_DouChk = False
   End If
   
   Set frm090801_1.m_frm0908A = Me
   
   frm090801_1.m_strCustChnName = Me.txt1(5).Text
   frm090801_1.lblName.Caption = Me.txt1(5).Text
  
   m_blnOneRec = False
   m_strCustCode = ""
   txt1(5).Tag = "" 'Added by Morgan 2012/9/11
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   Combo2.Tag = "": strNowCustNo = "" 'Add by Amy 2016/08/19
   If m_blnOneRec = True And m_strCustCode <> "" Then
     'Add by Amy 2016/08/19 記錄收據公司別(放於SetCustTxt前避免m_strCustCode被清空)
      strNowCustNo = m_strCustCode
      strCmpName = "Y"
      Combo2.Tag = GetReceiptCmp(Left(strNowCustNo, 8), Mid(strNowCustNo, 9, 1), "CFP", "020", False, strCmpName, Me.Name)
      If Combo2.Tag <> MsgText(601) And Combo2 <> MsgText(601) And Combo2.Tag <> frm210114_1.GetComp(Combo2) Then
        strMsg = "您輸入之收據公司別「" & Combo2 & "」與客戶檔設定值「" & strCmpName & "」不同" & vbCrLf & _
                     "是否依客戶檔設定覆蓋您的輸入值？"
        If MsgBox(strMsg, vbYesNo + vbCritical) = vbYes Then
            'Modified by Lydia 2024/08/06
            'Combo2 = strCmpName
            Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
        End If
      ElseIf strCmpName = MsgText(601) Then
        Combo2.ListIndex = 0
      Else
        'Modified by Lydia 2024/08/06
        'Combo2 = strCmpName
        Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
      End If
      'end 2016/08/19
      'Modify by Amy 2021/05/13 +if 讀取文檔要保留原文檔內容
      If Me.ActiveControl.Name = "cmdFind" Then
        Call SetCustTxt(m_strCustCode)
      End If
      txt1(5).Tag = txt1(5) 'Added by Morgan 2012/9/11
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
'Modified by Lydia 2022/01/18
'Dim tb As TextBox
Dim tb As Control
Dim op As OptionButton
Dim fN As Integer
Dim strBuffer As String
'Modify By Sindy 2010/3/17
'Dim AllObj(0 To 55) As String
'Modify By Sindy 2010/9/9
'Dim AllObj(0 To 57) As String
Dim AllObj(0 To 59) As String
Dim AllObjV As Variant
'Add by Amy 2016/08/19 目前收據公司別
Dim strNowCmp As String
Dim strMsg As String 'Added by Lydia 2019/04/16

   Select Case Index
      Case 0
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(0)) = "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "專利之名稱不可空白！", vbInformation, "錯誤！"
              'txt1(0).SetFocus
              'Txt1_GotFocus 0
              'Exit Sub
              strMsg = strMsg & "、專利之名稱"
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          'Modified by Lydia 2017/04/21 拿掉Trim
          If txt1(1) = "" And txt1(2) = "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "發明人姓名不可空白！", vbInformation, "錯誤！"
              'txt1(1).SetFocus
              'Txt1_GotFocus 1
              'Exit Sub
              strMsg = strMsg & "、發明人姓名"
          End If

          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(5)) = "" And Trim(txt1(6)) = "" Then
              MsgBox "申請人姓名不可空白！", vbInformation, "錯誤！"
              txt1(5).SetFocus
              txt1_GotFocus 5
              Exit Sub
          End If
          'Added by Lydia 2017/03/28 幣別改可輸入
          If Trim(txt1(55).Text) = "" Then
             MsgBox "幣別不可空白！", vbInformation, "錯誤！"
             txt1(55).SetFocus
             txt1_GotFocus 55
             Exit Sub
          End If
          'end 2017/03/28
          If Trim(txt1(13)) = "" And Trim(txt1(14)) = "" And Trim(txt1(15)) = "" And Trim(txt1(19)) = "" And Trim(txt1(20)) = "" And Trim(txt1(21)) = "" And Trim(txt1(25)) = "" And Trim(txt1(26)) = "" And Trim(txt1(27)) = "" And Trim(txt1(31)) = "" And Trim(txt1(32)) = "" And Trim(txt1(33)) = "" And Trim(txt1(37)) = "" And Trim(txt1(38)) = "" And Trim(txt1(39)) = "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "約定事項最少輸入一項！", vbInformation, "錯誤！"
              'Txt1(13).SetFocus
              'Txt1_GotFocus 13
              'Exit Sub
              strMsg = strMsg & "、約定事項：未輸入"
          End If
          If (Trim(txt1(13)) = "" Or Trim(txt1(14)) = "" Or Trim(txt1(15)) = "") And Trim(txt1(13)) & Trim(txt1(14)) & Trim(txt1(15)) <> "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "約定事項應該完整！", vbInformation, "錯誤！"
              'If Trim(txt1(15)) = "" Then
              '    txt1(15).SetFocus
              '    Txt1_GotFocus 15
              'End If
              'If Trim(txt1(14)) = "" Then
              '    txt1(14).SetFocus
              '    Txt1_GotFocus 14
              'End If
              'If Trim(txt1(13)) = "" Then
              '    txt1(13).SetFocus
              '    Txt1_GotFocus 13
              'End If
              'Exit Sub
              If strMsg = "" Or (strMsg <> "" And InStr(strMsg, "約定事項：輸入不完整") = 0) Then strMsg = strMsg & "、約定事項：輸入不完整"
          End If
          If (Trim(txt1(19)) = "" Or Trim(txt1(20)) = "" Or Trim(txt1(21)) = "") And Trim(txt1(19)) & Trim(txt1(20)) & Trim(txt1(21)) <> "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "約定事項應該完整！", vbInformation, "錯誤！"
              'If Trim(txt1(21)) = "" Then
              '    txt1(21).SetFocus
              '    Txt1_GotFocus 21
              'End If
              'If Trim(txt1(20)) = "" Then
              '    txt1(20).SetFocus
              '    Txt1_GotFocus 20
              'End If
              'If Trim(txt1(19)) = "" Then
              '    txt1(19).SetFocus
              '    Txt1_GotFocus 19
              'End If
              'Exit Sub
              If strMsg = "" Or (strMsg <> "" And InStr(strMsg, "約定事項：輸入不完整") = 0) Then strMsg = strMsg & "、約定事項：輸入不完整"
          End If
          If (Trim(txt1(25)) = "" Or Trim(txt1(26)) = "" Or Trim(txt1(27)) = "") And Trim(txt1(25)) & Trim(txt1(26)) & Trim(txt1(27)) <> "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "約定事項應該完整！", vbInformation, "錯誤！"
              'If Trim(txt1(27)) = "" Then
              '    txt1(27).SetFocus
              '    Txt1_GotFocus 27
              'End If
              'If Trim(txt1(26)) = "" Then
              '    txt1(26).SetFocus
              '    Txt1_GotFocus 26
              'End If
              'If Trim(txt1(25)) = "" Then
              '    txt1(25).SetFocus
              '    Txt1_GotFocus 25
              'End If
              'Exit Sub
              If strMsg = "" Or (strMsg <> "" And InStr(strMsg, "約定事項：輸入不完整") = 0) Then strMsg = strMsg & "、約定事項：輸入不完整"
          End If
          If (Trim(txt1(31)) = "" Or Trim(txt1(32)) = "" Or Trim(txt1(33)) = "") And Trim(txt1(31)) & Trim(txt1(32)) & Trim(txt1(33)) <> "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "約定事項應該完整！", vbInformation, "錯誤！"
             ' If Trim(txt1(33)) = "" Then
             '     txt1(33).SetFocus
             '     Txt1_GotFocus 33
             ' End If
             ' If Trim(txt1(32)) = "" Then
             '     txt1(32).SetFocus
             '     Txt1_GotFocus 32
             ' End If
             ' If Trim(txt1(31)) = "" Then
             '     txt1(31).SetFocus
             '     Txt1_GotFocus 31
             ' End If
             ' Exit Sub
              If strMsg = "" Or (strMsg <> "" And InStr(strMsg, "約定事項：輸入不完整") = 0) Then strMsg = strMsg & "、約定事項：輸入不完整"
          End If
          If (Trim(txt1(37)) = "" Or Trim(txt1(38)) = "" Or Trim(txt1(39)) = "") And Trim(txt1(37)) & Trim(txt1(38)) & Trim(txt1(39)) <> "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "約定事項應該完整！", vbInformation, "錯誤！"
              'If Trim(txt1(39)) = "" Then
              '    txt1(39).SetFocus
              '    Txt1_GotFocus 39
              'End If
              'If Trim(txt1(38)) = "" Then
              '    txt1(38).SetFocus
              '    Txt1_GotFocus 38
              'End If
              'If Trim(txt1(37)) = "" Then
              '    txt1(37).SetFocus
              '    Txt1_GotFocus 37
              'End If
              'Exit Sub
              If strMsg = "" Or (strMsg <> "" And InStr(strMsg, "約定事項：輸入不完整") = 0) Then strMsg = strMsg & "、約定事項：輸入不完整"
          End If
          If txt1(43) = "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "費用不可空白！", vbInformation, "錯誤！"
              'txt1(43).SetFocus
              'Txt1_GotFocus 43
              'Exit Sub
              strMsg = strMsg & "、合計(費用)"
          End If
          If opt1(0).Value <> True And opt1(1).Value <> True Then
              MsgBox "會不會稿要選擇一項！", vbInformation, "錯誤！"
              Exit Sub
          End If
          'Modified by Lydia 2019/04/16 必填欄位
      '    If txt1(44) = "" Then
      '        MsgBox "委任人不可空白！", vbInformation, "錯誤！"
      '        txt1(44).SetFocus
      '        txt1_GotFocus 44
      '        Exit Sub
      '    End If
          If Trim(txt1(44)) = "" Then
              MsgBox "委任人不可空白！", vbInformation, "錯誤！"
              txt1(44).SetFocus
              txt1_GotFocus 44
              Exit Sub
          End If
          'end 2019/04/16
          
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(49)) = "" Then
              MsgBox "經手人不可空白！", vbInformation, "錯誤！"
              txt1(49).SetFocus
              txt1_GotFocus 49
              Exit Sub
          End If
          
          '2011/10/18 ADD BY SONIA 檢查四縣市地址(指定聯絡人及地址因地址欄位置不明故不檢查)
          If txt1(3) <> "" Then
            If CheckTaiwanAddr(txt1(3), "000", "發明人地址") = False Then
               txt1(3).SetFocus
               txt1_GotFocus (3)
               Exit Sub
            End If
          End If
          If txt1(7) <> "" Then
            If CheckTaiwanAddr(txt1(7), "000", "申請人地址") = False Then
               txt1(7).SetFocus
               txt1_GotFocus (7)
               Exit Sub
            End If
          End If
          If txt1(46) <> "" Then
            If CheckTaiwanAddr(txt1(46), "000", "甲方委任人地址") = False Then
               txt1(46).SetFocus
               txt1_GotFocus (46)
               Exit Sub
            End If
          End If
          '2011/10/18 END
          'Add by Amy 2016/08/19 +受任人不可為空
          If Combo2 = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          
          'Added by Lydia 2019/04/16  開放部分欄位空白,統一彈訊息
          If strMsg <> "" Then
              If MsgBox("下列欄位空白，是否繼續列印？" & Replace(strMsg, "、", vbCrLf), vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                  Exit Sub
              End If
          End If
          'end 2019/04/16
         
          'Add By Sindy 2013/12/15 檢查是否為不開發票之客戶
          If Combo2 = "台一智權股份有限公司" Then
            'Added by Lydia 2020/07/20 國別若有"台灣"或"中華民國"時，受任人不可選擇"台一智權股份有限公司"
            If Trim(txt1(13) & txt1(19) & txt1(25) & txt1(31) & txt1(37)) <> "" And (InStr(txt1(13) & "," & txt1(19) & "," & txt1(25) & "," & txt1(31) & "," & txt1(37), "台灣") > 0 Or InStr(txt1(13) & "," & txt1(19) & "," & txt1(25) & "," & txt1(31) & "," & txt1(37), "臺灣") > 0 _
                   Or InStr(txt1(13) & "," & txt1(19) & "," & txt1(25) & "," & txt1(31) & "," & txt1(37), "中華民國") > 0) Then
               MsgBox "國別若有""台灣""或""中華民國""時，受任人不可選擇智權公司!!!", vbInformation
               Combo2.SetFocus
               Exit Sub
            End If
            'end 2020/07/20
            'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
            If PUB_ChkCU144isN("", "", txt1(5), "J", , "受任人") = True Then
               Combo2.SetFocus
               Exit Sub
            End If
          End If
          '2013/12/15 END
           '2009/11/13 MODIFY BY SONIA 杜副總提出
      '    If txt1(50) = "" Or txt1(51) = "" Or txt1(52) = "" Then
      '        MsgBox "日期需要正確！", vbInformation, "錯誤！"
      '        txt1(50).SetFocus
      '        txt1_GotFocus 50
      '        Exit Sub
      '    End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(50)) = "" Or Trim(txt1(51)) = "" Or Trim(txt1(52)) = "" Then
             If MsgBox("契約書日期不完整，是否確定？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
               txt1(50).SetFocus
               txt1_GotFocus 50
               Exit Sub
             End If
          End If
      '2009/11/13 END
      
          'Added by Lydia 2017/03/28
          If ChkSeal.Value = 1 Then
            'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
            'end 2017/04/27
             bolAddSeal = True
          Else
             bolAddSeal = False
          End If
          'end 2017/03/28
          
          'Modified by Lydia 2017/04/13
          'For iCount = 1 To Val(txtPCnt) 'edit by nickc 2006/09/27 2
          '    'add by nickc 2006/06/05
          '    Set Printer = Printers(Combo1.ListIndex)
          '    Screen.MousePointer = vbHourglass
          '    DoEvents
          '    StrMenu
          'Next iCount
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(False)
          Call runWordProc(False)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
          
          m_strCustCode = "" 'Added by Morgan 2012/9/11
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
          Screen.MousePointer = vbDefault
          Call RunEndProc(True)  'Added by Lydia 2022/01/20 刪除暫存檔
          'Modified by Lydia 2022/01/22 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/13
          If m_TempPDF <> "" Then ShowPrintOk
      Case 1
          frm210114.Show
          Unload Me
      Case 2
          For Each tb In txt1
              tb.Text = Empty
          Next
          For Each op In opt1
              op.Value = False
          Next
      Case 3
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowSave
          If cd1.FileName <> "" Then
              AllObj(0) = "案件委任契約書-CFP"
              'Modify By Sindy 2010/3/17
              'For iCount = 1 To 53
              For iCount = 1 To 56
                  AllObj(iCount) = txt1(iCount - 1).Text
              Next iCount
              'Modify By Sindy 2010/3/17
              'AllObj(54) = IIf(opt1(0).Value = True, "0", "1")
              'AllObj(55) = IIf(opt1(1).Value = True, "0", "1")
              AllObj(57) = IIf(opt1(0).Value = True, "0", "1")
              AllObj(58) = IIf(opt1(1).Value = True, "0", "1")
              AllObj(59) = Combo2.Text 'Add By Sindy 2010/9/9
              strBuffer = Join(AllObj, Chr(30))
              strBuffer = StrEncrypt(strBuffer)
              fN = FreeFile
              Open cd1.FileName For Output As fN
              Print #fN, strBuffer
              Close #fN
          End If
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2 <> MsgText(601) And Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
      Case 4
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowOpen
          If cd1.FileName <> "" Then
              fN = FreeFile
              Open cd1.FileName For Input As fN
              Input #fN, strBuffer
              Close #fN
              strBuffer = StrDecrypt(strBuffer)
              AllObjV = Split(strBuffer, Chr(30))
              If AllObjV(0) = "案件委任契約書-CFP" Then
                  cmdOK_Click 2
                  'Modify By Sindy 2010/3/17
                  'For iCount = 1 To 53
                  For iCount = 1 To 56
                       txt1(iCount - 1).Text = AllObjV(iCount)
                  Next iCount
                  'Modify By Sindy 2010/3/17
                  'opt1(0).Value = IIf(Val(AllObjV(54)) = 0, True, False)
                  'opt1(1).Value = IIf(Val(AllObjV(55)) = 0, True, False)
                  opt1(0).Value = IIf(Val(AllObjV(57)) = 0, True, False)
                  opt1(1).Value = IIf(Val(AllObjV(58)) = 0, True, False)
                  'Modify by Amy 2016/08/19 避免空值會Error
                  If AllObjV(59) = MsgText(601) Then
                    Combo2.ListIndex = 0
                  Else
                    Combo2.Text = AllObjV(59) 'Add By Sindy 2010/9/9
                  End If
                  'end 2016/08/19
                  
                  'Add By Sindy 2011/1/21 檢查地址欄
                  '申請人地址(中)
                  If txt1(5).Text <> "" And txt1(7).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(5).Text), Trim(txt1(7).Text), "申請人中文", True) = False Then
                        txt1(7).SetFocus
                     End If
                  End If
                  '申請人地址(英)
                  If txt1(6).Text <> "" And txt1(8).Text <> "" Then
                     If CheckCustomerAddr(2, Trim(txt1(6).Text), Trim(txt1(8).Text), "申請人英文", True) = False Then
                        txt1(8).SetFocus
                     End If
                  End If
                  '委任人地址
                  If txt1(44).Text <> "" And txt1(46).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(44).Text), Trim(txt1(46).Text), "委任人", True) = False Then
                        txt1(46).SetFocus
                     End If
                  End If
                  '2011/1/21 End
                  'Add by Amy 2016/08/19 讀取收據公司別
                  cmdFind_Click
              Else
                  MsgBox "錯誤格式，此份內容並非 CFP 格式！", vbExclamation
              End If
          End If
      'Added by Lydia 2017/03/24 空白委任書
      Case 5
          If Trim(Combo2) = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          'Modified by Lydia 2017/04/17 文雄表示用印由下方勾選,可直接空白列印
          If ChkSeal.Value = 1 Then
            If (InStr(UCase(Combo1.Text), "BATCH") > 0 Or InStr(UCase(Combo1.Text), "WRITER") > 0 Or InStr(UCase(Combo1.Text), "PDF") > 0) And Pub_StrUserSt03 <> "M51" Then
               MsgBox "空白用印的印表機不可選擇PDF列印！", vbInformation, "錯誤！"
               Combo1.SetFocus
               Exit Sub
            End If
            'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
            'end 2017/04/27
            bolAddSeal = True
          End If
          'end 2017/04/17
          Call cmdOK_Click(2) '清空資料
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(True)
          Call runWordProc(True)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
          
          m_strCustCode = ""
          bolAddSeal = False
          Screen.MousePointer = vbDefault
          Call RunEndProc(True)  'Added by Lydia 2022/01/20 刪除暫存檔
          'Modified by Lydia 2022/01/22 判斷是否有列印
          'ShowPrintOk
          If m_TempPDF <> "" Then ShowPrintOk
      'end 2017/03/24
      Case Else
   End Select
   Exit Sub
DialogCancel:
End Sub

'Add by Morgan 2011/2/24 只要鍵盤有動作就不斷線
Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Forms(0).Name) = "MDIMAIN" Then Forms(0).tmrConnect.Tag = 0
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   PUB_InitForm210114 Forms(0), Me 'Added by Lydia 2017/05/19 委任契約書表單大於主表單，控制主表單放大。
   MoveFormToCenter Me
   'Modified by Lydia 2017/04/28 改用模組
   'strSql = Printer.DeviceName
   'SeekPrintL = Printer.Orientation
   'For i = 0 To Printers.Count - 1
   '    Set Printer = Printers(i)
   '    Combo1.AddItem Printer.DeviceName, j
   '    j = j + 1
   '    If Printer.DeviceName = strSql Then
   '        SeekPrint = i
   '    End If
   'Next i
   'Set Printer = Printers(SeekPrint)
   'Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Me.Combo1, strPrinter, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數
   
    'Added by Lydia 2017/04/17 先用模組抓所有印表機後,排除特定印表機
    'Remove by Lydia 2017/06/07 改直接列印
    'For i = 0 To Combo1.ListCount - 1
    '   If InStr(UCase(Combo1.List(i)), "PDFCREATOR") > 0 And Trim(Combo1.List(i)) <> "" Then
    '      Combo1.RemoveItem i
    '      'If i = SeekPrint Then Combo1.Text = Combo1.List(0) 'Remove by Lydia 2017/04/28
    '   End If
    'Next
    'end 2017/04/17
    'end 2017/06/07
    
   'Add By Sindy 2013/12/15
   'Remove by Lydia 2020/03/25
   'If strSrvDate(1) >= InvoiceStartDate Then
   '   Combo2.AddItem "台一智權股份有限公司"
   'End If
   ''2013/12/15 END
   'end 2020/03/25
   
   'Added by Lydia 2020/03/25 設定公司別下拉選項
   Call PUB_SetCboTofrm210114(Me.Name, Me.Combo2, strCompSeal)
   
   'Modify by Amy 2016/08/19
   'Combo2.Text = Combo2.List(0)
   Combo2.ListIndex = 0
   
   'Added by Lydia 2017/03/28 預設項目
   txt1(53).Text = "申請費"
   txt1(54).Text = "要求審查費"

End Sub

Private Sub Form_Unload(Cancel As Integer)
   '還原預設印表機
   'Modified by Lydia 2017/04/28 記錄表單的印表機
   'Set Printer = Printers(SeekPrint)
   'Printer.Orientation = SeekPrintL
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   'end 2017/04/28
   
   Call RunEndProc(False)  'Added by Lydia 2022/01/20 刪除暫存檔
   
   Set frm210114_2 = Nothing
End Sub

'Modified by Lydia 2017/03/28
'Sub StrMenu()
Sub StrMenu(Optional ByVal bolSpace As Boolean = False)
Dim iY As Integer
Dim tmpI As Integer
Dim iStrL(1 To 47) As String
Dim iStrR(1 To 47) As String
Dim tBoxTop As Integer
'Added by Lydia 2017/03/28
Dim tObj As New StdPicture
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
'end 2017/03/28

   '??■□
   'Modified by Lydia 2017/03/27 變更格式
   'iStrL(1) = "國外專利案件委任契約書"
   'iStrL(2) = "委任人（甲方）茲委任受任人（乙方）辦理國外專利案件，雙方同意條件如下："
   iStrL(1) = "專利案件委任契約書"
   iStrL(2) = "委任人（甲方）茲委任受任人（乙方）辦理專利案件，雙方同意條件如下："
   'end 2017/03/27
   iStrL(3) = "第一條  專利之名稱：" & StrToStr(txt1(0) & String(74, " "), 37)
   iStrL(4) = ""
   iStrL(5) = "　　　　　　　（中）" & StrToStr(txt1(1) & String(74, " "), 37)
   iStrL(6) = "　　發明人姓名："
   iStrL(7) = "　　　　　　　（英）" & StrToStr(txt1(2) & String(74, " "), 37)
   iStrL(8) = ""
   iStrL(9) = "　　　　　　　（中）" & StrToStr(txt1(3) & String(74, " "), 37)
   iStrL(10) = "　　發明人地址："
   iStrL(11) = "　　　　　　　（英）" & StrToStr(txt1(4) & String(74, " "), 37)
   iStrL(12) = ""
   iStrL(13) = "　　　　　　　（中）" & StrToStr(txt1(5) & String(74, " "), 37)
   iStrL(14) = "　　申請人姓名："
   iStrL(15) = "　　　　　　　（英）" & StrToStr(txt1(6) & String(74, " "), 37)
   iStrL(16) = ""
   iStrL(17) = "　　　　　　　（中）" & StrToStr(txt1(7) & String(74, " "), 37)
   iStrL(18) = "　　申請人地址："
   iStrL(19) = "　　　　　　　（英）" & StrToStr(txt1(8) & String(74, " "), 37)
   iStrL(20) = "　　指定聯絡人"
   iStrL(21) = "　　及通訊地址：" & StrToStr(txt1(9) & String(78, " "), 39)
   iStrL(22) = "第二條　委辦範圍"
   'Added by Lydia 2017/03/28 有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
      For intI = 1 To 22
         If Trim(iStrL(intI)) <> "" Then
            strDetail = strDetail & RTrim(iStrL(intI)) & vbCrLf
         End If
      Next
   End If
   'end 2017/03/28
   
   'Modified by Lydia 2017/03/27 變更格式
   'iStrL(23) = "一、提出申請程序：乙方根據甲方所提供之發明創作資料或樣品，代撰專利說明書及圖式並為甲方委請各該"
   'iStrL(24) = "　　　　　　　　　申請國之專利代理人，代向各該國提出專利申請。"
   iStrL(23) = "一、提出申請程序：乙方根據甲方所提供之發明創作資料或樣品，代撰專利說明書及圖式並向我國專利主管"
   iStrL(24) = "　　　　　　　　　機關提出申請，或為甲方委請各該申請國之專利代理人，代向各該國提出專利申請。"
   'end 2017/03/27
   iStrL(25) = "二、中間處理程序：提出專利申請後依甲方請求或各該申請國專利主管機關之指示所需提出之修正、更正、"
   iStrL(26) = "　　　　　　　　　補充說明書之各程序。"
   iStrL(27) = "三、領證及繳納年費程序：依各該申請國專利法規之規定在申請程序中繳納維持費或審定核准後繳納證書費"
   iStrL(28) = "　　　　　　　　　、年費、維持費，領取證書。"
   iStrL(29) = "四、救 濟 程 序 ：審定不予專利時，所需進行之救濟程序，如請求再審、提起審判，以及答辯等程序。"
   iStrL(30) = "五、讓渡 或授權 ：乙方提供各該國之空白讓渡書或授權書供甲方簽署，或根據甲方所提供之讓渡書或授權"
   'Modified by Lydia 2017/03/27 變更格式
   'iStrL(31) = "　　　　　　　　　書，委請各該國專利代理人，代向該國專利主管機關辦理登記。"
   iStrL(31) = "　　　　　　　　　書，自行或委請各該國專利代理人，代向該國專利主管機關辦理登記。"
   'end 2017/03/27
   iStrL(32) = "六、專 利 調 卷 ：乙方根據甲方所提供之資料進行調卷。"
   'Modified by Lydia 2017/03/27 變更格式
   'iStrL(33) = "七、專 利 調 查 ：乙方根據甲方所提供之資料，委請該國代理人進行專利調查，此專利調查並不含專利調"
   'iStrL(34) = "　　　　　　　　　卷。"
   iStrL(33) = "七、專 利 調 查 ：乙方根據甲方所提供之資料，自行或委請該國代理人進行專利調查，此專利調查並不含"
   iStrL(34) = "　　　　　　　　　專利調卷。"
   'end 2017/03/27
   iStrL(35) = "八、其        他：如說明書、圖式之更正、新穎性調查、外國專利主管機關文件引證資料之翻譯、公告或"
   iStrL(36) = "　　　　　　　　　延期公告之申請等程序。"
   iStrL(37) = "第三條 委辦費用"
   'Modified by Lydia 2017/03/27 變更格式
   'iStrL(38) = "　　一、乙方受委辦前條第" & String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text) & String(6 - LenB(StrConv(String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text), vbFromUnicode)), " ") & "款" & String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text) & String(10 - LenB(StrConv(String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text), vbFromUnicode)), " ") & "程序之費用（包括國外代理人費用），約定如下："
   'Modified by Lydia 2023/02/17 直接代入
   'iStrL(38) = "　　一、乙方受委辦前條第" & String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text) & String(6 - LenB(StrConv(String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text), vbFromUnicode)), " ") & "款" & String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text) & String(10 - LenB(StrConv(String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text), vbFromUnicode)), " ") & "程序之費用（申請國外專利時，包括國外代理人費用），約定如下："
   'end 2017/03/27
   iStrL(38) = "　　一、乙方受委辦前條第 " & Trim(txt1(10).Text) & " 款 " & Trim(txt1(11).Text) & " 程序之費用（申請國外專利時，包括國外代理人費用），約定如下："
   iStrL(39) = "　　　　　　　　　　　　　　　　　　　　　　金　　　　　　　　　　　　　額　　　　　　　　　　　　"
   iStrL(40) = "　　　國     別   　　專利種類　　　　　　　　　　　　　　　　　　　　　　　　　　　 備　　　註"
   iStrL(41) = "　  　　　　　　　　　　　　　　　" & Pub_StrToCenter(txt1(53).Text, 16) & Pub_StrToCenter(txt1(54).Text, 16) & Pub_StrToCenter(txt1(12).Text, 16)
   iStrL(42) = "　" & Pub_StrToCenter(txt1(13).Text, 16) & Pub_StrToCenter(txt1(14).Text, 16) & Pub_StrToCenter(txt1(15).Text, 16) & Pub_StrToCenter(txt1(16).Text, 16) & Pub_StrToCenter(txt1(17).Text, 16) & Pub_StrToCenter(txt1(18).Text, 16)
   iStrL(43) = "　" & Pub_StrToCenter(txt1(19).Text, 16) & Pub_StrToCenter(txt1(20).Text, 16) & Pub_StrToCenter(txt1(21).Text, 16) & Pub_StrToCenter(txt1(22).Text, 16) & Pub_StrToCenter(txt1(23).Text, 16) & Pub_StrToCenter(txt1(24).Text, 16)
   iStrL(44) = "　" & Pub_StrToCenter(txt1(25).Text, 16) & Pub_StrToCenter(txt1(26).Text, 16) & Pub_StrToCenter(txt1(27).Text, 16) & Pub_StrToCenter(txt1(28).Text, 16) & Pub_StrToCenter(txt1(29).Text, 16) & Pub_StrToCenter(txt1(30).Text, 16)
   iStrL(45) = "　" & Pub_StrToCenter(txt1(31).Text, 16) & Pub_StrToCenter(txt1(32).Text, 16) & Pub_StrToCenter(txt1(33).Text, 16) & Pub_StrToCenter(txt1(34).Text, 16) & Pub_StrToCenter(txt1(35).Text, 16) & Pub_StrToCenter(txt1(36).Text, 16)
   iStrL(46) = "　" & Pub_StrToCenter(txt1(37).Text, 16) & Pub_StrToCenter(txt1(38).Text, 16) & Pub_StrToCenter(txt1(39).Text, 16) & Pub_StrToCenter(txt1(40).Text, 16) & Pub_StrToCenter(txt1(41).Text, 16) & Pub_StrToCenter(txt1(42).Text, 16)
   'Modified by Lydia 2017/03/28 幣別改可輸入
   'iStrL(47) = "合計新台幣　" & String(24, " ") & "元整，於本契約簽訂同時由甲方一次付清。"
   If Val(Trim(txt1(43))) = 0 Then
       strSpaceAmt = String(12, "　")
   Else
       'Modified by Lydia 2023/08/10 改變數控制
       'strSpaceAmt = Replace(ChangeNumber(txt1(43)), "元整", "")
       strSpaceAmt = ChangeNumber(txt1(43), False)
   End If
   iStrL(47) = "合計" & IIf(Trim(txt1(55).Text) = "", "　　　", Trim(txt1(55).Text)) & "　" & String(24, " ") & "元整，於本契約簽訂同時由甲方一次付清。"
   'end 2017/03/28
   
   'Added by Lydia 2017/03/28 有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
      strExc(1) = ""
      For intI = 13 To 42
          '有無費用項目
          strExc(1) = strExc(1) & Trim(Replace(Replace(txt1(intI), "　", ""), " ", ""))
      Next
      For intI = 37 To 46
          If Trim(Replace(Replace(iStrL(intI), "　", ""), " ", "")) <> "" Then
             Select Case intI
                 Case 40
                     strExc(1) = "　　　國     別   　　專利種類 " & Pub_StrToCenter(txt1(53).Text, 16) & Pub_StrToCenter(txt1(54).Text, 16) & Pub_StrToCenter(txt1(12).Text, 16) & " 備　　　註"
                     strDetail = strDetail & RTrim(strExc(1)) & vbCrLf
                 Case 41
                 Case Else
                      strDetail = strDetail & RTrim(iStrL(intI)) & vbCrLf
             End Select
          End If
      Next
      'Modified by Lydia 2023/08/10 增加判斷
      'strDetail = strDetail & "合計" & Trim(txt1(55).Text) & "　" & strSpaceAmt & "元整，於本契約簽訂同時由甲方一次付清。" & vbCrLf
      strDetail = strDetail & "合計" & Trim(txt1(55).Text) & "　" & strSpaceAmt & IIf(InStr(strSpaceAmt, "元") > 0, "整", "元整") & "，於本契約簽訂同時由甲方一次付清。" & vbCrLf
   End If
   'end 2017/03/28
   
'Modified by Lydia 2021/04/19 因為第五條內文增加，所以版面調整
'   iStrR(1) = ""
'   iStrR(2) = ""
'   iStrR(3) = ""
'   'Modified by Lydia 2017/03/27 變更格式
'   'iStrR(4) = "　　二、第二條所列之委辦範圍除本條第一款特予載明外，其費用由甲方負擔，其金額依當時外國代理人費"
'   'iStrR(5) = "　　　　用及本所服務費標準收取之。"
'   iStrR(4) = "　　二、第二條所列之委辦範圍除本條第一款特予載明外，其費用由甲方負擔，申請國外專利時，其金額依"
'   iStrR(5) = "　　　　當時外國代理人費用及本所服務費標準收取之。"
'   'end 2017/03/27
'   iStrR(6) = "　　三、案件之進行中，如需乙方派員前往現場研究或繪圖時，其出差旅費，由甲方負擔，並按實際處理時"
'   iStrR(7) = "　　　　間每小時新台幣貳仟元整計算另收費用。"
'   iStrR(8) = "　　　　本條所約定之費用如甲方未於乙方所指定之期限內付清，則乙方無義務辦理所受任之事項，且經乙"
'   'Modified by Lydia 2017/03/27 變更格式
'   'iStrR(9) = "　　　　方限期催告後，如甲方仍不履行時，則本契約當然終止，乙方並得通知各該外國代理人終止進行該"
'   'iStrR(10) = "　　　　程序及嗣後之一切程序。另乙方已先行代辦之服務費用，甲方仍應照付。"
'   iStrR(9) = "　　　　方限期催告後，如甲方仍不履行時，則本契約當然終止，乙方得終止進行該程序及嗣後之一切程序"
'   iStrR(10) = "　　　　，並通知各該外國代理人終止進行所有相關程序。另乙方已先行代辦之服務費用，甲方仍應照付。"
'   'end 2017/03/27
'   iStrR(11) = "第四條  乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方權益之"
'   iStrR(12) = "　　　　疏誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額的三倍為限。"
'   iStrR(13) = "第五條  甲方確保所交付予乙方之資料，均無虛偽情事，如因不實致生損害或法律責任時，概由甲方負責，"
'   iStrR(14) = "　　　　與乙方無關。"
'---------------------------------------------------
   iStrR(1) = ""
   iStrR(2) = ""
   iStrR(3) = "　　二、第二條所列之委辦範圍除本條第一款特予載明外，其費用由甲方負擔，申請國外專利時，其金額依"
   iStrR(4) = "　　　　當時外國代理人費用及本所服務費標準收取之。"
   iStrR(5) = "　　三、案件之進行中，如需乙方派員前往現場研究或繪圖時，其出差旅費，由甲方負擔，並按實際處理時"
   iStrR(6) = "　　　　間每小時新台幣貳仟元整計算另收費用。"
   iStrR(7) = "　　　　本條所約定之費用如甲方未於乙方所指定之期限內付清，則乙方無義務辦理所受任之事項，且經乙"
   iStrR(8) = "　　　　方限期催告後，如甲方仍不履行時，則本契約當然終止，乙方得終止進行該程序及嗣後之一切程序"
   iStrR(9) = "　　　　，並通知各該外國代理人終止進行所有相關程序。另乙方已先行代辦之服務費用，甲方仍應照付。"
   iStrR(10) = "第四條  乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方權益之"
   iStrR(11) = "　　　　疏誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額的三倍為限。"
   iStrR(12) = "第五條  甲方應確保所交付予乙方之資料及本契約書所載內容(包括發明人或創作人、申請人等資訊)均無虛"
   iStrR(13) = "　　　　偽情事，且甲方確實得到與委辦案件相關共同發明人及第三人之同意，有權委託乙方辦理案件，如"
   iStrR(14) = "　　　　因不實致生損害或法律責任時，概由甲方負責，與乙方無關。"
'end 2021/04/19
   iStrR(15) = "第六條  乙方於辦理過程中，應隨時將辦理經過如申請日、案號及其他重要函件，儘速通知或交付甲方。但"
   iStrR(16) = "　　　　甲方簽約後變更連絡處所，未即時通知乙方，因而連絡不及致誤時限者，乙方不負責任。"
   iStrR(17) = "第七條  凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責任。經乙方通知"
   iStrR(18) = "　　　　甲方繳費而未依限繳納者，亦同。"
   iStrR(19) = "第八條  甲方如逕自撤回所委辦之程序，或未經乙方同意終止契約時，所約定之費用，仍應全數給付。"
   iStrR(20) = "第九條  本約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需甲乙雙方於更動處"
   iStrR(21) = "　　　　蓋章始生效力，並由雙方各執乙份為憑。"
   iStrR(22) = ""
   iStrR(23) = "附  則  乙方所撰寫之專利說明書、圖式、再審文件或修正書，是否需要會稿，請於下方方格註明。"
   iStrR(24) = ""
   iStrR(25) = ""
   iStrR(26) = " 　   　　     會稿　　　　 " & IIf(opt1(0).Value = True, "Ｖ", "　") & "　　　　不會稿　　　　 " & IIf(opt1(1).Value = True, "Ｖ", "　")
   iStrR(27) = ""
   iStrR(28) = ""
   iStrR(29) = "　　甲方：委任人：" & StrToStr(txt1(44) & String(57, " "), 28)
   iStrR(30) = ""
   iStrR(31) = "　　　　　代表人：" & StrToStr(txt1(45) & String(57, " "), 28)
   iStrR(32) = ""
   iStrR(33) = "　　      地  址：" & StrToStr(txt1(46) & String(57, " "), 28)
   iStrR(34) = ""
   iStrR(35) = "　　　　　電  話：" & StrToStr(txt1(47) & String(28, " "), 14) & "傳  真：" & StrToStr(txt1(48) & String(26, " "), 13)
   iStrR(36) = ""
   iStrR(37) = ""
   iStrR(38) = "　　乙方：受任人：" & Combo2.Text '"台一國際專利法律事務所                               "
   iStrR(39) = ""
   iStrR(40) = "　　　　　經手人：" & StrToStr(txt1(49) & String(57, " "), 28)
   iStrR(41) = ""
   'Add By Sindy 2013/12/15
   'Modified by Lydia 2020/04/09 改用模組控制
   'If Combo2 = "台一智權股份有限公司" Then
   '   iStrR(42) = "　　　　　地　址：台北市長安東路二段一一０號四樓"
   'Else
   ''2013/12/15 END
   '   iStrR(42) = "　　　　　地　址：台北市長安東路二段一一二號九樓"
   'End If
   iStrR(42) = "　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   'end 2020/04/09
   iStrR(43) = ""
   iStrR(44) = "　　　　　電  話：(02)25061023(總機)  　　　　F A X ：(02)25011666"
   iStrR(45) = ""
   'Add By Sindy 2013/12/15
   If Combo2 = "台一智權股份有限公司" Then
      iStrR(46) = ""
   Else
   '2013/12/15 END
      iStrR(46) = "　　　　　網  址：www.taie.com.tw     　　　　E-mail：ipdept@taie.com.tw"   'modify by sonia 2020/4/8 原為lawoffice
   End If
   'Modified by Lydia 2017/04/25 縮短長度
   'iStrR(47) = " 中    華    民    國 " & String((10 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & txt1(50) & String((10 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & txt1(51) & String((10 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & "月" & String((8 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & txt1(52) & String((8 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & "日"
   iStrR(47) = " 中  華  民  國 " & String((10 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & txt1(50) & String((10 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & txt1(51) & String((10 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & "月" & String((8 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & txt1(52) & String((8 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & "日"
   
   'Added by Lydia 2017/03/28 有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
      strDetail = strDetail & vbCrLf & IIf(opt1(0).Value = True, "會稿", "不會稿") & vbCrLf & vbCrLf
      For intI = 29 To 40
         If Trim(iStrR(intI)) <> "" Then
            strDetail = strDetail & RTrim(iStrR(intI)) & vbCrLf
         End If
      Next
      strDetail = strDetail & iStrR(47)
      'Modified by Lydia 2017/04/17 空白用印改由勾選項目控制
      'If PUB_AddRecSeal("2", txtPCnt.Text, IIf(ChkSeal.Value = 1, "", "Y"), strDetail, Combo2.Text) Then
      'Remove by Lydia 2017/05/16 用印記錄移到pdf建立
      'If PUB_AddRecSeal("2", txtPCnt.Text, IIf(bolSpace = True, "Y", ""), strDetail, Combo2.Text) Then
      'End If
   End If
   'end 2017/03/28
        
   'Modified by Lydia 2017/04/27 實際列印的上邊界
   'iY = 0
   iY = d_Top
   
   Printer.FontBold = True
   Printer.PaperSize = 9

   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = 1 Then
       Printer.Orientation = 2
   'End If
   Printer.FontName = "標楷體"
   Printer.FontSize = 14
   'Modified by Lydia 2017/04/25 + d_Left(實際列印的左邊界)
   Printer.CurrentX = ((Printer.ScaleWidth / 2) - Printer.TextWidth(iStrL(1))) / 2 + d_Left
   iY = iY + Printer.TextHeight(iStrL(1))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStrL(1)) / 3) * 4)
   Printer.Print iStrL(1)
   Printer.FontSize = 11
   'Modified by Lydia 2017/04/25 + d_Left
   Printer.CurrentX = Printer.TextWidth("　") + d_Left
   iY = iY + Printer.TextHeight(iStrL(2))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStrL(2)) / 3) * 4)
   Printer.Print iStrL(2)
   Printer.FontSize = 8
   'Added by Lydia 2017/03/28 同步用印
   If bolAddSeal = True Then
      '列印座置抓乙方資料的起始
      'X軸
      'Modified by Lydia 2017/04/25 + d_Left
      strExc(1) = Printer.ScaleWidth / 2 + (Printer.TextWidth("　") * 37) + d_Left
      'Y軸
      strExc(2) = iY + ((Printer.TextHeight("　") / 3) * 4) * 36
      'Added by Lydia 2017/04/25 圖片尺寸
      strExc(3) = 1600 'width
      strExc(4) = 1600 'height
         
      'Added by Lydia 2020/03/25 已記錄公司名稱|用印編號
      intI = InStr(strCompSeal, Combo2.Text)
      If intI > 0 Then
         strExc(9) = Mid(strCompSeal, intI + Len(Combo2.Text))
         If InStr(strExc(9), ",") > 0 Then
             strExc(9) = Mid(strExc(9), 2, InStr(strExc(9), ",") - 2)
         Else
             strExc(9) = Mid(strExc(9), 2)
         End If
          If PUB_ReadDB2File(strSealFile, Val(strExc(9))) Then
             Set tObj = pvGetStdPicture(strSealFile)
             Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
          End If
      Else
      'end 2020/03/25
            If InStr(Combo2.Text, "專利法律") > 0 Then
              If PUB_ReadDB2File(strSealFile, 51) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
            If InStr(Combo2.Text, "專利商標") > 0 Then
              If PUB_ReadDB2File(strSealFile, 52) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
            If InStr(Combo2.Text, "台一智權") > 0 Then
              If PUB_ReadDB2File(strSealFile, 53) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
      End If 'Added by Lydia 2020/03/25
   End If
   'end 2017/03/28
   For tmpI = 3 To UBound(iStrL) - 1
       If iStrL(tmpI) & iStrR(tmpI) <> "" Then
           If tmpI = 35 Then
               tBoxTop = iY
           End If
           '畫格子
           Select Case tmpI
           Case 25
               'Modified by Lydia 2017/04/25 + d_Left
               'Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8), iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 38), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               'Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8), iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 28), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               'Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8), iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 23), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               'Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8), iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 13), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8) + d_Left, iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 38) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8) + d_Left, iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 28) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8) + d_Left, iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 23) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 8) + d_Left, iY)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 13) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
              
           Case 39
               'Modified by Lydia 2017/04/25 + d_Left
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 50), iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 50), iY + (((Printer.TextHeight("　") / 3) * 4) * 7) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 50), iY + (((Printer.TextHeight("　") / 3) * 4) * 6) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 50), iY + (((Printer.TextHeight("　") / 3) * 4) * 5) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 50), iY + (((Printer.TextHeight("　") / 3) * 4) * 4) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 50), iY + (((Printer.TextHeight("　") / 3) * 4) * 3) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 42), iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 18), iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 2), iY)-((Printer.TextWidth("　") * 10), iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
'
'               Printer.Line ((Printer.TextWidth("　") * 18), iY + (((Printer.TextHeight("　") / 3) * 4) * 1.6) - 50)-((Printer.TextWidth("　") * 42), iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 18), iY + (((Printer.TextHeight("　") / 3) * 4) * 1.6) - 50)-((Printer.TextWidth("　") * 34), iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
'               Printer.Line ((Printer.TextWidth("　") * 18), iY + (((Printer.TextHeight("　") / 3) * 4) * 1.6) - 50)-((Printer.TextWidth("　") * 26), iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 50) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 50) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 7) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 50) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 6) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 50) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 5) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 50) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 4) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 50) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 3) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 42) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 18) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 2) + d_Left, iY)-((Printer.TextWidth("　") * 10) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
               
               Printer.Line ((Printer.TextWidth("　") * 18) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 1.6) - 50)-((Printer.TextWidth("　") * 42) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 18) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 1.6) - 50)-((Printer.TextWidth("　") * 34) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B
               Printer.Line ((Printer.TextWidth("　") * 18) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 1.6) - 50)-((Printer.TextWidth("　") * 26) + d_Left, iY + (((Printer.TextHeight("　") / 3) * 4) * 8) - 50), , B

           Case Else
           End Select
           If tmpI = 39 Then
               'Modified by Lydia 2017/04/25 + d_Left
               Printer.CurrentX = Printer.TextWidth("　") + d_Left
               Printer.CurrentY = iY + (Printer.TextHeight("　") * 0.4)
               Printer.Print iStrL(tmpI)
           ElseIf tmpI = 41 Then
               'Modified by Lydia 2017/04/25 + d_Left
               Printer.CurrentX = Printer.TextWidth("　") + d_Left
               Printer.CurrentY = iY - (Printer.TextHeight("　") * 0.4)
               Printer.Print iStrL(tmpI)
           Else
               'Modified by Lydia 2017/04/25 + d_Left
               Printer.CurrentX = Printer.TextWidth("　") + d_Left
               Printer.CurrentY = iY
               Printer.Print iStrL(tmpI)
           End If
           If tmpI >= 26 Then
               Printer.FontSize = 10
   '            Printer.FontBold = True
               'Modified by Lydia 2017/04/25 + d_Left
               Printer.CurrentX = Printer.ScaleWidth / 2 + d_Left
               Printer.CurrentY = iY
               Printer.Print iStrR(tmpI)
               Printer.FontSize = 8
   '            Printer.FontBold = False
           Else
               'Modified by Lydia 2017/04/25 + d_Left
               Printer.CurrentX = Printer.ScaleWidth / 2 + d_Left
               Printer.CurrentY = iY
               Printer.Print iStrR(tmpI)
           End If
           iY = iY + ((Printer.TextHeight(iStrL(tmpI)) / 3) * 4)
           '畫線
           Select Case tmpI
           Case 3, 5, 7, 9, 11, 13, 15, 17, 19
                'Modified by Lydia 2017/04/25 + d_Left
                'Printer.Line (Printer.TextWidth("　") + (Printer.TextWidth("　") * 10), iY + 50)-(Printer.TextWidth("　") + (Printer.TextWidth("　") * 47), iY + 50)
                Printer.Line (Printer.TextWidth("　") + (Printer.TextWidth("　") * 10) + d_Left, iY + 50)-(Printer.TextWidth("　") + (Printer.TextWidth("　") * 47) + d_Left, iY + 50)
           Case 21
                'Modified by Lydia 2017/04/25 + d_Left
                'Printer.Line (Printer.TextWidth("　") + (Printer.TextWidth("　") * 8), iY - 50)-(Printer.TextWidth("　") + (Printer.TextWidth("　") * 47), iY - 50)
                Printer.Line (Printer.TextWidth("　") + (Printer.TextWidth("　") * 8) + d_Left, iY - 50)-(Printer.TextWidth("　") + (Printer.TextWidth("　") * 47) + d_Left, iY - 50)
           Case 29, 31, 33, 38, 40, 42
                'Modified by Lydia 2017/04/25 + d_Left
                'Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 11.5), iY + 50)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 47), iY + 50)
                Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 11.5) + d_Left, iY + 50)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 47) + d_Left, iY + 50)
           Case 35, 44, 46
               'Add By Sindy 2013/12/15
               If tmpI = 46 And Combo2 = "台一智權股份有限公司" Then
                  '不用畫線
               Else
               '2013/12/15 END
                  'Modified by Lydia 2017/04/25 + d_Left
                  'Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 11.5), iY + 50)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 28.5), iY + 50)
                  'Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 34), iY + 50)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 47), iY + 50)
                  Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 11.5) + d_Left, iY + 50)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 28.5) + d_Left, iY + 50)
                  Printer.Line ((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 34) + d_Left, iY + 50)-((Printer.ScaleWidth / 2) + (Printer.TextWidth("　") * 47) + d_Left, iY + 50)
               End If
           Case Else
           End Select
       End If
   Next tmpI
   Printer.FontSize = 8
   iY = iY + ((Printer.TextHeight(iStrL(tmpI)) / 3) * 4)
   'Modified by Lydia 2017/04/25 + d_Left
   Printer.CurrentX = Printer.TextWidth("　") + d_Left
   Printer.CurrentY = iY
   Printer.Print iStrL(47)
   'Printer.FontBold = True
   'Modified by Lydia 2017/03/28
   'Printer.CurrentX = (Printer.TextWidth("　") * 18) - Printer.TextWidth(Replace(ChangeNumber(txt1(43)), "元整", ""))
   'Modified by Lydia 2017/04/25 + d_Left
   Printer.CurrentX = (Printer.TextWidth("　") * 18) - Printer.TextWidth(strSpaceAmt) + d_Left
   Printer.CurrentY = iY
   'Modified by Lydia 2017/03/28
   'Printer.Print Replace(ChangeNumber(txt1(43)), "元整", "")
   Printer.Print strSpaceAmt
   'Printer.FontBold = False
   Printer.FontSize = 14
   'Memo by Lydia 2017/04/25 民國年月日的位置
   'Printer.CurrentX = (((Printer.ScaleWidth / 2) - Printer.TextWidth(iStrR(UBound(iStrR)))) / 2) + (Printer.ScaleWidth / 2)
   Printer.CurrentX = (Printer.ScaleWidth / 2) + Printer.TextWidth("　") + d_Left
   Printer.CurrentY = iY - 50
   Printer.Print iStrR(UBound(iStrR))
   Printer.FontSize = 8
   iY = iY + ((Printer.TextHeight(iStrL(tmpI)) / 3) * 4)
   'Modified by Lydia 2017/04/25 + d_Left
   'Printer.Line (Printer.TextWidth("　") + (Printer.TextWidth("　") * 5), iY - 50)-(Printer.TextWidth("　") + (Printer.TextWidth("　") * 17), iY - 50)
   Printer.Line (Printer.TextWidth("　") + (Printer.TextWidth("　") * 5) + d_Left, iY - 50)-(Printer.TextWidth("　") + (Printer.TextWidth("　") * 17) + d_Left, iY - 50)
   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = Val(txtPCnt) Then
       Printer.EndDoc
   'Else
   '    Printer.NewPage
   'End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

'Modified by Lydia 2022/01/18 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
'Add By Sindy 98/02/11
Dim intLen As Integer
   
   If KeyAscii <> 8 Then
      intLen = GetTextLength(txt1(Index))
      intLen = intLen + GetTextLength(Chr(KeyAscii))
      '2014/5/13 modify by sonia
      'If intLen > txt1(Index).MaxLength Then KeyAscii = 0
      If CheckLengthIsOK(txt1(Index).Text & Chr(KeyAscii), txt1(Index).MaxLength) = False Then
         KeyAscii = 0
      End If
      'end 2014/5/13
   End If
   '98/02/11 End
   If Index = 15 Or Index = 16 Or Index = 17 Or Index = 21 Or Index = 22 Or Index = 23 Or Index = 27 Or Index = 28 Or Index = 29 Or Index = 33 Or Index = 34 Or Index = 35 Or Index = 39 Or Index = 40 Or Index = 41 Or Index = 43 Then
       If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
           KeyAscii = 0
       End If
   End If
   '2009/11/13 ADD BY SONIA
   If Index = 50 Or Index = 51 Or Index = 52 Then
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
      End If
   End If
   '2009/11/13 END
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   'Modified by Lydia 2018/04/13
   'txt1(Index).Text = Replace(Replace(txt1(Index).Text, Chr(10), ""), Chr(13), "")
   txt1(Index).Text = PUB_StringFilter(txt1(Index).Text)
   Cancel = False
   If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
       Cancel = True
   End If
   'Add By Sindy 2013/12/16
   If strSrvDate(1) >= InvoiceStartDate Then
      If Index = 13 Or Index = 19 Or Index = 25 Or Index = 31 Or Index = 37 Then 'Add By Sindy 2014/5/6 +if
         If InStr(txt1(13), "大陸") > 0 Or InStr(txt1(19), "大陸") > 0 Or InStr(txt1(25), "大陸") > 0 _
             Or InStr(txt1(31), "大陸") > 0 Or InStr(txt1(37), "大陸") > 0 Then
            'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
            If PUB_ChkCU144isN("", "", txt1(5), IIf(Combo2 = "台一智權股份有限公司", "J", ""), False) = False Then
               Combo2.Text = Combo2.List(2)
            End If
         End If
      End If
   End If
   '2013/12/16 END
End Sub

Private Sub txtPCnt_GotFocus()
   txtPCnt.SelStart = 0
   txtPCnt.SelLength = Len(txtPCnt)
   End Sub

Private Sub txtPCnt_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
       KeyAscii = 0
   End If
End Sub

'Add By Sindy 2010/4/29
Private Function SetCustTxt(strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add by Lydia 2014/9/22
Dim rsB As New ADODB.Recordset, part1 As Integer, ppart1 As Integer, partCust As String
partCust = strCUCode

   SetCustTxt = False
   strCUCode = Left(strCUCode & "000000000", 9)
   'Modified by Morgan 2021/5/5
   'StrSQLa = "Select * From Customer,nation,potcustcont Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=substr(CU08, 1, 8) And pcc02(+)=substr(CU08, 9, 1) "
   StrSQLa = "Select * From Customer,nation,potcustcont Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=cu01 And pcc02(+)=cu127 "
   'end 2021/5/5
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      SetCustTxt = True
      '申請人中文
      Me.txt1(5).Text = "" & rsA("CU04").Value
      Me.txt1(44).Text = "" & rsA("CU04").Value
      '申請人英文
      Me.txt1(6).Text = "" & rsA("CU05").Value & "" & rsA("CU88").Value & "" & rsA("CU89").Value & "" & rsA("CU90").Value
      
     'Add by Lydia 2014/9/22 複選人員-中文名
     part1 = 1
     ppart1 = GetSubStringCount(partCust) '取得字串以逗點分隔的Sub字串總數
         Do While part1 < ppart1
            strCUCode = Mid(partCust, (part1 * 10) + 1, 9) '從第幾組代號開始，截取下一組代號
           'Modified by Morgan 2021/5/5
           'StrSQLa = " Select CU04,CU05,CU88,CU89,CU90 From Customer,nation,potcustcont " & _
                  " Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=substr(CU08, 1, 8) And pcc02(+)=substr(CU08, 9, 1) "
           StrSQLa = " Select CU04,CU05,CU88,CU89,CU90 From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
           'end 2021/5/5
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            rsB.CursorLocation = adUseClient
            rsB.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
                rsB.MoveFirst
                Me.txt1(5).Text = LTrim(RTrim(Me.txt1(5).Text)) + "、" & rsB("CU04").Value
                If Len(LTrim(RTrim(Me.txt1(6).Text))) > 0 Then
                  Me.txt1(6).Text = LTrim(RTrim(Me.txt1(6).Text)) + ", " + "" & rsB("CU05").Value & "" & rsB("CU88").Value & "" & rsB("CU89").Value & "" & rsB("CU90").Value
                Else
                  Me.txt1(6).Text = "" & rsB("CU05").Value & "" & rsB("CU88").Value & "" & rsB("CU89").Value & "" & rsB("CU90").Value
                End If
            
            End If
            
            part1 = part1 + 1
            
            If part1 = ppart1 Then
                If CheckLengthIsOK(txt1(5).Text, txt1(5).MaxLength) = False Then
                    Me.txt1(5).SetFocus
                ElseIf CheckLengthIsOK(txt1(6).Text, txt1(6).MaxLength) = False Then
                    Me.txt1(6).SetFocus
                End If
            End If

        
        Loop
           
           
'      'ID No.
'      Me.txt1(6).Text = "" & rsA("CU11").Value
      '申請地址
      Me.txt1(7).Text = "" & rsA("CU23").Value
      Me.txt1(46).Text = "" & rsA("CU23").Value
      '申請英文地址
      Me.txt1(8).Text = "" & rsA("CU24").Value & "" & rsA("CU25").Value & "" & rsA("CU26").Value & "" & rsA("CU27").Value & "" & rsA("CU28").Value
'      '國籍
'      Me.txt1(8).Text = "" & rsA("NA03").Value
      '聯絡人地址
      'Modified by Morgan 2021/5/5
      'If "" & rsA("CU08").Value <> "" Then
      If "" & rsA("pcc22").Value <> "" Then
      'end 2021/5/5
         Me.txt1(9).Text = "" & rsA("pcc22").Value
      Else
         Me.txt1(9).Text = "" & rsA("CU31").Value
      End If
      '電話1
      Me.txt1(47).Text = "" & rsA("CU16").Value
      '傳真1
      Me.txt1(48).Text = "" & rsA("CU18").Value
      '代表人1中文
      Me.txt1(45).Text = "" & rsA("CU07").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function


'Add by Lydia 2014/9/22
Private Sub cmdFind2_Click()
   If Me.txt1(1).Text = "" Then
      MsgBox "請輸入發明人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(1).SetFocus
      Exit Sub
   End If
   
   frm090801_1.m_Type = 3
   If ChkDou.Value = 1 Then
     frm090801_1.m_DouChk = True '可複選
   Else
     frm090801_1.m_DouChk = False
   End If
   Set frm090801_1.m_frm0908A = Me
   
   frm090801_1.m_strCustChnName = Me.txt1(1).Text
   frm090801_1.lblName.Caption = Me.txt1(1).Text
 
   m_blnOneRec = False
   m_strCustCode = ""
'   txt1(1).Tag = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   
   If CheckLengthIsOK(txt1(1).Text, txt1(1).MaxLength) = False Then
      Me.txt1(1).SetFocus
   ElseIf CheckLengthIsOK(txt1(2).Text, txt1(2).MaxLength) = False Then
      Me.txt1(2).SetFocus
   End If
   
   
End Sub

'Add by Amy 2016/08/19
Private Sub UpdReceiptCmp(ByVal stNowCustNo As String, ByVal stNowCmp As String)
    Dim strUpd As String
    
    'Add by Amy 2016/12/30 +同業務區或為MCTF同組人員才可回寫收據公司別
    If ChkSameCuArea(stNowCustNo, strUserNum) = False Then Exit Sub
    
    'Added by Lydia 2022/08/30 受任人若選擇台一國際智慧財產事務所時，更新客戶檔之相關欄位請更新為NULL，台一智權股份有限公司才更新為J。
    If stNowCmp <> "J" Then
        stNowCmp = ""
    End If
    'end 2022/08/30
    
    'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time(CU84,Cu85,Cu86)
    'strUpd = "Update Customer Set CU84='" & strUserNum & "',CU85=to_number(to_char(sysdate,'YYYYMMDD')),CU86=to_number(to_char(sysdate,'HH24MI')),CU161='" & stNowCmp & "'  " & _
                  "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(strNowCustNo, 9, 1) & "' "
    strUpd = "Update Customer Set CU161='" & stNowCmp & "'  " & _
                  "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(strNowCustNo, 9, 1) & "' "
    Pub_SeekTbLog strUpd
    'Modified by Lydia 2019/04/23 觸發Trigger
    'cnnConnection.Execute strUpd
    cnnConnection.Execute "begin user_data.user_enabled:=1; " & strUpd & " ; end; "
End Sub

'Added by Lydia 2017/04/13 列印:先轉PDF,列印後刪檔
Private Sub Print2PDF(ByVal bSpace As Boolean)
Dim strFileName As String
Dim strOldName As String 'Added by Lydia 2017/06/07

'Added by Lydia 2017/04/25 VB印表機實際列印的左邊界、右邊界
Set Printer = Printers(PUB_PrinterIndex(Combo1.Text))
d_Top = Format((Printer.Width - Printer.ScaleWidth) / 2, "0") '外專-橫印
d_Left = Format((Printer.Height - Printer.ScaleHeight) / 2, "0")
d_Top = 0
d_Left = 0
'end 2017/04/25

strDetail = "" 'Added by Lydia 2017/05/16
strOldName = App.Title 'Added by Lydia 2017/06/07

Screen.MousePointer = vbHourglass
'    'Modified by Lydia 2022/01/19 先產生Word檔，後轉成PDF檔逐一列印
'    For iCount = 1 To Val(txtPCnt)
'        iPrintC = iCount
'        'Modified by Lydia 2017/06/06 改用App.Title變更印表機列印文件名稱(執行exe檔有效,VB跑無效)
'        'strFileName = strUserNum & "_CFP_" & IIf(bSpace = False, IIf(Trim(txt1(5)) <> "", Mid(Trim(txt1(5)), 1, 4), Mid(Trim(txt1(6)), 1, 4)), "空白") & iCount & ".pdf"
'        'If Dir(App.path & "\" & strFileName) <> "" Then
'        '   Kill App.path & "\" & strFileName
'        'End If
'        ''轉PDF
'        'frmPDF.Show
'        'frmPDF.StartProcess App.path, strFileName
'        'Call StrMenu(bSpace)
'        'frmPDF.EndtProcess
'        'Unload frmPDF
'        strFileName = strUserNum & "_CFP_" & IIf(bSpace = False, IIf(Trim(txt1(5)) <> "", Mid(Trim(txt1(5)), 1, 4), Mid(Trim(txt1(6)), 1, 4)), "空白") & iCount
'        App.Title = strFileName
'        Call StrMenu(bSpace)
'        'end 2017/06/07
'
'        'Added by Morgan 2017/4/18
'        '因pdf列印無法橫印,改將pdf右轉90度存檔後再印(reader沒有設定自動轉向時會縮小或被截掉)
'        'Remove by Lydia 2017/06/07
'        ' PUB_RotatePDF App.path & "\" & strFileName
'         'end 2017/4/18
'
'        'Added by Lydia 2017/05/16 用印記錄移到pdf建立
'        If iCount = 1 And strDetail <> "" Then
'           'If Dir(App.path & "\" & strFileName) <> "" Then 'Remove by Lydia 2020/03/16 因為不存檔案所以取消檔案檢查(自2017/06/08~2020/03/16無用印記錄)
'              If PUB_AddRecSeal("2", txtPCnt.Text, IIf(bSpace = True, "Y", ""), strDetail, Combo2.Text) Then
'              End If
'           'End If 'Remove by Lydia 2020/03/16
'        End If
'        'end 2017/05/16
'
'        'Remove by Lydia 2017/06/07
'        ''列印PDF
'        'PUB_PrintPDF App.path & "\" & strFileName, Me.Combo1
'        ''刪除PDF
'        'Kill App.path & "\" & strFileName
'    Next iCount
    Call runWordProc(bSpace)
    If m_TempPDF <> "" Then
        For iCount = 1 To Val(txtPCnt)
            If iCount = 1 Then
                PUB_RotatePDF App.path & "\" & strUserNum & "\" & m_TempPDF
            End If
            strFileName = strUserNum & "_CFP_" & m_TempFN & iCount
            PUB_PrintPDF App.path & "\" & strUserNum & "\" & m_TempPDF, Combo1.Text
            App.Title = strFileName
        Next iCount
    End If
'--------------先產生Word檔，後轉成PDF檔逐一列印

    App.Title = strOldName 'Added by Lydia 2017/06/07
    
End Sub

'Added by Morgan 2017/4/18
Public Function PUB_RotatePDF(pFileName As String) As Boolean
   Dim strCmd As String
   Dim strInput As String, strOutput As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim intI As Integer
   Dim oFilObj As FileSystemObject
   
   Set oFilObj = New FileSystemObject
   
   '合併程式遇中文檔名會錯,先更名
   strInput = App.path & "\$input.pdf"
   If oFilObj.FileExists(strInput) = True Then oFilObj.DeleteFile strInput, True
   oFilObj.MoveFile pFileName, strInput

   '輸出檔
   strOutput = App.path & "\$output.pdf"
   If oFilObj.FileExists(strOutput) = True Then oFilObj.DeleteFile strOutput, True
   
   strCmd = pub_PdftkEXE & " """ & strInput & """ cat 1-endR output """ & strOutput & """"
   process_id = SHELL(strCmd, vbHide)
   process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   If process_handle <> 0 Then
      For intI = 1 To 10
         If PUB_CheckIsRunning(pub_PdftkName) = True Then
            Sleep 1000
         Else
            Exit For
         End If
      Next
      If intI > 10 Then
         TerminateProcess process_handle, 0&
         CloseHandle process_handle
         MsgBox "PDF轉向失敗！"
         GoTo ErrHnd
      Else
         CloseHandle process_handle
      End If
   Else
      MsgBox "PDF轉向失敗！"
      GoTo ErrHnd
   End If
   '刪除輸入檔
   If oFilObj.FileExists(strInput) = True Then oFilObj.DeleteFile strInput, True
   '更名
   oFilObj.MoveFile strOutput, pFileName
   
   PUB_RotatePDF = True
     
ErrHnd:

   Set oFilObj = Nothing
End Function

'Added by Lydia 2022/01/18 下載Word範本套印
Private Sub runWordProc(ByVal pSpace As Boolean)
Dim iStrL(1 To 47) As String  '用印記錄(全文)
Dim iStrR(1 To 47) As String  '用印記錄(全文)
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
Dim strName As String
Dim strText As String
Dim intA As Integer
Dim m_FileName As String, m_TempFileName As String
Dim m_DefPath As String
Dim oShape

On Error GoTo ErrHand
   
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000300-0-02 智權部委任契約書_CFP.docx", "M51", "000300", "0", "02", "4", "1")
      
   m_DefPath = App.path & "\" & strUserNum
   'Added by Lydia 2022/01/25
   m_TempPDF = ""
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   'end 2022/01/25
   
   '下載範本檔: M51-000300-0-02 智權部委任契約書_CFP.docx
   m_TempFN = Pub_RepFileName(IIf(pSpace = False, Mid(Trim(txt1(5)), 1, 4), "空白")) 'Move by Lydia 2022/01/25 從m_TempFileName移過來
   'Modified by Lydia 2022/01/25 改成Word直接印，所以範本一開始就先命名好
   'm_FileName = "$$" & Me.Name & ".docx"
   m_FileName = "$$" & strUserNum & "_CFP_" & m_TempFN & ".docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000300-0-02", , m_DefPath) = False Then
        Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   'Remove by Lydia 2022/01/25 不用改存PDF檔
   'm_TempFileName = "$$" & strUserNum & "_CFP_" & m_TempFN & ".pdf"
   'If Dir(m_DefPath & "\" & m_TempFileName) <> "" Then
   '   Kill m_DefPath & "\" & m_TempFileName
   'End If
   'end 2022/01/25
   '改成直接用範本檔 Q: AddToRecentFiles:=False還是會新增到最近開啟記錄
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
      
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 0 To 57
         strName = "PS" & Format(intA, "000")
         strText = ""
         .Selection.Font.Bold = True '因為字小,所以全部用粗體字
'-------第一條
         If intA = 0 Then
              '專利之名稱
              strText = PUB_StrToStr(txt1(0), 64)
         ElseIf intA = 1 Then
              '發明人姓名-中文
              strText = PUB_StrToStr(txt1(1), 64)
         ElseIf intA = 2 Then
              '發明人姓名-英文
              strText = PUB_StrToStr(txt1(2), 64)
         ElseIf intA = 3 Then
               '發明人地址-中文
              strText = PUB_StrToStr(txt1(3), 64)
         ElseIf intA = 4 Then
               '發明人地址-英文
              strText = PUB_StrToStr(txt1(4), 64)
         ElseIf intA = 5 Then
               '申請人姓名-中文
              strText = PUB_StrToStr(txt1(5), 64)
         ElseIf intA = 6 Then
               '申請人姓名-英文
              strText = PUB_StrToStr(txt1(6), 64)
         ElseIf intA = 7 Then
               '申請人地址-中文
              strText = PUB_StrToStr(txt1(7), 64)
         ElseIf intA = 8 Then
               '申請人地址-英文
              strText = PUB_StrToStr(txt1(8), 64)
         ElseIf intA = 9 Then
              '指定聯絡人及通訊地址
              strText = PUB_StrToStr(txt1(9), 60)
'-------第三條
         ElseIf intA = 10 Then
              '前條第X款: 置中
              'Added by Lydia 2023/02/17 判斷長度不用置中;
              If LenB(StrConv(txt1(10).Text, vbFromUnicode)) >= 6 Then
                  strText = " " & Trim(txt1(10)) & " "
              Else
              'end 2023/02/17
                  strText = String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text) & String(6 - LenB(StrConv(String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text), vbFromUnicode)), " ")
              End If 'Added by Lydia 2023/02/17
         ElseIf intA = 11 Then
              'XX程序: 置中
              'Added by Lydia 2023/02/17 判斷長度不用置中;
              If LenB(StrConv(txt1(11).Text, vbFromUnicode)) >= 10 Then
                  strText = " " & Trim(txt1(11)) & " "
              Else
              'end 2023/02/17
                  strText = String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text) & String(10 - LenB(StrConv(String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text), vbFromUnicode)), " ")
              End If 'Added by Lydia 2023/02/17
         ElseIf intA = 13 Or intA = 19 Or intA = 25 Or intA = 31 Or intA = 37 Then
              '國別
              strText = PUB_StrToStr(txt1(intA), 12)
         ElseIf intA = 14 Or intA = 20 Or intA = 26 Or intA = 32 Or intA = 38 Then
              '專利種類
              strText = PUB_StrToStr(txt1(intA), 14)
         ElseIf intA = 53 Or intA = 15 Or intA = 21 Or intA = 27 Or intA = 33 Or intA = 39 Then
              '金額：1
              strText = PUB_StrToStr(txt1(intA), 14)
         ElseIf intA = 54 Or intA = 16 Or intA = 22 Or intA = 28 Or intA = 34 Or intA = 40 Then
              '金額：2
              strText = PUB_StrToStr(txt1(intA), 14)
         ElseIf intA = 12 Or intA = 17 Or intA = 23 Or intA = 29 Or intA = 35 Or intA = 41 Then
              '金額：3
              strText = PUB_StrToStr(txt1(intA), 14)
         ElseIf intA = 18 Or intA = 24 Or intA = 30 Or intA = 36 Or intA = 42 Then
              '備註
              strText = PUB_StrToStr(txt1(intA), 18)
         ElseIf intA = 43 Then
              '合計
              If Val(Trim(txt1(43))) = 0 Then
                  strExc(1) = String(12, "　")
              Else
                  'Modified by Lydia 2023/08/10 改變數控制
                  'strExc(1) = Replace(ChangeNumber(txt1(43)), "元整", "")
                  strExc(1) = ChangeNumber(txt1(43), False)
              End If
              'Added by Lydia 2022/06/24 判斷超過字元長度不限制
              If GetTextLength(strExc(1)) > 22 Then
                 strText = IIf(Trim(txt1(55).Text) = "", "　　　", Trim(txt1(55).Text)) & "　" & strExc(1) & "　"
              Else
              'end 2022/06/24
                 strText = IIf(Trim(txt1(55).Text) = "", "　　　", Trim(txt1(55).Text)) & "　" & PUB_StrToStr(strExc(1), 22, True, True) & "　"
              End If 'Added by Lydia 2022/06/24
'------其他
         ElseIf intA = 44 Then
              '會稿
              strText = IIf(opt1(0).Value = True, "V", "")
         ElseIf intA = 45 Then
              '不會稿
              strText = IIf(opt1(1).Value = True, "V", "")
         ElseIf intA = 46 Then
              '委任人
              strText = PUB_StrToStr(txt1(44).Text, 50)
         ElseIf intA = 47 Then
              '委任人-代表人
              strText = PUB_StrToStr(txt1(45).Text, 50)
         ElseIf intA = 48 Then
              '委任人-地址
              strText = PUB_StrToStr(txt1(46).Text, 50)
         ElseIf intA = 49 Then
              '委任人-電話
              strText = PUB_StrToStr(txt1(47).Text, 18)
         ElseIf intA = 50 Then
              '委任人-傳真
              strText = PUB_StrToStr(txt1(48).Text, 18)
         ElseIf intA = 51 Then
              '受任人
              strText = Combo2.Text
         ElseIf intA = 52 Then
              '經手人
              strText = PUB_StrToStr(txt1(49).Text, 50)
         ElseIf intA = 55 Then
              '受任人-地址
              strText = PUB_SetAddrTofrm210114(Combo2.Text)
         ElseIf intA = 56 Then
              strText = "    中    華    民    國 " & String((6 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & txt1(50) & String((6 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & "年" & String((6 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & txt1(51) & String((6 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & "月" & String((6 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & txt1(52) & String((6 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & "日"
         ElseIf intA = 57 Then
              strText = ""
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 0 And intA <= 9) Or (intA >= 46 And intA <= 49) Or intA = 52 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "細明體-ExtB"
            End If
            If intA = 53 Or intA = 54 Then
                '委辦費用
                .Selection.Font.Size = 8
            End If
            If intA = 44 Or intA = 45 Then
                '會稿/不會稿勾選
                .Selection.Font.Size = 14
            End If
            If intA = 57 And bolAddSeal = True Then  '公司章: 放在受任人的儲存格
                strExc(9) = Mid(strCompSeal, InStr(strCompSeal, Combo2))
                If InStr(strExc(9), ",") > 0 Then
                    strExc(9) = Right(Mid(strExc(9), 1, InStr(strExc(9), ",") - 1), 2)
                Else
                    strExc(9) = Right(strExc(9), 2)
                End If
                If PUB_ReadDB2File(m_DefPath & "\$$" & Me.Name & "TempFile", Val(strExc(9))) Then
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=m_DefPath & "\$$" & Me.Name & "TempFile", LinkToFile:=False, SaveWithDocument:=True)
                    '--------設定圖片=文蓋圖(文字在前)
                        oShape.Fill.Visible = msoFalse
                        oShape.Fill.Solid
                        oShape.Fill.Transparency = 0#
                        oShape.Line.Weight = 0.75
                        oShape.Line.DashStyle = msoLineSolid
                        oShape.Line.Style = msoLineSingle
                        oShape.Line.Transparency = 0#
                        oShape.Line.Visible = msoFalse
                        oShape.LockAspectRatio = msoTrue
                        oShape.Rotation = 0#
                        oShape.PictureFormat.Brightness = 0.5
                        oShape.PictureFormat.Contrast = 0.5
                        oShape.PictureFormat.ColorType = msoPictureAutomatic
                        oShape.PictureFormat.CropLeft = 0#
                        oShape.PictureFormat.CropRight = 0#
                        oShape.PictureFormat.CropTop = 0#
                        oShape.PictureFormat.CropBottom = 0#
                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(7.8)
                        oShape.Top = .CentimetersToPoints(0.3)
                        oShape.LockAnchor = False
                        oShape.LayoutInCell = True
                        oShape.WrapFormat.AllowOverlap = True
                        oShape.WrapFormat.Side = wdWrapBoth
                        oShape.WrapFormat.DistanceTop = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceBottom = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceLeft = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.DistanceRight = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.Type = 3
                        oShape.ZOrder 5 '文蓋圖(文字在前)
                        '---------------------------
                End If
          
            End If
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 0 And intA <= 9) Or (intA >= 46 And intA <= 49) Or intA = 52 Then
               '有Unicode字需要換字型=>還原
               .Selection.Font.Name = "標楷體"
            End If
            If intA = 53 Or intA = 54 Then
                '委辦費用=>還原
                .Selection.Font.Size = 11
            End If
            If intA = 44 Or intA = 45 Then
                '會稿/不會稿勾選=>還原
                .Selection.Font.Size = 11
            End If
         End If

      Next intA
'      '因為先全部以細明體-ExtB,最後全選改字型;
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   'Modified by Lydia 2022/01/19 改存成PDF檔
   'Memo by Lydia 2022/01/25  因為受PDF redirect設定灰階列印影響，改成Word直接印
   intA = IIf(Val(txtPCnt) = 0, 1, Val(txtPCnt))
   For intI = 1 To intA
       g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=1, Pages:="1", Collate:=True
   Next intI
   
   '保留: 存檔
   'g_WordAp.ActiveDocument.Close wdSaveChanges
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   m_TempPDF = m_FileName 'Added by Lydia 2022/01/25
   
   'Mark by Lydia 2022/01/25 因為受PDF redirect設定灰階列印影響，改成Word直接印
   'If PUB_PrintWord2PDF(g_WordAp, m_DefPath, m_TempFileName, m_TempPDF) = False Then
   '    Exit Sub
   'End If
   'end 2022/01/19
   
If bolAddSeal = True Then  '用印記錄
   strDetail = ""
   iStrL(1) = "專利案件委任契約書"
   iStrL(2) = "委任人（甲方）茲委任受任人（乙方）辦理專利案件，雙方同意條件如下："
   iStrL(3) = "第一條  專利之名稱：" & StrToStr(txt1(0) & String(74, " "), 37)
   iStrL(4) = ""
   iStrL(5) = "　　　　　　　（中）" & StrToStr(txt1(1) & String(74, " "), 37)
   iStrL(6) = "　　發明人姓名："
   iStrL(7) = "　　　　　　　（英）" & StrToStr(txt1(2) & String(74, " "), 37)
   iStrL(8) = ""
   iStrL(9) = "　　　　　　　（中）" & StrToStr(txt1(3) & String(74, " "), 37)
   iStrL(10) = "　　發明人地址："
   iStrL(11) = "　　　　　　　（英）" & StrToStr(txt1(4) & String(74, " "), 37)
   iStrL(12) = ""
   iStrL(13) = "　　　　　　　（中）" & StrToStr(txt1(5) & String(74, " "), 37)
   iStrL(14) = "　　申請人姓名："
   iStrL(15) = "　　　　　　　（英）" & StrToStr(txt1(6) & String(74, " "), 37)
   iStrL(16) = ""
   iStrL(17) = "　　　　　　　（中）" & StrToStr(txt1(7) & String(74, " "), 37)
   iStrL(18) = "　　申請人地址："
   iStrL(19) = "　　　　　　　（英）" & StrToStr(txt1(8) & String(74, " "), 37)
   iStrL(20) = "　　指定聯絡人"
   iStrL(21) = "　　及通訊地址：" & StrToStr(txt1(9) & String(78, " "), 39)
   iStrL(22) = "第二條　委辦範圍"
      For intI = 1 To 22
         If Trim(iStrL(intI)) <> "" Then
            strDetail = strDetail & RTrim(iStrL(intI)) & vbCrLf
         End If
      Next

   iStrL(23) = "一、提出申請程序：乙方根據甲方所提供之發明創作資料或樣品，代撰專利說明書及圖式並向我國專利主管"
   iStrL(24) = "　　　　　　　　　機關提出申請，或為甲方委請各該申請國之專利代理人，代向各該國提出專利申請。"
   iStrL(25) = "二、中間處理程序：提出專利申請後依甲方請求或各該申請國專利主管機關之指示所需提出之修正、更正、"
   iStrL(26) = "　　　　　　　　　補充說明書之各程序。"
   iStrL(27) = "三、領證及繳納年費程序：依各該申請國專利法規之規定在申請程序中繳納維持費或審定核准後繳納證書費"
   iStrL(28) = "　　　　　　　　　、年費、維持費，領取證書。"
   iStrL(29) = "四、救 濟 程 序 ：審定不予專利時，所需進行之救濟程序，如請求再審、提起審判，以及答辯等程序。"
   iStrL(30) = "五、讓渡 或授權 ：乙方提供各該國之空白讓渡書或授權書供甲方簽署，或根據甲方所提供之讓渡書或授權"
   iStrL(31) = "　　　　　　　　　書，自行或委請各該國專利代理人，代向該國專利主管機關辦理登記。"
   iStrL(32) = "六、專 利 調 卷 ：乙方根據甲方所提供之資料進行調卷。"
   iStrL(33) = "七、專 利 調 查 ：乙方根據甲方所提供之資料，自行或委請該國代理人進行專利調查，此專利調查並不含"
   iStrL(34) = "　　　　　　　　　專利調卷。"
   iStrL(35) = "八、其        他：如說明書、圖式之更正、新穎性調查、外國專利主管機關文件引證資料之翻譯、公告或"
   iStrL(36) = "　　　　　　　　　延期公告之申請等程序。"
   iStrL(37) = "第三條 委辦費用"
   iStrL(38) = "　　一、乙方受委辦前條第" & String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text) & String(6 - LenB(StrConv(String(Int((6 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text), vbFromUnicode)), " ") & "款" & String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text) & String(10 - LenB(StrConv(String(Int((10 - LenB(StrConv(txt1(11).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(11).Text), vbFromUnicode)), " ") & "程序之費用（申請國外專利時，包括國外代理人費用），約定如下："
   iStrL(39) = "　　　　　　　　　　　　　　　　　　　　　　金　　　　　　　　　　　　　額　　　　　　　　　　　　"
   iStrL(40) = "　　　國     別   　　專利種類　　　　　　　　　　　　　　　　　　　　　　　　　　　 備　　　註"
   iStrL(41) = "　  　　　　　　　　　　　　　　　" & Pub_StrToCenter(txt1(53).Text, 16) & Pub_StrToCenter(txt1(54).Text, 16) & Pub_StrToCenter(txt1(12).Text, 16)
   iStrL(42) = "　" & Pub_StrToCenter(txt1(13).Text, 16) & Pub_StrToCenter(txt1(14).Text, 16) & Pub_StrToCenter(txt1(15).Text, 16) & Pub_StrToCenter(txt1(16).Text, 16) & Pub_StrToCenter(txt1(17).Text, 16) & Pub_StrToCenter(txt1(18).Text, 16)
   iStrL(43) = "　" & Pub_StrToCenter(txt1(19).Text, 16) & Pub_StrToCenter(txt1(20).Text, 16) & Pub_StrToCenter(txt1(21).Text, 16) & Pub_StrToCenter(txt1(22).Text, 16) & Pub_StrToCenter(txt1(23).Text, 16) & Pub_StrToCenter(txt1(24).Text, 16)
   iStrL(44) = "　" & Pub_StrToCenter(txt1(25).Text, 16) & Pub_StrToCenter(txt1(26).Text, 16) & Pub_StrToCenter(txt1(27).Text, 16) & Pub_StrToCenter(txt1(28).Text, 16) & Pub_StrToCenter(txt1(29).Text, 16) & Pub_StrToCenter(txt1(30).Text, 16)
   iStrL(45) = "　" & Pub_StrToCenter(txt1(31).Text, 16) & Pub_StrToCenter(txt1(32).Text, 16) & Pub_StrToCenter(txt1(33).Text, 16) & Pub_StrToCenter(txt1(34).Text, 16) & Pub_StrToCenter(txt1(35).Text, 16) & Pub_StrToCenter(txt1(36).Text, 16)
   iStrL(46) = "　" & Pub_StrToCenter(txt1(37).Text, 16) & Pub_StrToCenter(txt1(38).Text, 16) & Pub_StrToCenter(txt1(39).Text, 16) & Pub_StrToCenter(txt1(40).Text, 16) & Pub_StrToCenter(txt1(41).Text, 16) & Pub_StrToCenter(txt1(42).Text, 16)
   If Val(Trim(txt1(43))) = 0 Then
       strSpaceAmt = String(12, "　")
   Else
       'Modified by Lydia 2023/08/10 改變數控制
       'strSpaceAmt = Replace(ChangeNumber(txt1(43)), "元整", "")
       strSpaceAmt = ChangeNumber(txt1(43), False)
   End If
   iStrL(47) = "合計" & IIf(Trim(txt1(55).Text) = "", "　　　", Trim(txt1(55).Text)) & "　" & String(24, " ") & "元整，於本契約簽訂同時由甲方一次付清。"
        strExc(1) = ""
      For intI = 13 To 42
          '有無費用項目
          strExc(1) = strExc(1) & Trim(Replace(Replace(txt1(intI), "　", ""), " ", ""))
      Next
      For intI = 37 To 46
          If Trim(Replace(Replace(iStrL(intI), "　", ""), " ", "")) <> "" Then
             Select Case intI
                 Case 40
                     strExc(1) = "　　　國     別   　　專利種類 " & Pub_StrToCenter(txt1(53).Text, 16) & Pub_StrToCenter(txt1(54).Text, 16) & Pub_StrToCenter(txt1(12).Text, 16) & " 備　　　註"
                     strDetail = strDetail & RTrim(strExc(1)) & vbCrLf
                 Case 41
                 Case Else
                      strDetail = strDetail & RTrim(iStrL(intI)) & vbCrLf
             End Select
          End If
      Next
      'Modified by Lydia 2023/08/10 增加判斷
      'strDetail = strDetail & "合計" & Trim(txt1(55).Text) & "　" & strSpaceAmt & "元整，於本契約簽訂同時由甲方一次付清。" & vbCrLf
      strDetail = strDetail & "合計" & Trim(txt1(55).Text) & "　" & strSpaceAmt & IIf(InStr(strSpaceAmt, "元") > 0, "整", "元整") & "，於本契約簽訂同時由甲方一次付清。" & vbCrLf
      
   iStrR(1) = ""
   iStrR(2) = ""
   iStrR(3) = "　　二、第二條所列之委辦範圍除本條第一款特予載明外，其費用由甲方負擔，申請國外專利時，其金額依"
   iStrR(4) = "　　　　當時外國代理人費用及本所服務費標準收取之。"
   iStrR(5) = "　　三、案件之進行中，如需乙方派員前往現場研究或繪圖時，其出差旅費，由甲方負擔，並按實際處理時"
   iStrR(6) = "　　　　間每小時新台幣貳仟元整計算另收費用。"
   iStrR(7) = "　　　　本條所約定之費用如甲方未於乙方所指定之期限內付清，則乙方無義務辦理所受任之事項，且經乙"
   iStrR(8) = "　　　　方限期催告後，如甲方仍不履行時，則本契約當然終止，乙方得終止進行該程序及嗣後之一切程序"
   iStrR(9) = "　　　　，並通知各該外國代理人終止進行所有相關程序。另乙方已先行代辦之服務費用，甲方仍應照付。"
   iStrR(10) = "第四條  乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方權益之"
   iStrR(11) = "　　　　疏誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額的三倍為限。"
   iStrR(12) = "第五條  甲方應確保所交付予乙方之資料及本契約書所載內容(包括發明人或創作人、申請人等資訊)均無虛"
   iStrR(13) = "　　　　偽情事，且甲方確實得到與委辦案件相關共同發明人及第三人之同意，有權委託乙方辦理案件，如"
   iStrR(14) = "　　　　因不實致生損害或法律責任時，概由甲方負責，與乙方無關。"
   iStrR(15) = "第六條  乙方於辦理過程中，應隨時將辦理經過如申請日、案號及其他重要函件，儘速通知或交付甲方。但"
   iStrR(16) = "　　　　甲方簽約後變更連絡處所，未即時通知乙方，因而連絡不及致誤時限者，乙方不負責任。"
   iStrR(17) = "第七條  凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責任。經乙方通知"
   iStrR(18) = "　　　　甲方繳費而未依限繳納者，亦同。"
   iStrR(19) = "第八條  甲方如逕自撤回所委辦之程序，或未經乙方同意終止契約時，所約定之費用，仍應全數給付。"
   iStrR(20) = "第九條  本約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需甲乙雙方於更動處"
   iStrR(21) = "　　　　蓋章始生效力，並由雙方各執乙份為憑。"
   iStrR(22) = ""
   iStrR(23) = "附  則  乙方所撰寫之專利說明書、圖式、再審文件或修正書，是否需要會稿，請於下方方格註明。"
   iStrR(24) = ""
   iStrR(25) = ""
   iStrR(26) = " 　   　　     會稿　　　　 " & IIf(opt1(0).Value = True, "Ｖ", "　") & "　　　　不會稿　　　　 " & IIf(opt1(1).Value = True, "Ｖ", "　")
   iStrR(27) = ""
   iStrR(28) = ""
   iStrR(29) = "　　甲方：委任人：" & StrToStr(txt1(44) & String(57, " "), 28)
   iStrR(30) = ""
   iStrR(31) = "　　　　　代表人：" & StrToStr(txt1(45) & String(57, " "), 28)
   iStrR(32) = ""
   iStrR(33) = "　　      地  址：" & StrToStr(txt1(46) & String(57, " "), 28)
   iStrR(34) = ""
   iStrR(35) = "　　　　　電  話：" & StrToStr(txt1(47) & String(28, " "), 14) & "傳  真：" & StrToStr(txt1(48) & String(26, " "), 13)
   iStrR(36) = ""
   iStrR(37) = ""
   iStrR(38) = "　　乙方：受任人：" & Combo2.Text '"台一國際專利法律事務所                               "
   iStrR(39) = ""
   iStrR(40) = "　　　　　經手人：" & StrToStr(txt1(49) & String(57, " "), 28)
   iStrR(41) = ""
   iStrR(42) = "　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   iStrR(43) = ""
   iStrR(44) = "　　　　　電  話：(02)25061023(總機)  　　　　F A X ：(02)25011666"
   iStrR(45) = ""
   If Combo2 = "台一智權股份有限公司" Then
      iStrR(46) = ""
   Else
      iStrR(46) = "　　　　　網  址：www.taie.com.tw     　　　　E-mail：ipdept@taie.com.tw"
   End If
   iStrR(47) = " 中  華  民  國 " & String((10 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & txt1(50) & String((10 - LenB(StrConv((txt1(50)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & txt1(51) & String((10 - LenB(StrConv((txt1(51)), vbFromUnicode))) / 2, " ") & "月" & String((8 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & txt1(52) & String((8 - LenB(StrConv((txt1(52)), vbFromUnicode))) / 2, " ") & "日"
   
      strDetail = strDetail & vbCrLf & IIf(opt1(0).Value = True, "會稿", "不會稿") & vbCrLf & vbCrLf
      For intI = 29 To 40
         If Trim(iStrR(intI)) <> "" Then
            strDetail = strDetail & RTrim(iStrR(intI)) & vbCrLf
         End If
      Next
      strDetail = strDetail & iStrR(47)
      
      If PUB_AddRecSeal("2", txtPCnt.Text, IIf(pSpace = True, "Y", ""), strDetail, Combo2.Text) Then
      End If
End If
          
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2022/01/20 刪除暫存檔
Private Sub RunEndProc(ByVal bolSleep As Boolean)
   
   If bolSleep = True Then Sleep 3000
   PUB_KillTempFile (strUserNum & "\$$" & strUserNum & "*_CFP*.*")
   PUB_KillTempFile (strUserNum & "\$$" & Me.Name & "*.*")
    
End Sub
