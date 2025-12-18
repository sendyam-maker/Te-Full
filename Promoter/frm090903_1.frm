VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090903_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "待命名"
   ClientHeight    =   7080
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8892
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8892
   Begin VB.CommandButton cmdOpen 
      Caption         =   "外文本(&P)"
      Height          =   300
      Left            =   120
      TabIndex        =   179
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔(&S)"
      Height          =   360
      Index           =   0
      Left            =   5400
      TabIndex        =   176
      Top             =   15
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認回報(&A)"
      Height          =   360
      Index           =   1
      Left            =   6357
      TabIndex        =   177
      Top             =   15
      Width           =   1160
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Left            =   7615
      TabIndex        =   178
      Top             =   15
      Width           =   1160
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6612
      Left            =   120
      TabIndex        =   40
      Top             =   420
      Width           =   8652
      _ExtentX        =   15261
      _ExtentY        =   11663
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   4057
      TabCaption(0)   =   "案件名稱"
      TabPicture(0)   =   "frm090903_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblData(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblData(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblData(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblData(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblData(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblData(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblData(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblData(10)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label4(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label5(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label4(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblData(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label4(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblData(5)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label6(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblData(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label5(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label4(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(6)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(7)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Line1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(46)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblData(12)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblData(13)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblData(15)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblData(16)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label5(2)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label1(54)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lblCMboth"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtData(3)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtData(4)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtData(5)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Frame1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Frame2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Frame3"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Frame4"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdOK(2)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Combo1"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Frame5(8)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "CmdPA174"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "ChkPA174"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Chk2(51)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "說明書"
      TabPicture(1)   =   "frm090903_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5(12)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "申請專利範圍 ＆ 圖示"
      TabPicture(2)   =   "frm090903_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5(6)"
      Tab(2).Control(1)=   "Frame5(7)"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox Chk2 
         Caption         =   "專利權期間延長相關"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   238
         Top             =   2070
         Width           =   2175
      End
      Begin VB.CheckBox ChkPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   75
         TabIndex        =   235
         Top             =   2610
         Width           =   1035
      End
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   210
         Style           =   1  '圖片外觀
         TabIndex        =   234
         Top             =   2880
         Width           =   840
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame5"
         Height          =   5865
         Index           =   12
         Left            =   -74880
         TabIndex        =   200
         Top             =   420
         Width           =   8385
         Begin VB.CheckBox Chk2 
            Caption         =   "其它問題"
            Height          =   255
            Index           =   19
            Left            =   30
            TabIndex        =   90
            Top             =   5520
            Width           =   1215
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "缺頁"
            Height          =   255
            Index           =   18
            Left            =   390
            TabIndex        =   88
            Top             =   5235
            Width           =   855
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "建議內容"
            Height          =   255
            Index           =   17
            Left            =   870
            TabIndex        =   86
            Top             =   4920
            Width           =   1095
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "缺摘要"
            Height          =   255
            Index           =   16
            Left            =   390
            TabIndex        =   85
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame5"
            Height          =   660
            Index           =   5
            Left            =   720
            TabIndex        =   227
            Top             =   3990
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   14
               Left            =   960
               TabIndex        =   79
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_6 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   80
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_6 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   81
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   15
               Left            =   960
               TabIndex        =   83
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   18
               Left            =   4200
               TabIndex        =   82
               Top             =   0
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   19
               Left            =   2040
               TabIndex        =   84
               Top             =   320
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "符號說明"
               Height          =   255
               Index           =   31
               Left            =   0
               TabIndex        =   231
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   32
               Left            =   3720
               TabIndex        =   230
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   33
               Left            =   1800
               TabIndex        =   229
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   34
               Left            =   3480
               TabIndex        =   228
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame5"
            Height          =   660
            Index           =   4
            Left            =   720
            TabIndex        =   222
            Top             =   3300
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   12
               Left            =   960
               TabIndex        =   73
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_5 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   74
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_5 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   75
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   13
               Left            =   960
               TabIndex        =   77
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   16
               Left            =   4200
               TabIndex        =   76
               Top             =   0
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   17
               Left            =   2040
               TabIndex        =   78
               Top             =   320
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "實施方式"
               Height          =   255
               Index           =   26
               Left            =   0
               TabIndex        =   226
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   27
               Left            =   3720
               TabIndex        =   225
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   28
               Left            =   1800
               TabIndex        =   224
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   30
               Left            =   3480
               TabIndex        =   223
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame5"
            Height          =   660
            Index           =   3
            Left            =   720
            TabIndex        =   216
            Top             =   2610
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   10
               Left            =   960
               TabIndex        =   67
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_4 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   68
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_4 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   69
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   11
               Left            =   960
               TabIndex        =   71
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   14
               Left            =   4200
               TabIndex        =   70
               Top             =   0
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   15
               Left            =   2040
               TabIndex        =   72
               Top             =   320
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "圖式簡單"
               Height          =   240
               Index           =   20
               Left            =   0
               TabIndex        =   221
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   21
               Left            =   3720
               TabIndex        =   220
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   22
               Left            =   1800
               TabIndex        =   219
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   23
               Left            =   3480
               TabIndex        =   218
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "說明"
               Height          =   255
               Index           =   35
               Left            =   360
               TabIndex        =   217
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame5"
            Height          =   660
            Index           =   2
            Left            =   720
            TabIndex        =   211
            Top             =   1920
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   9
               Left            =   960
               TabIndex        =   65
               Top             =   320
               Width           =   1095
            End
            Begin VB.OptionButton Opt3_3 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   63
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton Opt3_3 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   62
               Top             =   0
               Width           =   615
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   8
               Left            =   960
               TabIndex        =   61
               Top             =   0
               Width           =   735
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   13
               Left            =   2040
               TabIndex        =   66
               Top             =   320
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   12
               Left            =   4200
               TabIndex        =   64
               Top             =   0
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   16
               Left            =   3480
               TabIndex        =   215
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   17
               Left            =   1800
               TabIndex        =   214
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   18
               Left            =   3720
               TabIndex        =   213
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "發明內容"
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   212
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame5"
            Height          =   660
            Index           =   1
            Left            =   720
            TabIndex        =   206
            Top             =   1230
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   7
               Left            =   960
               TabIndex        =   59
               Top             =   320
               Width           =   1095
            End
            Begin VB.OptionButton Opt3_2 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   57
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton Opt3_2 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   56
               Top             =   0
               Width           =   615
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   180
               Index           =   6
               Left            =   960
               TabIndex        =   55
               Top             =   0
               Width           =   735
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   11
               Left            =   2040
               TabIndex        =   60
               Top             =   320
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   10
               Left            =   4200
               TabIndex        =   58
               Top             =   0
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   12
               Left            =   3480
               TabIndex        =   210
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   13
               Left            =   1800
               TabIndex        =   209
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   14
               Left            =   3720
               TabIndex        =   208
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "先前技術"
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   207
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame5"
            Height          =   660
            Index           =   0
            Left            =   720
            TabIndex        =   201
            Top             =   540
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   4
               Left            =   960
               TabIndex        =   49
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_1 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   50
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_1 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   51
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   5
               Left            =   960
               TabIndex        =   53
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   8
               Left            =   4200
               TabIndex        =   52
               Top             =   0
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   9
               Left            =   2040
               TabIndex        =   54
               Top             =   320
               Width           =   3000
               VariousPropertyBits=   671105051
               Size            =   "5292;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label1 
               Caption         =   "技術領域"
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   205
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   10
               Left            =   3720
               TabIndex        =   204
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   11
               Left            =   1800
               TabIndex        =   203
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   29
               Left            =   3480
               TabIndex        =   202
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "內容不完整"
            Height          =   255
            Index           =   3
            Left            =   420
            TabIndex        =   48
            Top             =   260
            Width           =   1215
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "說明書"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   47
            Top             =   0
            Width           =   1215
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   22
            Left            =   1230
            TabIndex        =   91
            Top             =   5520
            Width           =   6795
            VariousPropertyBits=   671105051
            Size            =   "11994;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   21
            Left            =   1950
            TabIndex        =   89
            Top             =   5220
            Width           =   3000
            VariousPropertyBits=   671105051
            Size            =   "5292;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   20
            Left            =   1950
            TabIndex        =   87
            Top             =   4920
            Width           =   3000
            VariousPropertyBits=   671105051
            Size            =   "5292;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblPS 
            Caption         =   "P.S 如果要設定""缺""或""須修正"", 請在""位置""欄位內輸入文字或空白鍵."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   2550
            TabIndex        =   233
            Top             =   260
            Width           =   5715
         End
         Begin VB.Label Label1 
            Caption         =   "頁數："
            Height          =   255
            Index           =   36
            Left            =   1350
            TabIndex        =   232
            Top             =   5235
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Height          =   2580
         Index           =   8
         Left            =   6240
         TabIndex        =   187
         Top             =   3315
         Width           =   2175
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Height          =   480
            Index           =   10
            Left            =   480
            TabIndex        =   195
            Top             =   400
            Width           =   1215
            Begin VB.OptionButton Opt4s2 
               Caption         =   "提申後"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   14
               Top             =   0
               Width           =   975
            End
            Begin VB.OptionButton Opt4s2 
               Caption         =   "提申前"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不請款"
            Height          =   255
            Index           =   48
            Left            =   600
            TabIndex        =   16
            Top             =   900
            Width           =   1020
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不需收文"
            Height          =   255
            Index           =   46
            Left            =   240
            TabIndex        =   21
            Top             =   2280
            Width           =   1620
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "需收文告代"
            Height          =   255
            Index           =   47
            Left            =   240
            TabIndex        =   17
            Top             =   1200
            Width           =   1620
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "需收文主動修正"
            Height          =   255
            Index           =   45
            Left            =   240
            TabIndex        =   13
            Top             =   120
            Width           =   1620
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Height          =   840
            Index           =   9
            Left            =   480
            TabIndex        =   192
            Top             =   1440
            Width           =   1575
            Begin VB.OptionButton Opt4s 
               Caption         =   "當日告代"
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton Opt4s 
               Caption         =   "提申前告代"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   330
               Width           =   1335
            End
            Begin VB.OptionButton Opt4s 
               Caption         =   "提申後告代"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   90
               Width           =   1335
            End
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   4680
         TabIndex        =   38
         Text            =   "Combo1"
         Top             =   6204
         Width           =   2655
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "列印(&P)"
         Height          =   330
         Index           =   2
         Left            =   7560
         TabIndex        =   39
         Top             =   6168
         Width           =   860
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Height          =   2145
         Index           =   7
         Left            =   -74880
         TabIndex        =   183
         Top             =   4140
         Width           =   8415
         Begin VB.CheckBox Chk2 
            Caption         =   "其它問題"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   141
            Top             =   1760
            Width           =   1095
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "用途說明"
            Height          =   255
            Index           =   39
            Left            =   4920
            TabIndex        =   148
            Top             =   1500
            Width           =   1695
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "色彩"
            Height          =   255
            Index           =   38
            Left            =   4920
            TabIndex        =   147
            Top             =   1215
            Width           =   975
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不主張設計的部分"
            Height          =   255
            Index           =   37
            Left            =   4920
            TabIndex        =   146
            Top             =   930
            Width           =   1935
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "超過一個實施例"
            Height          =   255
            Index           =   36
            Left            =   4920
            TabIndex        =   145
            Top             =   645
            Width           =   1695
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不完整：圖"
            Height          =   255
            Index           =   35
            Left            =   4920
            TabIndex        =   143
            Top             =   360
            Width           =   1245
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "格式不符（圖"
            Height          =   255
            Index           =   34
            Left            =   480
            TabIndex        =   138
            Top             =   960
            Width           =   1425
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "彩圖"
            Height          =   255
            Index           =   33
            Left            =   1320
            TabIndex        =   136
            Top             =   660
            Width           =   720
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "缺圖（　　　　）圖"
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   135
            Top             =   660
            Width           =   2055
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "建議指定代表圖：圖"
            Height          =   255
            Index           =   31
            Left            =   480
            TabIndex        =   133
            Top             =   360
            Width           =   2025
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "圖式"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   132
            Top             =   120
            Width           =   975
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   45
            Left            =   6240
            TabIndex        =   144
            Top             =   360
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   46
            Left            =   1200
            TabIndex        =   142
            Top             =   1760
            Width           =   6800
            VariousPropertyBits=   671105051
            Size            =   "11994;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   44
            Left            =   2280
            TabIndex        =   140
            Top             =   1240
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   43
            Left            =   1920
            TabIndex        =   139
            Top             =   945
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   42
            Left            =   2520
            TabIndex        =   137
            Top             =   645
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   41
            Left            =   2520
            TabIndex        =   134
            Top             =   360
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   48
            Left            =   3960
            TabIndex        =   185
            Top             =   1240
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，說明"
            Height          =   255
            Index           =   37
            Left            =   1560
            TabIndex        =   184
            Top             =   1240
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Height          =   3795
         Index           =   6
         Left            =   -74880
         TabIndex        =   165
         Top             =   360
         Width           =   8415
         Begin VB.CheckBox Chk2 
            Caption         =   "其它問題"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   130
            Top             =   3405
            Width           =   1095
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "混雜式請求項：請求項"
            Height          =   255
            Index           =   28
            Left            =   600
            TabIndex        =   128
            Top             =   3135
            Width           =   2175
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不予專利（請求項"
            Height          =   255
            Index           =   27
            Left            =   600
            TabIndex        =   124
            Top             =   2820
            Width           =   1785
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "標的不一致"
            Height          =   255
            Index           =   26
            Left            =   600
            TabIndex        =   119
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "屬於引用記載形式之獨立項：請求項"
            Height          =   255
            Index           =   25
            Left            =   600
            TabIndex        =   117
            Top             =   1605
            Width           =   3375
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "多附多（附屬項"
            Height          =   255
            Index           =   24
            Left            =   600
            TabIndex        =   114
            Top             =   1320
            Width           =   1785
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "依附關係不明確（附屬項"
            Height          =   255
            Index           =   23
            Left            =   600
            TabIndex        =   111
            Top             =   1020
            Width           =   2340
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "依附關係錯誤（附屬項"
            Height          =   255
            Index           =   22
            Left            =   600
            TabIndex        =   108
            Top             =   720
            Width           =   2235
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "項號錯誤：請求項"
            Height          =   255
            Index           =   21
            Left            =   600
            TabIndex        =   106
            Top             =   420
            Width           =   1755
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "申請專利範圍"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   105
            Top             =   120
            Width           =   1575
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   40
            Left            =   1200
            TabIndex        =   131
            Top             =   3405
            Width           =   6800
            VariousPropertyBits=   671105051
            Size            =   "11994;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   39
            Left            =   2760
            TabIndex        =   129
            Top             =   3120
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   38
            Left            =   6240
            TabIndex        =   127
            Top             =   2805
            Width           =   1300
            VariousPropertyBits=   671105051
            Size            =   "2293;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   37
            Left            =   4320
            TabIndex        =   126
            Top             =   2812
            Width           =   1300
            VariousPropertyBits=   671105051
            Size            =   "2293;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   36
            Left            =   2400
            TabIndex        =   125
            Top             =   2805
            Width           =   1300
            VariousPropertyBits=   671105051
            Size            =   "2293;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   35
            Left            =   4680
            TabIndex        =   123
            Top             =   2505
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   34
            Left            =   2280
            TabIndex        =   122
            Top             =   2505
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   33
            Left            =   4080
            TabIndex        =   121
            Top             =   2205
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   32
            Left            =   1680
            TabIndex        =   120
            Top             =   2205
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   31
            Left            =   3960
            TabIndex        =   118
            Top             =   1590
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   30
            Left            =   5640
            TabIndex        =   116
            Top             =   1305
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   29
            Left            =   2400
            TabIndex        =   115
            Top             =   1305
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   28
            Left            =   6240
            TabIndex        =   113
            Top             =   1005
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   27
            Left            =   3000
            TabIndex        =   112
            Top             =   1005
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   26
            Left            =   6120
            TabIndex        =   110
            Top             =   720
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   25
            Left            =   2880
            TabIndex        =   109
            Top             =   720
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   24
            Left            =   2400
            TabIndex        =   107
            Top             =   405
            Width           =   1695
            VariousPropertyBits=   671105051
            Size            =   "2990;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   44
            Left            =   7560
            TabIndex        =   182
            Top             =   2813
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，法條"
            Height          =   255
            Index           =   39
            Left            =   5640
            TabIndex        =   181
            Top             =   2820
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "，標的："
            Height          =   255
            Index           =   38
            Left            =   3720
            TabIndex        =   180
            Top             =   2820
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "，標的："
            Height          =   255
            Index           =   52
            Left            =   3960
            TabIndex        =   175
            Top             =   2520
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "被依附之請求項"
            Height          =   255
            Index           =   51
            Left            =   960
            TabIndex        =   174
            Top             =   2520
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "，標的："
            Height          =   255
            Index           =   50
            Left            =   3360
            TabIndex        =   173
            Top             =   2220
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "附屬項"
            Height          =   255
            Index           =   49
            Left            =   960
            TabIndex        =   172
            Top             =   2220
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   47
            Left            =   7365
            TabIndex        =   171
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，應依附於請求項"
            Height          =   255
            Index           =   45
            Left            =   4095
            TabIndex        =   170
            Top             =   1320
            Width           =   1560
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   43
            Left            =   7980
            TabIndex        =   169
            Top             =   1020
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，應依附於請求項"
            Height          =   255
            Index           =   42
            Left            =   4695
            TabIndex        =   168
            Top             =   1020
            Width           =   1560
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   41
            Left            =   7860
            TabIndex        =   167
            Top             =   735
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，應依附於請求項"
            Height          =   255
            Index           =   40
            Left            =   4575
            TabIndex        =   166
            Top             =   735
            Width           =   1560
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   1068
         Left            =   120
         TabIndex        =   162
         Top             =   5115
         Width           =   6015
         Begin VB.CheckBox Chk27 
            Caption         =   "湃傳思"
            Height          =   255
            Index           =   5
            Left            =   1140
            TabIndex        =   37
            Top             =   792
            Width           =   850
         End
         Begin VB.CheckBox Chk27 
            Caption         =   "捷恩凱"
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   239
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Chk27 
            Caption         =   "其他"
            Height          =   255
            Index           =   0
            Left            =   3900
            TabIndex        =   35
            Top             =   480
            Width           =   705
         End
         Begin VB.CheckBox Chk27 
            Caption         =   "迅達"
            Height          =   255
            Index           =   3
            Left            =   2060
            TabIndex        =   33
            Top             =   480
            Width           =   850
         End
         Begin VB.CheckBox Chk27 
            Caption         =   "百靈"
            Height          =   255
            Index           =   4
            Left            =   2980
            TabIndex        =   34
            Top             =   480
            Width           =   850
         End
         Begin VB.CheckBox Chk27 
            Caption         =   "舜禹"
            Height          =   255
            Index           =   1
            Left            =   1140
            TabIndex        =   32
            Top             =   480
            Width           =   850
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   47
            Left            =   4680
            TabIndex        =   36
            Top             =   465
            Width           =   1200
            VariousPropertyBits=   671105051
            Size            =   "2117;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   270
            Index           =   2
            Left            =   2760
            TabIndex        =   31
            Top             =   152
            Width           =   495
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "命名人員"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   197
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "是否欲翻譯此案件者：　　　   (A:下班翻譯 B.上班翻譯)"
            Height          =   255
            Index           =   25
            Left            =   880
            TabIndex        =   164
            Top             =   165
            Width           =   4575
         End
         Begin VB.Label Label1 
            Caption         =   "指定翻譯："
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   163
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Height          =   720
         Left            =   120
         TabIndex        =   160
         Top             =   4395
         Width           =   5895
         Begin VB.OptionButton Opt5 
            Caption         =   "生醫"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   22
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Opt5 
            Caption         =   "化學"
            Height          =   255
            Index           =   1
            Left            =   2235
            TabIndex        =   23
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Opt5 
            Caption         =   "化工"
            Height          =   255
            Index           =   2
            Left            =   3165
            TabIndex        =   24
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Opt5 
            Caption         =   "材料"
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   25
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Opt5 
            Caption         =   "電子"
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   26
            Top             =   420
            Width           =   855
         End
         Begin VB.OptionButton Opt5 
            Caption         =   "機械"
            Height          =   255
            Index           =   5
            Left            =   2235
            TabIndex        =   27
            Top             =   420
            Width           =   855
         End
         Begin VB.OptionButton Opt5 
            Caption         =   "電機"
            Height          =   255
            Index           =   6
            Left            =   3165
            TabIndex        =   28
            Top             =   420
            Width           =   855
         End
         Begin VB.OptionButton Opt5 
            Caption         =   "其他"
            Height          =   255
            Index           =   7
            Left            =   4080
            TabIndex        =   29
            Top             =   420
            Width           =   675
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   23
            Left            =   4800
            TabIndex        =   30
            Top             =   405
            Width           =   855
            VariousPropertyBits=   671105051
            Size            =   "1508;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "案件類別："
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   161
            Top             =   160
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   800
         Left            =   120
         TabIndex        =   159
         Top             =   3600
         Width           =   5895
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   11
            Left            =   120
            TabIndex        =   199
            Top             =   130
            Width           =   4695
            Begin VB.CheckBox Chk2 
               Caption         =   "有序列表"
               Height          =   255
               Index           =   50
               Left            =   3120
               TabIndex        =   8
               Top             =   0
               Width           =   1215
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "可一案兩請"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   1215
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "彩圖提申"
               Height          =   255
               Index           =   49
               Left            =   1560
               TabIndex        =   7
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.CommandButton CmdFile 
            Caption         =   "上傳RES檔案"
            Height          =   495
            Left            =   4920
            TabIndex        =   12
            Top             =   240
            Width           =   900
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "本案說明書內容與"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   450
            Width           =   1750
         End
         Begin VB.Label Label4 
            Caption         =   "%相同"
            Height          =   252
            Index           =   6
            Left            =   4380
            TabIndex        =   237
            Top             =   456
            Width           =   564
         End
         Begin VB.Label Label4 
            Caption         =   "之內容"
            Height          =   228
            Index           =   4
            Left            =   3180
            TabIndex        =   236
            Top             =   480
            Width           =   612
         End
         Begin MSForms.TextBox txtData 
            Height          =   276
            Index           =   6
            Left            =   1872
            TabIndex        =   10
            Top             =   456
            Width           =   1284
            VariousPropertyBits=   671105051
            Size            =   "2265;487"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtData 
            Height          =   270
            Index           =   7
            Left            =   3840
            TabIndex        =   11
            Top             =   450
            Width           =   495
            VariousPropertyBits=   671105051
            Size            =   "1931;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   3315
         Width           =   5655
         Begin VB.OptionButton Opt1 
            Caption         =   "整體"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   2
            Top             =   0
            Width           =   800
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "部分"
            Height          =   255
            Index           =   1
            Left            =   2240
            TabIndex        =   3
            Top             =   0
            Width           =   800
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "圖像"
            Height          =   255
            Index           =   2
            Left            =   3160
            TabIndex        =   4
            Top             =   0
            Width           =   800
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "成組"
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   5
            Top             =   0
            Width           =   800
         End
         Begin VB.Label Label1 
            Caption         =   "設計案屬性"
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   188
            Top             =   0
            Width           =   975
         End
      End
      Begin MSForms.TextBox txtData 
         Height          =   285
         Index           =   5
         Left            =   1590
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   2970
         Width           =   7005
         VariousPropertyBits=   671105051
         Size            =   "12356;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtData 
         Height          =   285
         Index           =   4
         Left            =   1590
         TabIndex        =   1
         Top             =   2640
         Width           =   7005
         VariousPropertyBits=   671105051
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtData 
         Height          =   285
         Index           =   3
         Left            =   1590
         TabIndex        =   0
         Top             =   2325
         Width           =   7005
         VariousPropertyBits=   671105051
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCMboth 
         AutoSize        =   -1  'True
         Caption         =   "lblCMboth"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   198
         Top             =   2070
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "P.S.其他指定翻譯請在名稱後面加A或B"
         ForeColor       =   &H00FF00FF&
         Height          =   252
         Index           =   54
         Left            =   240
         TabIndex        =   196
         Top             =   6216
         Width           =   3252
      End
      Begin VB.Label Label5 
         Caption         =   "命名人員："
         Height          =   180
         Index           =   2
         Left            =   6120
         TabIndex        =   194
         Top             =   825
         Width           =   900
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   16
         Left            =   4010
         TabIndex        =   193
         Top             =   825
         Width           =   855
         ForeColor       =   255
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "1508;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   15
         Left            =   1200
         TabIndex        =   191
         Top             =   1125
         Width           =   1215
         ForeColor       =   255
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "2143;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   13
         Left            =   5580
         TabIndex        =   190
         Top             =   520
         Width           =   555
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "979;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   12
         Left            =   4680
         TabIndex        =   189
         Top             =   520
         Width           =   855
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "1508;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "印表機設定："
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   46
         Left            =   3600
         TabIndex        =   186
         Top             =   6252
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label1 
         Caption         =   "(外)："
         Height          =   255
         Index           =   7
         Left            =   1110
         TabIndex        =   157
         Top             =   2985
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "(英)："
         Height          =   255
         Index           =   6
         Left            =   1110
         TabIndex        =   156
         Top             =   2655
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱　 (中)："
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   155
         Top             =   2340
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "案件性質："
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   153
         Top             =   1125
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "譯畢期限："
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   152
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "總收文號："
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   151
         Top             =   1125
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   4
         Left            =   4005
         TabIndex        =   150
         Top             =   1125
         Width           =   1920
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "3387;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Caption         =   "收文日期："
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   149
         Top             =   1485
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   5
         Left            =   4010
         TabIndex        =   104
         Top             =   1485
         Width           =   855
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "1508;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   103
         Top             =   1800
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   280
         Index           =   6
         Left            =   4010
         TabIndex        =   102
         Top             =   1800
         Width           =   1455
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "2566;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "分案組別："
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   101
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   3
         Left            =   6120
         TabIndex        =   100
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "急件，請於　   　　　　　　　前譯畢名稱"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   99
         Top             =   540
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "國　　籍："
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   98
         Top             =   1485
         Width           =   1245
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   10
         Left            =   7380
         TabIndex        =   97
         Top             =   1485
         Width           =   1155
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "2037;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   8
         Left            =   7080
         TabIndex        =   96
         Top             =   840
         Width           =   1000
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "1764;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   9
         Left            =   7080
         TabIndex        =   95
         Top             =   1125
         Width           =   1000
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "1764;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "1.專利種類："
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   94
         Top             =   1485
         Width           =   1095
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   2
         Left            =   1400
         TabIndex        =   93
         Top             =   1485
         Width           =   1200
         ForeColor       =   255
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "2117;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "2.中說類型："
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   92
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblData 
         Height          =   280
         Index           =   3
         Left            =   1395
         TabIndex        =   46
         Top             =   1800
         Width           =   1560
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "2752;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "法定期限："
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   45
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "本所期限："
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   44
         Top             =   825
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   0
         Left            =   1200
         TabIndex        =   43
         Top             =   540
         Width           =   1000
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "1764;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   1
         Left            =   1200
         TabIndex        =   42
         Top             =   825
         Width           =   1000
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "1764;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   11
         Left            =   7080
         TabIndex        =   41
         Top             =   1800
         Width           =   1440
         BackColor       =   -2147483632
         VariousPropertyBits=   27
         Size            =   "2540;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "frm090903_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/27 改成Form2.0 ; lblData(index)、txtData(index)
'Created by Lydia 2017/11/14 外專新案未命名區-待命名明細
Option Explicit
Private Const m_FS As Integer = 16  '輸入欄位對應Table欄位的起始位置
'Modified by Lydia 2023/02/16 改成共用常數m_FE=>TF_TCT, m_NotFS=>TF_TCTnotFS
'Private Const m_FE As Integer = 119 '輸入欄位對應Table欄位的終止位置
'Private Const m_NotFS As String = "112,113,114,115" 'Added by Lydia 2018/03/01 排除不修改的欄位
'end 2023/02/16
Private Const m_Frame5 As Integer = 12 'Added by Lydia 2019/09/19 Frame5的index

Dim strPrinter As String
Dim m_PrevForm As Form '前一畫面
Dim oLbl As Control
Dim oTxt As Control
Dim oChk As CheckBox
Dim oOpt As OptionButton

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存案件名稱資料檔欄位的串列
Dim m_TCTList() As FIELDITEM
Dim m_TCTCount As Integer

Dim m_UserNo As String   '傳入員工編號
'Modified by Lydia 2018/04/18 +9-申請日,10-公告日(PA14),11-目前准/駁(PA16)
'Modified by Lydia 2020/02/21 +12-名稱有特殊字(PA174)
Dim strCase(1 To 12) As String '1~4本所案號pa01~pa04,5-專利種類pa08,6-申請國家pa09,7-分案組別pa150,8-設計案屬性pa158
Dim m_TCT01 As String  '收文號=PK
Dim m_TCT02 As String  '分案組別
Dim m_TCT04 As String  '工程師主管
Dim m_TCT07 As String  '工程師主任
Dim m_TCT10 As String  '命名人員編號
Dim m_TCT14 As String 'Added by Lydia 2019/09/19 重送確認記錄
Dim m_TCT20 As String 'Added by Lydia 2018/04/19 何時告代
Dim m_TCT117 As String 'Added by Lydia 2018/04/19 何時主動修正
Dim m_PA05 As String, m_PA06 As String  'Added by Lydia 2018/04/20 專利檔中文、英文名稱
Dim bolEmail As Boolean '是否發Email通知
Dim m_TCT27kind As String '欲翻譯此案件者可輸入的選項
Dim strCP203 As String, strCP901 As String '主動修正和告代的收文號
Dim m_UserSt16 As String 'Added by Lydia 2017/12/29 傳入員工編號的工程師組別
Dim n_CP118 As String, n_CP27 As String 'Added by Lydia 2018/04/18 新申請案：是否電子送件、發文日
Dim tCP06 As String, tCP27 As String 'Added by Lydia 2018/04/18 新案翻譯：所限、發文日
Dim m_CP203 As String 'Added by Lydia 2018/05/15 承辦收的A類主動修正
Dim m_TF01 As String, m_TF01pty As String 'Added by Lydia 2018/06/01 記錄中說收文號和案件性質
'Added by Lydia 2018/07/12
Dim m_TF01t As String, m_TF19 As String, m_TF20 As String, m_TF29 As String '翻譯費用檔的PK, 相似度,相似案號,待比對
Dim m_TF01cp14 As String, m_TF01cp27 As String '中說-承辦, 發文日
Dim m_GrpManList As String   '所有工程師主管(含F編號)
Public m_strSaveFiles As String '上傳檔案
Dim strResPath As String   '上傳相似比對結果存放路徑
'Added by Lydia 2018/10/18
Dim m_PA63 As String '客戶有提供彩圖(來自新案建檔)
Dim m_TCT01cp64 As String '新案收文號的進度備註
Dim bUpdCP64 As Boolean '黑白圖提申
'Added by Lydia 2020/02/21
Public bolAskPA174 As Boolean '存檔前檢查有修改案件名稱，將原始檔之維護word檔自動打開，是否有上傳
Dim cmdState As Integer
Dim strNotBList As String 'Added by Lydia 2023/01/18 命名作業不可新增告代和主動修正
Dim m_PA75 As String, m_PA26 As String, m_PA27 As String, m_PA28 As String, m_PA29 As String, m_PA30 As String  'Added by Lydia 2023/01/18 FC代理人和申請人1~5
Dim m_PA176 As String 'Added by Lydia 2023/03/10 是否新藥專利(P案的說明)/專利權期間延長相關 (FMP案的說明)
'Modified by Lydia 2025/06/05 更改名稱
'Dim m_strBASF As String 'Added by Lydia 2025/04/23 BASF集團的X編號--公告1120419-05
Dim m_str所內譯 As String
Dim m_str所內譯例外 As String 'Added by Lydia 2025/07/01

' 清除案件名稱資料檔檔案欄位串列
Private Sub ClearTCTFieldList()
   If m_TCTCount > 0 Then
      Erase m_TCTList
   End If
   m_TCTCount = 0
End Sub

' 設定案件名稱資料檔欄位串列中的欄位內容
Private Sub SetTCTFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TCTCount - 1
      If m_TCTList(nPos).fiName = strFieldName Then
         bFind = True
         m_TCTList(nPos).fiOldData = strFieldData
         m_TCTList(nPos).fiNewData = strFieldData
         m_TCTList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TCTList(m_TCTCount + 1)
      m_TCTList(m_TCTCount).fiName = strFieldName
      m_TCTList(m_TCTCount).fiOldData = strFieldData
      m_TCTList(m_TCTCount).fiNewData = strFieldData
      m_TCTList(m_TCTCount).fiType = nFieldType
      m_TCTCount = m_TCTCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetTCTFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TCTCount - 1
      If m_TCTList(nPos).fiName = strFieldName Then
         bFind = True
         m_TCTList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

Public Sub SetParent(ByRef fm As Form, ByVal pCase As String, ByVal pNo As String, ByVal pUser As String)
   Set m_PrevForm = fm
   m_TCT01 = pNo
   m_UserNo = pUser
   Call ChgCaseNo(Replace(pCase, "-", ""), strCase)
   m_UserSt16 = PUB_GetStaffST16(m_UserNo) 'Added by Lydia 2017/12/29
End Sub

Private Sub ClearForm(Optional ByVal bolRest As Boolean)
   
   ClearTCTFieldList
   
   For Each oLbl In lblData
      oLbl.Caption = ""
      If bolRest = True Then oLbl.BackColor = &H8000000F
   Next

   bolEmail = False
   
   For Each oTxt In txtData
      oTxt.Text = ""
      oTxt.Tag = ""
      oTxt.Visible = True
   Next
   For Each oChk In Chk2
      oChk.Value = 0
      oChk.Tag = ""
      oChk.Visible = True
   Next
   'Added by Lydia 2025/03/13
   For Each oChk In Chk27
      oChk.Value = 0
      oChk.Tag = ""
   Next
   'end 2025/03/13
   For Each oOpt In Opt1
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt3_1
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt3_2
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt3_3
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt3_4
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt3_5
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt3_6
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt4s
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   For Each oOpt In Opt5
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   'Added by Lydia 2018/04/18
   For Each oOpt In Opt4s2
      oOpt.Value = 0
      oOpt.Tag = ""
   Next
   
   'Added by Lydia 2018/07/12
   CmdFile.Visible = False
   m_strSaveFiles = ""
   
   'Added by Lydia 2018/10/22
   lblCMboth.Caption = ""
   lblCMboth.Tag = ""
   
   'Added by Lydia 2020/02/21
   ChkPA174.Value = vbUnchecked
   bolAskPA174 = False
   cmdState = -1
End Sub

Private Sub Chk2_Click(Index As Integer)

   If Chk2(Index).Value = vbChecked Then
      '說明書-內容項勾選
      If Index >= 3 And Index <= 19 And Chk2(2).Value = vbUnchecked Then
          Chk2(2).Value = vbChecked
      End If
      If Index >= 4 And Index <= 15 And Chk2(3).Value = vbUnchecked Then
          Chk2(3).Value = vbChecked
      End If
      '申請專利範圍-內容項勾選
      If Index >= 21 And Index <= 29 And Chk2(20).Value = vbUnchecked Then
          Chk2(20).Value = vbChecked
      End If
      '圖式-內容項勾選
      If Index >= 31 And Index <= 40 And Chk2(30).Value = vbUnchecked Then
          Chk2(30).Value = vbChecked
      End If
      If Index = 33 And Chk2(32).Value = vbUnchecked Then
          Chk2(32).Value = vbChecked
      End If
      '外翻
      'Mark by Lydia 2025/03/13 新增國外翻譯社
      'If Index >= 41 And Index <= 44 Then
      '    txtData(2).Text = ""
      '    For intI = 41 To 44
      '         If intI <> Index Then Chk2(intI).Value = vbUnchecked
      '    Next
      'End If
      'end 2025/03/13
      'Added by Lydia 2017/12/27 主動修正和告代改成Check
      If Index >= 45 And Index <= 47 Then
          Frame5(8).BackColor = &H8000000F
          Select Case Index
               Case 45 '收文主動修正
                    Chk2(46) = vbUnchecked
               Case 46 '不需收文
                    Chk2(45) = vbUnchecked
                    Chk2(47) = vbUnchecked
                    Chk2(48) = vbUnchecked 'Added by Lydia 2018/03/01
                    For Each oOpt In Opt4s
                         oOpt.Value = 0
                    Next
                    'Added by Lydia 2018/04/18
                    For Each oOpt In Opt4s2
                         oOpt.Value = 0
                    Next
               Case 47 '收文告代
                    Chk2(46) = vbUnchecked
          End Select
      End If
      'end 2017/12/27
      
      'Added by Lydia 2018/03/01 不請款
      If Index = 48 And Chk2(45).Value = vbUnchecked Then
          Chk2(45).Value = vbChecked
      End If
      'end 2018/03/01
      
      'Added by Lydia 2023/01/18 命名作業不可新增告代和主動修正
      If strNotBList <> "" And (Index = 45 Or Index = 47) Then
         If InStr(strNotBList, ",") > 0 Then
            strExc(9) = Mid(strNotBList, 1, InStr(strNotBList, ",") - 1)
         Else
            strExc(9) = strNotBList
         End If
         strExc(10) = Pub_GetITS01Type(strExc(9))
         'Added by Lydia 2025/10/02 針對Spruson & Ferguson (Asia)Pte Ltd (Y21071)之「Y01 命名作業」進行細部設定 (僅阻擋新案「235核對中說格式」或「209檢視中說」案件的「告代」與「主動修正」收文，並跳出說明提醒，但不阻擋其他如「201新案翻譯」、「210  製作中說」新案的「告代」與「主動修正」收文）)
         strExc(2) = ""
         If InStr("Y21071000,", m_PA75) > 0 And InStr("201,210", m_TF01pty) > 0 Then
            strExc(2) = "B"
         End If
         'end 2025/10/02
         If PUB_GetITStoList(Me.Name, strExc(10), strExc(9), False, False, , , "Y01") = True Then
         End If
         If strExc(2) = "" Then 'Added by Lydia 2025/10/02
            Chk2(46).Value = vbChecked
         End If 'Added by Lydia 2025/10/02
      End If
      'end 2023/01/18
   Else
       'Added by Lydia 2017/12/27 不收文告代
       If Index = 47 Then
          Frame5(9).BackColor = &H8000000F
          For Each oOpt In Opt4s
               oOpt.Value = 0
          Next
       End If
       'end 2017/12/27
       'Added by Lydia 2018/04/18 不收文主動修正
       If Index = 45 Then
          Frame5(10).BackColor = &H8000000F
          Chk2(48).Value = vbUnchecked
          For Each oOpt In Opt4s2
               oOpt.Value = 0
          Next
       End If
       'end 2018/04/18
   End If
End Sub

'Added by Lydia 2025/03/13
Private Sub Chk27_Click(Index As Integer)
   If Chk27(Index).Value = vbChecked Then
      txtData(2).Text = ""
      If Index <> 0 Then txtData(47) = ""
      For Each oChk In Chk27
         If oChk.Index <> Index Then
            oChk.Value = vbUnchecked
         End If
      Next
   End If
End Sub

Private Sub cmdExit_Click()
Dim bolCheck As Boolean

   If cmdOK(0).Enabled = True And cmdOK(1).Enabled = True Then
      If CheckDataDiff(bolCheck) Then
         If bolCheck = True Then
            If MsgBox("你並未存檔，確定離開嗎?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
      End If
   End If
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bDiff As Boolean
    
   cmdState = Index 'Added by Lydia 2020/02/21
   
   Select Case Index
      Case 0  '存檔
         'Modified by Lydia 2018/12/21 + bolJump
         'If TxtValidate1 = True Then
         If TxtValidate1(True) = True Then
            If CheckDataDiff(bDiff) = True Then
               'Modified by Lydia 2018/09/20 不修改,只上傳RES檔(相似比對結果)
               'If bDiff = True Then
               'Modified by Lydia 2018/10/18 +黑白圖提申
               'If bDiff = True Or (bDiff = False And m_strSaveFiles <> "") Then
               'Modified by Lydia 2020/03/03 +特殊字存檔後,繼續執行作業
               'If bDiff = True Or (bDiff = False And (m_strSaveFiles <> "" Or bUpdCP64 = True)) Then
               If bDiff = True Or (bDiff = False And ((m_strSaveFiles <> "" Or bUpdCP64 = True) Or (bolAskPA174 = True And ChkPA174.Value = vbChecked))) Then
                  'Modified by Lydia 2022/05/12 避免重複Click
                  'If OnSaveData(Index) = False Then Exit Sub
                  cmdOK(Index).Enabled = False
                  Screen.MousePointer = vbHourglass
                  If OnSaveData(Index) = False Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                  End If
                  Screen.MousePointer = vbDefault
                  'end 2022/05/12
                  ClearForm
                  If ReadData = True Then
                  End If
               End If
            End If
         End If
      Case 1  '確認回報
         If TxtValidate1 = True Then
            'Added by Lydia 2017/12/27
            If Trim(txtData(3)) = "待命名" And Trim(txtData(4)) = "" Then
                 MsgBox "案件名稱中文為待命名並且英文為空白時，不可確認回報!", vbExclamation
                 Exit Sub
            End If
            'end 2017/12/27
            'Added by Lydia 2021/05/06 命名作業若有指定翻譯人員，查詢該認領人員之新案翻譯未上完稿日案件
            If m_TF01pty = "201" And m_TF01cp14 = "" And Val(m_TF01cp27) = 0 And (txtData(2) <> "" Or txtData(47) <> "") Then
                 If txtData(2) = "A" Or txtData(2) = "B" Then
                     strExc(2) = m_TCT10
                 Else
                     strExc(2) = GetPrjSalesNM_2(Mid(txtData(47).Text, 1, Len(txtData(47).Text) - 1), , , , True)
                 End If
                 strExc(4) = Pub_GetEngEP09List(strExc(2))
                 If strExc(4) <> "" Then
                     If MsgBox(GetStaffName(strExc(2)) & " 尚未完稿案件：" & strExc(4) & vbCrLf & vbCrLf & "是否繼續確認？", vbInformation + vbYesNo + vbDefaultButton2, "翻譯人員檢查") = vbNo Then
                         Exit Sub
                     End If
                 End If
            End If
            'end 2021/05/06
            
            If CheckDataDiff(bDiff) = True Then
               'Modified by Lydia 2022/05/12 避免重複Click
               'If OnSaveData(Index) = False Then Exit Sub
               cmdOK(Index).Enabled = False
               Screen.MousePointer = vbHourglass
               If OnSaveData(Index) = False Then
                   Screen.MousePointer = vbDefault
                   Exit Sub
               End If
               Screen.MousePointer = vbDefault
               'end 2022/05/12
               Unload Me
            End If
         End If
      Case 2  '列印
         '先存檔,後列印
         If TxtValidate1 = True Then
            If CheckDataDiff(bDiff) = True Then
               If bDiff = True Then
                  If OnSaveData(Index) = False Then Exit Sub
                  ClearForm
                  If ReadData = True Then
                  End If
               End If
            End If
         'Modified by Lydia 2017/12/26 從下面移上來
         Call frm090902_2.PUB_PrintTCTcon(m_TCT01, Me.Combo1.Text, strPrinter)
         Unload frm090902_2
         End If
   End Select

End Sub

Private Sub Form_Load()
   ClearForm True
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, Me.Combo1, strPrinter, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數
   
   Frame1.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   Frame3.BackColor = &H8000000F
   Frame4.BackColor = &H8000000F
   'Modified by Lydia 2018/04/18 max=9 => 10
   'Modified by Lydia 2019/09/19 10=>改成m_Frame5
   For intI = 0 To m_Frame5
      Frame5(intI).BackColor = &H8000000F
   Next intI
   
   SSTab1.Tab = 0
   'Modified by Lydia 2018/04/26 +檢查名稱
   'If ReadData = True Then
   If ReadData(True) = True Then
   End If

   'Added by Lydia 2017/12/29 檢查專利檔和命名記錄檔的工程師組別
   'Mark by Lydia 2024/03/25 原因:機械組案件101,102由外專工程師(可能是電子組或化學組), 103由內專工程師翻譯
   ''If strCase(7) <> m_UserSt16 Then
   '     MsgBox "命名記錄檔的工程師組別和專利基本檔不一致，請通知程序人員到新案建檔設定工程師組別! "
   '     cmdOK(0).Enabled = False
   '     cmdOK(1).Enabled = False
   '     cmdOK(2).Enabled = False
   '     cmdOpen.Enabled = False
   'End If
   'end 2024/03/25
   'end 2017/12/29
   
   'Added by Lydia 2018/07/12
   m_GrpManList = Pub_GetSt16Man(True) '所有工程師主管(含F編號)
   strResPath = Pub_GetSpecMan("FCP相似比對結果暫存")
   
   Chk2(51).Top = 3345   'Added by Lydia 2023/03/10
   
   'Added by Lydia 2025/03/13 以後原本的4會改成Z
   If strSrvDate(1) < "20250314" Then Chk27(4).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

   PUB_SendMailCache 'Added by Lydia 2018/09/17
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
      
   Set frm090903_1 = Nothing
   If TypeName(m_PrevForm) <> "Nothing" Then
        m_PrevForm.Show
        If m_PrevForm.Name = "frm090903" Then
           m_PrevForm.doQuery False
        End If
   End If
   Set m_PrevForm = Nothing
   
End Sub

'Modified by Lydia 2018/04/26 +是否檢查中英文名稱bChk
Private Function ReadData(Optional bChk As Boolean = False) As Boolean
Dim rsRd As New ADODB.Recordset
  
    '先清除案件名稱資料檔串列
    ClearTCTFieldList
    
    '改成模組控制,若基本資料顯示有變,要注意frm090902_1,frm090902_2,frm090903_1的欄位
    'Modified by Lydia 2018/04/18
    'If PUB_GetTCTread(Me, strCase, m_TCT27kind) = True Then
    If PUB_GetTCTread(Me, strCase, m_TCT27kind, n_CP118, n_CP27, tCP06, tCP27) = True Then
       ReadData = True
       'Added by Lydia 2018/04/20 基本檔中文、英文名稱
       m_PA05 = txtData(3).Text
       m_PA06 = txtData(4).Text
       'end 2018/04/20
       Call SetCaseTitle
       'Added by Lydia 2020/02/21 基本檔：名稱有特殊字
       If strCase(12) = "Y" Then
          ChkPA174.Value = vbChecked
       End If
           
        'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地，外文本->原始檔區
        If PUB_ChkCPExist(strCase, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            cmdOpen.Caption = Replace(cmdOpen.Caption, "外文本", "原始檔")
            cmdOpen.Tag = strExc(1)
        End If
        'Mark by Lydia 2020/03/18 以收文為準
        'If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
        '    cmdOpen.Caption = Replace(cmdOpen.Caption, "外文本", "原始檔")
        'End If
        'end 2020/01/20
    Else
       MsgBox "查無資料 !", vbExclamation
       Unload Me
    End If
    'Added by Lydia 2018/04/20 檢查基本檔和命名記錄的名稱
    'Modified by Lydia 2018/04/26 櫃台現在會幫忙Key英文名稱,所以改成分開檢查
    'If m_PA05 & m_PA06 <> "待命名" And m_PA05 & m_PA06 <> txtData(3).Text & txtData(4).Text Then
    '    If MsgBox("命名作業的中英文名稱和專利基本檔不一致，是否代入專利基本檔的中英文名稱？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
    '                 txtData(3).Text = m_PA05
    '                 txtData(4).Text = m_PA06
    '    End If
    'End If
    If bChk = True Then
            If m_PA05 <> "待命名" And txtData(3).Text <> "待命名" And m_PA05 <> txtData(3).Text Then
                If MsgBox("命名作業的中文名稱和專利基本檔不一致，是否代入專利基本檔的中文名稱？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     txtData(3).Text = m_PA05
                End If
            End If
            If m_PA06 <> "" And txtData(4).Text <> "" And m_PA06 <> txtData(4).Text Then
                If MsgBox("命名作業的英文名稱和專利基本檔不一致，是否代入專利基本檔的英文名稱？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     txtData(4).Text = m_PA06
                End If
            End If
    End If
    'end 2018/04/26
   
    'Modified by Lydia 2025/06/05 更改名稱
    'm_strBASF = Pub_GetSpecMan("外專翻譯分案-BASF") & ","   'Added by Lydia 2025/04/23
    m_str所內譯 = Pub_GetSpecMan("外專翻譯分案-所內譯") & ","
    m_str所內譯例外 = Pub_GetSpecMan("外專翻譯分案-所內譯例外") & "," 'Added by Lydia 2025/07/01
    
End Function

'設案件命名欄位
Private Sub SetCaseTitle()
Dim rsA As New ADODB.Recordset
Dim Str01 As String
Dim intA As Integer
    
    'Modified by Lydia 2018/03/06
    'Str01 = "SELECT A.* FROM TRANSCASETITLE A WHERE TCT01='" & m_TCT01 & "' "
    'Modified by Lydia 2018/10/18 另外抓PA63,CP64
    'Str01 = "SELECT A.*,s1.st02 as TCT10n FROM TRANSCASETITLE A,staff s1 WHERE TCT01='" & m_TCT01 & "' and tct10=s1.st01(+) "
    'Modified by Lydia 2023/01/18 +pa75,pa26,pa27,pa28,pa29,pa30
    'Modified by Lydia 2023/03/10 + CP10,PA176
    Str01 = "SELECT A.*,s1.st02 as TCT10n,pa63,cp64,pa75,pa26,pa27,pa28,pa29,pa30,CP10,PA176 " & _
                "FROM TRANSCASETITLE A,staff s1,caseprogress,patent " & _
                "WHERE TCT01='" & m_TCT01 & "' and tct10=s1.st01(+) and tct01=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
    
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, Str01)
    If intA = 1 Then
        With rsA
           'Added by Lydia 2018/10/18
           m_PA63 = "" & rsA.Fields("pa63")  '客戶有提供彩圖
           m_TCT01cp64 = "" & rsA.Fields("cp64") '新案收文號的進度備註
           'end 2018/10/18
           
           '譯畢期限
           If "" & .Fields("TCT02") <> "" Then
              lblData(12).Caption = ChangeTStringToTDateString(TransDate(.Fields("TCT02"), 1))
              Label5(1).Visible = True
           Else
              Label5(1).Visible = False
           End If

           If "" & .Fields("TCT03") <> "" Then
              lblData(13).Caption = Format("" & .Fields("TCT03"), "00:00")
           End If
           '工程師主管
           m_TCT04 = "" & .Fields("TCT04")
           '工程師主任
           m_TCT07 = "" & .Fields("TCT07")
           '命名人員
           m_TCT10 = Trim("" & .Fields("TCT10"))
           lblData(8).Caption = Trim("" & .Fields("TCT10n")) 'Added by Lydia 2018/03/16 顯示命名人員
           m_TCT14 = "" & .Fields("TCT14")   'Added by Lydia 2019/09/19 重送確認記錄
           'Added by Lydia 2018/04/19 記錄告代和主動修正的狀態
           m_TCT20 = "" & .Fields("TCT20")
           m_TCT117 = "" & .Fields("TCT117")
           'Added by Lydia 2023/03/10 FMP大陸新藥發明專利權期限補償控管
           'FMP大陸發明案+化學組
           m_PA176 = "" & .Fields("PA176")
           If strCase(1) = "P" And strCase(6) = "020" And strCase(7) = "2" And "" & .Fields("cp10") = "101" Then
               Chk2(51).Visible = True
               If m_PA176 = "Y" And Chk2(51).Value = 0 Then
                   Chk2(51).Value = 1
               End If
           Else
               Chk2(51).Visible = False
           End If
           'end 2023/03/10
           
           For intA = m_FS To TF_TCT
              If InStr(TF_TCTnotFS, Format(intA, "000")) = 0 Then 'Added by Lydia 2018/03/01 加判斷
                    If GetDataDef("S", intA, "" & .Fields(intA - 1), .Fields(intA - 1).DefinedSize) Then
                      '儲存串列
                      SetTCTFieldOldData "TCT" & Format(intA, "00"), "" & .Fields(intA - 1), 0 '0-文字,1-數字
                    End If
              End If
           Next intA
           'Added by Lydia 2023/01/18 FC代理人和申請人1~5
           m_PA75 = "" & .Fields("PA75")
           m_PA26 = "" & .Fields("PA26")
           m_PA27 = "" & .Fields("PA27")
           m_PA28 = "" & .Fields("PA28")
           m_PA29 = "" & .Fields("PA29")
           m_PA30 = "" & .Fields("PA30")
           'end 2023/01/18
        End With
        
        'Added by Lydia 2018/05/15 承辦收的A類主動修正
        If PUB_ChkCPExist(strCase, "203", , m_CP203, , "A") = True Then
        End If
        
    End If
    
    'Added by Lydia 2018/06/01 記錄中說收文號和案件性質
    'Modified by Lydia 2018/07/12 +cp14,cp27,tf29
    'Str01 = "select cp09,cp10,tf01,tf19,tf20 from caseprogress,transfee " & _
                "where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' " & _
                "and cp10 in (" & GetAddStr(FcpTctPtys) & ") and cp159=0 and cp09=tf01(+) order by cp09 "
    Str01 = "select cp09,cp10,cp14,cp27,tf01,tf19,tf20,tf29 from caseprogress,transfee " & _
                "where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' " & _
                "and cp10 in ('201','209','235') and cp159=0 and cp09=tf01(+) order by cp09 "
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, Str01)
    If intA = 1 Then
         m_TF01 = "" & rsA.Fields("cp09")
         m_TF01pty = "" & rsA.Fields("cp10")
         'Added by Lydia 2018/07/12
         m_TF01cp14 = "" & rsA.Fields("cp14")
         m_TF01cp27 = "" & rsA.Fields("cp27")
         ' 翻譯費用檔
         m_TF01t = "" & rsA.Fields("tf01") 'PK
         m_TF19 = "" & rsA.Fields("tf19") '相似度
         If "" & rsA.Fields("tf20") <> "" Then
             'Modified by Lydia 2023/05/30 改成要輸入全部案號
             'm_TF20 = Mid("" & rsA.Fields("tf20"), 4, 6)  '相似案號(只取流水號)
             m_TF20 = rsA.Fields("tf20")
         Else
             m_TF20 = ""
         End If
         m_TF29 = "" & rsA.Fields("tf29") '待比對
         'end 2018/07/12
    End If
    'end 2018/06/01
    
    Set rsA = Nothing
    
    'Added by Lydia 2018/06/01 限新案翻譯才可輸入欲翻譯人員
    If m_TF01pty <> "201" Then
        Frame4.Enabled = False
    Else
        Frame4.Enabled = True
    End If
    'end 2018/06/01
    
    'Added by Lydia 2018/07/12 上傳相似比對結果檔案,新案翻譯才需要
    'Modified by Lydia 2018/08/09 工程師約定8/13上線
    'Modified by Lydia 2018/12/20 相似案號或相似度有值,就顯示
    'If strSrvDate(1) >= "20180813" And m_TF01 <> "" And m_TF01pty = "201" And txtData(6) <> "" And txtData(7) <> "" Then
    If strSrvDate(1) >= "20180813" And m_TF01 <> "" And m_TF01pty = "201" And (txtData(6) <> "" Or txtData(7) <> "") Then
        CmdFile.Visible = True
    End If
    
    'Added by Lydia 2023/04/26 (電子電機、化學、機械 三組)直接以彩圖製作成ori版本提申(自112年5月1日起實施); 日文組仍維持原程式=>需判斷是否以彩圖提申
    'Modified by Lydia 2024/10/17 debug: 工程師組別strCase(5)>>strCase(7)
    If strSrvDate(1) >= "20230501" And strCase(7) <> "3" And m_PA63 = "Y" Then
        Chk2(49).Value = vbChecked
    End If
    'end 2023/04/26
    
    'Added by Lydia 2023/01/18 命名作業不可新增告代和主動修正
    strNotBList = Pub_GetITSforHandle(strCase(1) & strCase(2) & strCase(3) & strCase(4), m_PA75, m_PA26 & "," & m_PA27 & "," & m_PA28 & "," & m_PA29 & "," & m_PA30)
        
End Sub

'檢查資料是否有變更
Private Function CheckDataDiff(ByRef bolDiff As Boolean) As Boolean
Dim inR As Integer
Dim tmpA1 As Variant 'Added by Lydia 2018/03/01

On Error GoTo Err02

    CheckDataDiff = False
    bolDiff = False
    tmpA1 = Empty
    tmpA1 = Split(TF_TCTnotFS, ",")
    For inR = m_FS To TF_TCT
       If InStr(TF_TCTnotFS, Format(inR, "000")) = 0 Then 'Added by Lydia 2018/03/01 加判斷
            If GetDataDef("U", inR) Then
               If bolDiff = False Then
                  'Added by Lydia 2018/03/01 跳過UpdateID
                  If inR - m_FS < m_TCTCount - 1 Then
                        If m_TCTList(inR - m_FS).fiNewData <> m_TCTList(inR - m_FS).fiOldData Then
                           bolDiff = True
                        End If
                  Else
                        If m_TCTList(inR - m_FS - (UBound(tmpA1) + 1)).fiNewData <> m_TCTList(inR - m_FS - (UBound(tmpA1) + 1)).fiOldData Then
                           bolDiff = True
                        End If
                  End If
                  'end 2018/03/01
               End If
            End If
       End If
    Next inR
    
    CheckDataDiff = True
Err02:
    Exit Function
End Function

'設定資料和欄位顯示
Private Function GetDataDef(ByVal nType As String, ByVal nIndex As Integer, Optional ByVal nData As String = "", Optional ByVal nMax As Integer = 0) As Boolean
'nType = S(預設) ,U(取得新值), W(欄位描述)
'nMax =輸入欄位最大長度
Dim strM01 As String
Dim intM As Integer
Dim iChar As Integer

'因為O12的欄位有設char,會=O8欄位長度*4
'Modified by Lydia 2022/03/22 已改用新的Provider,可以直接取得真實長度
'If InStr(UCase(Forms(0).Caption), "O12") > 0 Then
'   iChar = 4
'Else
'   iChar = 1
'End If
 iChar = 1
 'end 2022/03/22
 
On Error GoTo Err01
   Select Case nIndex
       Case 16 '案件名稱(中)
           If nType = "S" Then
              If nData <> "" Then
                 txtData(3).Text = nData
              '案件名稱(中)預設'待命名'但可修改
              ElseIf Trim(txtData(3).Text) = "" Then
                 txtData(3).Text = "待命名"
              End If
              If nMax > 0 Then
                 txtData(3).Tag = "TCT" & Format(nIndex, "00")
                 txtData(3).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(3).Tag, txtData(3).Text
           ElseIf nType = "W" Then
           End If
       Case 17 '案件名稱(英)
           If nType = "S" Then
              If nData <> "" Then
                 txtData(4).Text = nData
              End If
              If nMax > 0 Then
                 txtData(4).Tag = "TCT" & Format(nIndex, "00")
                 txtData(4).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(4).Tag, txtData(4).Text
           ElseIf nType = "W" Then
           End If
       Case 18 '設計案屬性
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 4 Then
                 Opt1(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt1(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt1
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt1(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       'Modified by Lydia 2018/03/01 +不請款(116)
       'Modified by Lydia 2018/04/18 +提申後/提申前(117)
       Case 19, 116, 117 '收文主動修正203
           If nIndex = 19 Then 'Added by Lydia 2018/03/01 判斷
                If nType = "S" Then
                   If nData = "Y" Then
                       Chk2(45).Value = vbChecked
                   ElseIf nData = "N" Then
                       Chk2(46).Value = vbChecked
                   End If
                   If nMax > 0 Then
                      Chk2(45).Tag = "TCT" & Format(nIndex, "00")
                   End If
                ElseIf nType = "U" Then
                   strM01 = ""
                   If Chk2(45).Value = vbChecked Then
                        strM01 = "Y"
                   ElseIf Chk2(46).Value = vbChecked Then
                        strM01 = "N"
                   End If
                   SetTCTFieldNewData Chk2(45).Tag, strM01
                ElseIf nType = "W" Then
                End If
           'Added by Lydia 2018/04/18 提申後/提申前(117)
           ElseIf nIndex = 117 Then
                If nType = "S" Then
                   If Val(nData) >= 1 And Val(nData) <= 2 Then
                      Opt4s2(Val(nData) - 1).Value = 1
                      Chk2(45).Value = vbChecked
                   End If
                   If nMax > 0 Then
                      Opt4s2(0).Tag = "TCT" & Format(nIndex, "00")
                   End If
                ElseIf nType = "U" Then
                   intM = 0
                   For Each oOpt In Opt4s2
                      If oOpt.Value = True Then intM = oOpt.Index + 1
                   Next
                   SetTCTFieldNewData Opt4s2(0).Tag, IIf(intM > 0, intM, "")
                ElseIf nType = "W" Then
                End If
           'end 2018/04/18
           Else '是否不請款
                If nType = "S" Then
                   If nData = "Y" Then
                      Chk2(48).Value = vbChecked
                   Else
                      Chk2(48).Value = vbUnchecked
                   End If
                   If nMax > 0 Then
                      Chk2(48).Tag = "TCT" & Format(nIndex, "00")
                   End If
                ElseIf nType = "U" Then
                   strM01 = ""
                   If Chk2(48).Value = vbChecked Then
                      strM01 = "Y"
                   End If
                   SetTCTFieldNewData Chk2(48).Tag, strM01
                ElseIf nType = "W" Then
                End If
           End If
       Case 20 '告代(告知代理人901)
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 3 Then
                 Opt4s(Val(nData) - 1).Value = 1
                 Chk2(47).Value = vbChecked
              End If
              If nMax > 0 Then
                 Opt4s(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt4s
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt4s(0).Tag, IIf(intM > 0, intM, "")
           ElseIf nType = "W" Then
           End If
       Case 21 '一案兩請
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(0).Value = vbChecked
              Else
                 Chk2(0).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(0).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(0).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 22 '有相似案
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(1).Value = vbChecked
              Else
                 Chk2(1).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(1).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(1).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(1).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 23 '相似案號
           If nType = "S" Then
              If nData <> "" Then
                 'Modified by Lydia 2023/05/30 改成要輸入全部案號
                 'txtData(6).Text = Mid(nData, 4, 6)
                 txtData(6).Text = nData
              End If
              If nMax > 0 Then
                 txtData(6).Tag = "TCT" & Format(nIndex, "00")
                 'Modified by Lydia 2023/05/30
                 'txtData(6).MaxLength = 6
                 txtData(6).MaxLength = 12
              End If
           ElseIf nType = "U" Then
              'Modified by Lydia 2023/05/30 改成要輸入全部案號
              'strM01 = ""
              'If Trim(txtData(6).Text) <> "" Then
              '   strM01 = "FCP" & txtData(6).Text & "000"
              'End If
              'SetTCTFieldNewData txtData(6).Tag, strM01
              SetTCTFieldNewData txtData(6).Tag, txtData(6).Text
           ElseIf nType = "W" Then
           End If
       Case 24 '相似內容%
           If nType = "S" Then
              If nData <> "" Then
                 txtData(7).Text = nData
              End If
              If nMax > 0 Then
                 txtData(7).Tag = "TCT" & Format(nIndex, "00")
                 txtData(7).MaxLength = nMax
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(7).Tag, txtData(7).Text
           ElseIf nType = "W" Then
           End If
       Case 25  '案件類別
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 8 Then
                 Opt5(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt5(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt5
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt5(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       Case 26  '其他類別內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(23).Text = nData
              End If
              If nMax > 0 Then
                 txtData(23).Tag = "TCT" & Format(nIndex, "00")
                 txtData(23).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(23).Tag, txtData(23).Text
           ElseIf nType = "W" Then
           End If
       Case 27  '欲翻譯此案件者/指定翻譯
           If nType = "S" Then
              If nData = "A" Or nData = "B" Then
                 txtData(2).Text = nData
              'Modified by Lydia 2025/03/13 新增國外翻譯社
              'ElseIf Val(nData) >= 1 And Val(nData) <= 4 Then
              '   Chk2(40 + Val(nData)).Value = vbChecked
              Else
                 If nData = "Z" Or (strSrvDate(1) < "20250314" And nData = "4") Then '114/3/14以後原本的4會改成Z
                     Chk27(0).Value = vbChecked
                 ElseIf Val(nData) >= 1 Then
                     Chk27(Val(nData)).Value = vbChecked
                 End If
              'end 2025/03/13
              End If
              If nMax > 0 Then
                 txtData(2).Tag = "TCT" & Format(nIndex, "00")
                 txtData(2).MaxLength = nMax
              End If
           ElseIf nType = "U" Then
              strM01 = txtData(2).Text
              'Modified by Lydia 2025/03/13 新增國外翻譯社
              'intM = 0
              'For intI = 41 To 44
              '   If Chk2(intI).Value = vbChecked Then intM = intI - 40
              'Next
              intM = -1
              For Each oChk In Chk27
                 If oChk.Value = vbChecked Then
                    If oChk.Index = 0 Then
                       strM01 = "Z"
                    Else
                       intM = oChk.Index
                    End If
                 End If
              Next
              'end 2025/03/13
              SetTCTFieldNewData txtData(2).Tag, IIf(intM = 0 And strM01 = "", Empty, IIf(intM > 0, intM, strM01))
           ElseIf nType = "W" Then
           End If
       Case 28  '其他指定翻譯
           If nType = "S" Then
              If nData <> "" Then
                 txtData(47).Text = nData
              End If
              If nMax > 0 Then
                 txtData(47).Tag = "TCT" & Format(nIndex, "00")
                 txtData(47).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(47).Tag, txtData(47).Text
           ElseIf nType = "W" Then
           End If
       Case 29  '說明書
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(2).Value = vbChecked
              Else
                 Chk2(2).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(2).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(2).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(2).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 30  '內容不完整
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(3).Value = vbChecked
              Else
                 Chk2(3).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(3).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(3).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(3).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 31  '技術領域標題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(4).Value = vbChecked
              Else
                 Chk2(4).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(4).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(4).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(4).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 32  '技術領域標題內容
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 2 Then
                 Opt3_1(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt3_1(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt3_1
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt3_1(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       Case 33  '技術領域標題位置
           If nType = "S" Then
              If nData <> "" Then
                 txtData(8).Text = nData
              End If
              If nMax > 0 Then
                 txtData(8).Tag = "TCT" & Format(nIndex, "00")
                 txtData(8).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(8).Tag, txtData(8).Text
           ElseIf nType = "W" Then
           End If
       Case 34  '技術領域建議
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(5).Value = vbChecked
              Else
                 Chk2(5).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(5).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(5).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(5).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 35  '技術領域建議內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(9).Text = nData
              End If
              If nMax > 0 Then
                 txtData(9).Tag = "TCT" & Format(nIndex, "00")
                 txtData(9).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(9).Tag, txtData(9).Text
           ElseIf nType = "W" Then
           End If
       Case 36  '先前技術標題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(6).Value = vbChecked
              Else
                 Chk2(6).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(6).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(6).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(6).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 37  '先前技術標題內容
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 2 Then
                 Opt3_2(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt3_2(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt3_2
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt3_2(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       Case 38  '先前技術標題位置
           If nType = "S" Then
              If nData <> "" Then
                 txtData(10).Text = nData
              End If
              If nMax > 0 Then
                 txtData(10).Tag = "TCT" & Format(nIndex, "00")
                 txtData(10).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(10).Tag, txtData(10).Text
           ElseIf nType = "W" Then
           End If
       Case 39  '先前技術建議
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(7).Value = vbChecked
              Else
                 Chk2(7).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(7).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(7).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(7).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 40  '先前技術建議內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(11).Text = nData
              End If
              If nMax > 0 Then
                 txtData(11).Tag = "TCT" & Format(nIndex, "00")
                 txtData(11).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(11).Tag, txtData(11).Text
           ElseIf nType = "W" Then
           End If
       Case 41  '發明內容標題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(8).Value = vbChecked
              Else
                 Chk2(8).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(8).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(8).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(8).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 42  '發明內容標題內容
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 2 Then
                 Opt3_3(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt3_3(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt3_3
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt3_3(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       Case 43  '發明內容標題位置
           If nType = "S" Then
              If nData <> "" Then
                 txtData(12).Text = nData
              End If
              If nMax > 0 Then
                 txtData(12).Tag = "TCT" & Format(nIndex, "00")
                 txtData(12).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(12).Tag, txtData(12).Text
           ElseIf nType = "W" Then
           End If
       Case 44  '發明內容建議
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(9).Value = vbChecked
              Else
                 Chk2(9).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(9).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(9).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(9).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 45  '發明內容建議內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(13).Text = nData
              End If
              If nMax > 0 Then
                 txtData(13).Tag = "TCT" & Format(nIndex, "00")
                 txtData(13).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(13).Tag, txtData(13).Text
           ElseIf nType = "W" Then
           End If
       Case 46  '圖式簡單說明標題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(10).Value = vbChecked
              Else
                 Chk2(10).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(10).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(10).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(10).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 47  '圖式簡單說明標題內容
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 2 Then
                 Opt3_4(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt3_4(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt3_4
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt3_4(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       Case 48  '圖式簡單說明標題位置
           If nType = "S" Then
              If nData <> "" Then
                 txtData(14).Text = nData
              End If
              If nMax > 0 Then
                 txtData(14).Tag = "TCT" & Format(nIndex, "00")
                 txtData(14).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(14).Tag, txtData(14).Text
           ElseIf nType = "W" Then
           End If
       Case 49  '圖式簡單說明建議
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(11).Value = vbChecked
              Else
                 Chk2(11).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(11).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(11).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(11).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 50  '圖式簡單說明建議內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(15).Text = nData
              End If
              If nMax > 0 Then
                 txtData(15).Tag = "TCT" & Format(nIndex, "00")
                 txtData(15).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(15).Tag, txtData(15).Text
           ElseIf nType = "W" Then
           End If
       Case 51  '實施方式標題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(12).Value = vbChecked
              Else
                 Chk2(12).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(12).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(12).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(12).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 52  '實施方式標題內容
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 2 Then
                 Opt3_5(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt3_5(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt3_5
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt3_5(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       Case 53  '實施方式標題位置
           If nType = "S" Then
              If nData <> "" Then
                 txtData(16).Text = nData
              End If
              If nMax > 0 Then
                 txtData(16).Tag = "TCT" & Format(nIndex, "00")
                 txtData(16).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(16).Tag, txtData(16).Text
           ElseIf nType = "W" Then
           End If
       Case 54  '實施方式建議
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(13).Value = vbChecked
              Else
                 Chk2(13).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(13).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(13).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(13).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 55  '實施方式建議內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(17).Text = nData
              End If
              If nMax > 0 Then
                 txtData(17).Tag = "TCT" & Format(nIndex, "00")
                 txtData(17).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(17).Tag, txtData(17).Text
           ElseIf nType = "W" Then
           End If
       Case 56  '符號說明標題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(14).Value = vbChecked
              Else
                 Chk2(14).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(14).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(14).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(14).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 57  '符號說明標題內容
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 2 Then
                 Opt3_6(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 Opt3_6(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In Opt3_6
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData Opt3_6(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
           End If
       Case 58  '符號說明標題位置
           If nType = "S" Then
              If nData <> "" Then
                 txtData(18).Text = nData
              End If
              If nMax > 0 Then
                 txtData(18).Tag = "TCT" & Format(nIndex, "00")
                 txtData(18).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(18).Tag, txtData(18).Text
           ElseIf nType = "W" Then
           End If
       Case 59  '符號說明建議
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(15).Value = vbChecked
              Else
                 Chk2(15).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(15).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(15).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(15).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 60  '符號說明建議內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(19).Text = nData
              End If
              If nMax > 0 Then
                 txtData(19).Tag = "TCT" & Format(nIndex, "00")
                 txtData(19).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(19).Tag, txtData(19).Text
           ElseIf nType = "W" Then
           End If
       Case 61  '缺摘要
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(16).Value = vbChecked
              Else
                 Chk2(16).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(16).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(16).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(16).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 62  '缺摘要建議內容
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(17).Value = vbChecked
              Else
                 Chk2(17).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(17).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(17).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(17).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 63  '缺摘要建議內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(20).Text = nData
              End If
              If nMax > 0 Then
                 txtData(20).Tag = "TCT" & Format(nIndex, "00")
                 txtData(20).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(20).Tag, txtData(20).Text
           ElseIf nType = "W" Then
           End If
       Case 64  '缺頁
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(18).Value = vbChecked
              Else
                 Chk2(18).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(18).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(18).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(18).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 65  '頁數(描述)
           If nType = "S" Then
              If nData <> "" Then
                 txtData(21).Text = nData
              End If
              If nMax > 0 Then
                 txtData(21).Tag = "TCT" & Format(nIndex, "00")
                 txtData(21).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(21).Tag, txtData(21).Text
           ElseIf nType = "W" Then
           End If
       Case 66  '其它問題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(19).Value = vbChecked
              Else
                 Chk2(19).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(19).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(19).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(19).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 67  '其它問題內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(22).Text = nData
              End If
              If nMax > 0 Then
                 txtData(22).Tag = "TCT" & Format(nIndex, "00")
                 txtData(22).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(22).Tag, txtData(22).Text
           ElseIf nType = "W" Then
           End If
       Case Else '避免程式過長,分成2段
           If GetDataDef2(nType, nIndex, iChar, nData, nMax) = False Then
              Exit Function
           End If
   End Select
   
   GetDataDef = True
   Exit Function
   
Err01:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

Private Function GetDataDef2(ByVal nType As String, ByVal nIndex As Integer, ByVal iChar As Integer, Optional ByVal nData As String = "", Optional ByVal nMax As Integer = 0) As Boolean
'nType = S(預設) ,U(取得新值), W(欄位描述)
'nMax =輸入欄位最大長度
Dim strM01 As String
Dim intM As Integer

On Error GoTo Err01_2

   GetDataDef2 = False
   Select Case nIndex
       Case 68  '申請專利範圍
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(20).Value = vbChecked
              Else
                 Chk2(20).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(20).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(20).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(20).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 69  '項號錯誤
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(21).Value = vbChecked
              Else
                 Chk2(21).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(21).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(21).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(21).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 70  '項號錯誤請求項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(24).Text = nData
              End If
              If nMax > 0 Then
                 txtData(24).Tag = "TCT" & Format(nIndex, "00")
                 txtData(24).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(24).Tag, txtData(24).Text
           ElseIf nType = "W" Then
           End If
       Case 71  '依附關係錯誤
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(22).Value = vbChecked
              Else
                 Chk2(22).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(22).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(22).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(22).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 72  '依附關係錯誤附屬項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(25).Text = nData
              End If
              If nMax > 0 Then
                 txtData(25).Tag = "TCT" & Format(nIndex, "00")
                 txtData(25).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(25).Tag, txtData(25).Text
           ElseIf nType = "W" Then
           End If
       Case 73  '依附關係錯誤請求項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(26).Text = nData
              End If
              If nMax > 0 Then
                 txtData(26).Tag = "TCT" & Format(nIndex, "00")
                 txtData(26).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(26).Tag, txtData(26).Text
           ElseIf nType = "W" Then
           End If
       Case 74  '依附關係不明確
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(23).Value = vbChecked
              Else
                 Chk2(23).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(23).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(23).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(23).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 75  '依附關係不明確附屬項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(27).Text = nData
              End If
              If nMax > 0 Then
                 txtData(27).Tag = "TCT" & Format(nIndex, "00")
                 txtData(27).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(27).Tag, txtData(27).Text
           ElseIf nType = "W" Then
           End If
       Case 76  '依附關係不明確請求項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(28).Text = nData
              End If
              If nMax > 0 Then
                 txtData(28).Tag = "TCT" & Format(nIndex, "00")
                 txtData(28).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(28).Tag, txtData(28).Text
           ElseIf nType = "W" Then
           End If
       Case 77  '多附多 'Memo by Lydia 2018/04/17 原"不當依附"
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(24).Value = vbChecked
              Else
                 Chk2(24).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(24).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(24).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(24).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 78  '多附多附屬項 'Memo by Lydia 2018/04/17 原"不當依附"
           If nType = "S" Then
              If nData <> "" Then
                 txtData(29).Text = nData
              End If
              If nMax > 0 Then
                 txtData(29).Tag = "TCT" & Format(nIndex, "00")
                 txtData(29).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(29).Tag, txtData(29).Text
           ElseIf nType = "W" Then
           End If
       Case 79  '多附多請求項 'Memo by Lydia 2018/04/17 原"不當依附"
           If nType = "S" Then
              If nData <> "" Then
                 txtData(30).Text = nData
              End If
              If nMax > 0 Then
                 txtData(30).Tag = "TCT" & Format(nIndex, "00")
                 txtData(30).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(30).Tag, txtData(30).Text
           ElseIf nType = "W" Then
           End If
       Case 80  '引用記載形式
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(25).Value = vbChecked
              Else
                 Chk2(25).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(25).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(25).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(25).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 81  '引用記載形式請求項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(31).Text = nData
              End If
              If nMax > 0 Then
                 txtData(31).Tag = "TCT" & Format(nIndex, "00")
                 txtData(31).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(31).Tag, txtData(31).Text
           ElseIf nType = "W" Then
           End If
       Case 82  '標的不一致
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(26).Value = vbChecked
              Else
                 Chk2(26).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(26).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(26).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(26).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 83  '標的不一致附屬項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(32).Text = nData
              End If
              If nMax > 0 Then
                 txtData(32).Tag = "TCT" & Format(nIndex, "00")
                 txtData(32).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(32).Tag, txtData(32).Text
           ElseIf nType = "W" Then
           End If
       Case 84  '標的不一致附屬項標的
           If nType = "S" Then
              If nData <> "" Then
                 txtData(33).Text = nData
              End If
              If nMax > 0 Then
                 txtData(33).Tag = "TCT" & Format(nIndex, "00")
                 txtData(33).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(33).Tag, txtData(33).Text
           ElseIf nType = "W" Then
           End If
       Case 85  '標的不一致請求項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(34).Text = nData
              End If
              If nMax > 0 Then
                 txtData(34).Tag = "TCT" & Format(nIndex, "00")
                 txtData(34).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(34).Tag, txtData(34).Text
           ElseIf nType = "W" Then
           End If
       Case 86  '標的不一致請求項標的
           If nType = "S" Then
              If nData <> "" Then
                 txtData(35).Text = nData
              End If
              If nMax > 0 Then
                 txtData(35).Tag = "TCT" & Format(nIndex, "00")
                 txtData(35).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(35).Tag, txtData(35).Text
           ElseIf nType = "W" Then
           End If
       Case 87  '不予專利
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(27).Value = vbChecked
              Else
                 Chk2(27).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(27).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(27).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(27).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 88  '不予專利請求項
           If nType = "S" Then
              If nData <> "" Then
                 txtData(36).Text = nData
              End If
              If nMax > 0 Then
                 txtData(36).Tag = "TCT" & Format(nIndex, "00")
                 txtData(36).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(36).Tag, txtData(36).Text
           ElseIf nType = "W" Then
           End If
       Case 89  '不予專利請求項標的
           If nType = "S" Then
              If nData <> "" Then
                 txtData(37).Text = nData
              End If
              If nMax > 0 Then
                 txtData(37).Tag = "TCT" & Format(nIndex, "00")
                 txtData(37).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(37).Tag, txtData(37).Text
           ElseIf nType = "W" Then
           End If
       Case 90  '不予專利請求項法條
           If nType = "S" Then
              If nData <> "" Then
                 txtData(38).Text = nData
              End If
              If nMax > 0 Then
                 txtData(38).Tag = "TCT" & Format(nIndex, "00")
                 txtData(38).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(38).Tag, txtData(38).Text
           ElseIf nType = "W" Then
           End If
       Case 91  '混雜式請求項
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(28).Value = vbChecked
              Else
                 Chk2(28).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(28).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(28).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(28).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 92  '混雜式請求項內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(39).Text = nData
              End If
              If nMax > 0 Then
                 txtData(39).Tag = "TCT" & Format(nIndex, "00")
                 txtData(39).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(39).Tag, txtData(39).Text
           ElseIf nType = "W" Then
           End If
       Case 93  '其它問題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(29).Value = vbChecked
              Else
                 Chk2(29).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(29).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(29).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(29).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 94  '其它問題內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(40).Text = nData
              End If
              If nMax > 0 Then
                 txtData(40).Tag = "TCT" & Format(nIndex, "00")
                 txtData(40).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(40).Tag, txtData(40).Text
           ElseIf nType = "W" Then
           End If
       Case 95  '圖式
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(30).Value = vbChecked
              Else
                 Chk2(30).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(30).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(30).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(30).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 96  '建議指定代表圖
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(31).Value = vbChecked
              Else
                 Chk2(31).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(31).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(31).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(31).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 97  '建議指定代表圖內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(41).Text = nData
              End If
              If nMax > 0 Then
                 txtData(41).Tag = "TCT" & Format(nIndex, "00")
                 txtData(41).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(41).Tag, txtData(41).Text
           ElseIf nType = "W" Then
           End If
       Case 98  '缺圖
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(32).Value = vbChecked
              Else
                 Chk2(32).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(32).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(32).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(32).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 99  '彩圖
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(33).Value = vbChecked
              Else
                 Chk2(33).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(33).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(33).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(33).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 100 '缺圖內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(42).Text = nData
              End If
              If nMax > 0 Then
                 txtData(42).Tag = "TCT" & Format(nIndex, "00")
                 txtData(42).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(42).Tag, txtData(42).Text
           ElseIf nType = "W" Then
           End If
       Case 101 '格式不符
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(34).Value = vbChecked
              Else
                 Chk2(34).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(34).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(34).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(34).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 102 '格式不符圖內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(43).Text = nData
              End If
              If nMax > 0 Then
                 txtData(43).Tag = "TCT" & Format(nIndex, "00")
                 txtData(43).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(43).Tag, txtData(43).Text
           ElseIf nType = "W" Then
           End If
       Case 103 '格式不符說明
           If nType = "S" Then
              If nData <> "" Then
                 txtData(44).Text = nData
              End If
              If nMax > 0 Then
                 txtData(44).Tag = "TCT" & Format(nIndex, "00")
                 txtData(44).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(44).Tag, txtData(44).Text
           ElseIf nType = "W" Then
           End If
       Case 104 '其它問題
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(40).Value = vbChecked
              Else
                 Chk2(40).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(40).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(40).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(40).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 105 '其它問題內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(46).Text = nData
              End If
              If nMax > 0 Then
                 txtData(46).Tag = "TCT" & Format(nIndex, "00")
                 txtData(46).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(46).Tag, txtData(46).Text
           ElseIf nType = "W" Then
           End If
       Case 106 '不完整
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(35).Value = vbChecked
              Else
                 Chk2(35).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(35).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(35).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(35).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 107 '不完整圖內容
           If nType = "S" Then
              If nData <> "" Then
                 txtData(45).Text = nData
              End If
              If nMax > 0 Then
                 txtData(45).Tag = "TCT" & Format(nIndex, "00")
                 txtData(45).MaxLength = nMax / iChar
              End If
           ElseIf nType = "U" Then
              SetTCTFieldNewData txtData(45).Tag, txtData(45).Text
           ElseIf nType = "W" Then
           End If
       Case 108 '超過一個實施例
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(36).Value = vbChecked
              Else
                 Chk2(36).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(36).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(36).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(36).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 109 '不主張設計的部分
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(37).Value = vbChecked
              Else
                 Chk2(37).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(37).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(37).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(37).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 110 '色彩
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(38).Value = vbChecked
              Else
                 Chk2(38).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(38).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(38).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(38).Tag, strM01
           ElseIf nType = "W" Then
           End If
       Case 111 '用途說明
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(39).Value = vbChecked
              Else
                 Chk2(39).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(39).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(39).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(39).Tag, strM01
           ElseIf nType = "W" Then
           End If
       'Added by Lydia 2018/04/18
       Case 118 '彩圖提申
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(49).Value = vbChecked
              Else
                 Chk2(49).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(49).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(49).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(49).Tag, strM01
           ElseIf nType = "W" Then
           End If
       'end 2018/04/18
       'Added by Lydia 2021/04/09
       Case 119  '有序列表
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(50).Value = vbChecked
              Else
                 Chk2(50).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(50).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(50).Value = vbChecked Then
                 strM01 = "Y"
              End If
              SetTCTFieldNewData Chk2(50).Tag, strM01
           ElseIf nType = "W" Then
           End If
       'end 2021/04/09
       'Added by Lydia 2023/03/10
       Case 120  '專利權期間延長相關
           If nType = "S" Then
              If nData = "Y" Then
                 Chk2(51).Value = vbChecked
              Else
                 Chk2(51).Value = vbUnchecked
              End If
              If nMax > 0 Then
                 Chk2(51).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              strM01 = ""
              If Chk2(51).Visible = True Then
                 If Chk2(51).Value = vbChecked Then
                    strM01 = "Y"
                 Else
                    strM01 = "N"
                 End If
                 SetTCTFieldNewData Chk2(51).Tag, strM01
              End If
           ElseIf nType = "W" Then
           End If
       'end 2023/03/10
       Case Else
   End Select
   
   GetDataDef2 = True
   Exit Function
   
Err01_2:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

Private Function OnSaveData(ByVal iType As Integer) As Boolean
Dim strTmp As String
Dim strExSql As String
Dim nIndex As Integer
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim tmpArr As Variant 'Added by Lydia 2018/09/27

   'Added by Lydia 2018/07/12 上傳檔案
   If m_strSaveFiles <> "" Then
      'Modified by Lydia 2018/09/27 開放可上傳多個檔案
      'strExc(3) = m_strSaveFiles
      tmpArr = Empty
      tmpArr = Split(m_strSaveFiles, "&")
      For nIndex = 0 To UBound(tmpArr)
          If Trim("" & tmpArr(nIndex)) <> "" Then
                strExc(3) = Trim(tmpArr(nIndex))
      'end 2018/09/27
                If InStr(strExc(3), " (") > 0 Then strExc(3) = RTrim(Mid(strExc(3), 1, InStr(strExc(3), " (")))
                If Dir(strExc(3)) <> "" Then '檔案路徑正確
                      'Modified by Lydia 2018/09/27 用原檔名
                      'If Pub_FtpPutTyping2(strExc(3), strResPath & "/" & strCase(1) & strCase(2) & ".RES" & Mid(strExc(3), InStrRev(strExc(3), "."))) = False Then
                      If Pub_FtpPutTyping2(strExc(3), strResPath & "/" & Mid(strExc(3), InStrRev(strExc(3), "\") + 1)) = False Then
                          Exit Function
                      End If
                End If
      'Added by Lydia 2018/09/27
          End If
      Next nIndex
      'end 2018/09/27
   End If
   'end 2018/07/12
   
   OnSaveData = False
   
   '更新輸入欄位
   strExSql = "UPDATE TransCaseTitle SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TCTCount - 1
      strTmp = Empty
      If m_TCTList(nIndex).fiOldData <> m_TCTList(nIndex).fiNewData Then
         If m_TCTList(nIndex).fiType = 0 Then
            If m_TCTList(nIndex).fiNewData = Empty Then
               strTmp = m_TCTList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_TCTList(nIndex).fiName & " = '" & ChgSQL(m_TCTList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TCTList(nIndex).fiNewData = Empty Then
               strTmp = m_TCTList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_TCTList(nIndex).fiName & " = " & m_TCTList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strExSql = strExSql & strTmp
            bFirst = False
         Else
            strExSql = strExSql & "," & strTmp
         End If
      End If
   Next nIndex
      
   '抓時間
   strExc(1) = Mid(Format(ServerTime, "000000"), 1, 4)
   
   If bDifference = True Then
      strExSql = strExSql & ",TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Val(strExc(1))
   End If
   
   '確認回報
   If iType = 1 Then
      strExSql = strExSql & IIf(bDifference = True, ", ", " ") & "TCT11=" & strSrvDate(1) & ", TCT12=" & Val(strExc(1))
      '主任=命名人員
      If m_TCT07 = m_TCT10 Then
         strExSql = strExSql & ", TCT08=" & strSrvDate(1) & ", TCT09=" & Val(strExc(1))
      End If
   End If

   'Added by Lydia 2018/09/20 判斷沒有變更,清空字串
   If InStr(strExSql, "TCT") = 0 Then
       strExSql = ""
   Else
   'end 2018/09/20
      strExSql = strExSql & " WHERE TCT01 = '" & m_TCT01 & "' "
   End If
      
   cnnConnection.BeginTrans
   If bDifference = True Or iType = 1 Then 'Move by Lydia 2018/09/27 從BeginTrans上面移過來
        If strExSql <> "" Then cnnConnection.Execute strExSql, intI '更新命名記錄 'Modified by Lydia 2018/09/20 +判斷非空白
        'Added by Lydia 2018/04/20 提申後從命名系統修改專利名稱則有欄位註記
        If Val(strCase(9)) > 0 And (m_PA05 <> txtData(3) Or m_PA06 <> txtData(4)) Then
             strExc(0) = "update transcasetitle set tct15='Y' where tct01='" & m_TCT01 & "' and tct15 is null "
             cnnConnection.Execute strExc(0), intI
        End If
        'end 2018/04/20
   End If
'Move by Lydia 2018/09/27 從"提申後從命名系統修改專利名稱則有欄位註記"下面移過來
        'Added by Lydia 2018/07/12
        '回寫翻譯費用檔
        If m_TF01t <> "" And (txtData(6) <> m_TF20 Or txtData(7) <> m_TF19 Or m_strSaveFiles <> "") Then
             'Modified by Lydia 2023/05/30 直接存入txtData(6)
             strExc(0) = "update transfee set tf20=" & CNULL(txtData(6)) & ",tf19=" & CNULL(txtData(7), True) & _
                              IIf(m_strSaveFiles <> "", ", tf29=null ", "") & " where tf01=" & CNULL(m_TF01t)
             cnnConnection.Execute strExc(0), intI
             'Added by Lydia 2018/09/17 命名作業拿掉相似度,記錄被取消的相似度在進度備註,並且發email通知程序(ex.FCP-59523)
             If m_TF29 = "Y" And txtData(6) = "" And Val(txtData(7)) = 0 Then
                   'Modified by Lydia 2023/05/30 FCP" & m_TF20 & "000=> " & m_TF20 & "
                   strExc(0) = "Update CaseProgress Set cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & "命名作業-取消相似案號" & m_TF20 & "，相似度" & m_TF19 & "％;'||cp64 where cp09=" & CNULL(m_TF01t)
                   cnnConnection.Execute strExc(0), intI
                   strExc(1) = strCase(1) & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", strCase(3) & strCase(4), "") & " 取消相似度"
                   strExc(2) = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4))
                   'Modified by Lydia 2023/05/30 FCP" & m_TF20 & "000=> " & m_TF20 & "
                   strExc(3) = "命名作業-取消相似案號" & m_TF20 & "，相似度" & m_TF19 & "%" & vbCrLf & _
                                    "請管制人員至新案建檔之翻譯頁籤，取消待比對。"
                   strExc(0) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                       " values( '" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd')" & _
                       ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & strExc(3) & "',null)"
                   cnnConnection.Execute strExc(0)
             End If
             'end 2018/09/17
        'Added by Lydia 2018/08/22 若輸入相似度後,程序未執行新案建檔(ex.FCP-59234)
        ElseIf m_TF01t = "" And m_TF01 <> "" And InStr("201,209,235", m_TF01pty) > 0 And (txtData(6) <> "" Or txtData(7) <> "") Then
             'Modified by Lydia 2023/05/30 直接存入txtData(6)
             strExc(0) = "insert into transfee (TF01,TF20,TF19) values (" & CNULL(m_TF01) & ", " & CNULL(txtData(6)) & ", " & CNULL(txtData(7), True) & ") "
             cnnConnection.Execute strExc(0), intI
        'end 2018/08/22
        End If
'end 2018/09/27
   
   'Added by Lydia 2018/10/18 黑白圖提申記錄在新案進度的備註
   If bUpdCP64 = True Then
        strExc(0) = "Update caseprogress set cp64=" & CNULL(ChangeTStringToTDateString(strSrvDate(1)) & " 命名-黑白圖提申(" & strUserNum & ");") & "||cp64  where cp09=" & CNULL(m_TCT01)
        cnnConnection.Execute strExc(0), intI
   End If
   'end 2018/10/18
   
   'Added by Lydia 2020/02/21 存檔「名稱有特殊字」
   strExc(1) = ""
   If strCase(12) <> IIf(ChkPA174.Value = 0, "", "Y") Then
       strExc(1) = strExc(1) & ", pa174=" & CNULL(IIf(ChkPA174.Value = 0, "", "Y"))
   End If
   If strExc(1) <> "" Then
       strExc(0) = "Update patent set " & Mid(strExc(1), 2) & " where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' "
       'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
       Pub_SeekTbLog strExc(0), , , , Me.Caption & "(" & Me.Name & ")"
       cnnConnection.Execute strExc(0), intI
   End If
   'end 2020/02/21
   
   'Added by Lydia 2021/04/22 工程師完成命名時能自動發email通知上級主管
   If iType = 1 Then
       strExc(1) = ""
       If m_TCT04 = m_TCT10 Then
            '命名人員=各組主管，不發email
       Else
            If m_TCT10 = m_TCT07 Then  '副理／主任=命名人員
                strExc(1) = m_TCT04  '各組主管
            Else
                'Modified by Lydia 2021/05/12 直接分給工程師=> email給各組主管
                If m_TCT07 = "" Then
                    strExc(1) = m_TCT04
                Else
                'end 2021/05/12
                    strExc(1) = m_TCT07 '副理／主任
                End If
            End If
       End If
       If strExc(1) <> "" Then
           strExc(2) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & "命名人員已完成命名，請主管進行確認。"
           strExc(3) = Mid(strExc(2), 1, Len(strExc(2)) - 1) & vbCrLf
           strExc(3) = strExc(3) & "命名人員：" & m_TCT10 & " " & lblData(8).Caption
           strExc(0) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
               ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "',null )"
           cnnConnection.Execute strExc(0)
       End If
   End If
   'end 2021/04/22
   
   cnnConnection.CommitTrans 'Move by Lydia 2018/09/27 移下來
   
   OnSaveData = True

Err03:
   If Err.Number <> 0 Then
      MsgBox Err.Description
      cnnConnection.RollbackTrans
   End If
End Function

Private Sub Opt1_Click(Index As Integer)
    If Opt1(Index).Value = True Then
       Frame1.BackColor = &H8000000F
    End If
End Sub

Private Sub Opt3_1_Click(Index As Integer)
    If Opt3_1(Index).Value = True And Chk2(4).Value = vbUnchecked Then
       Chk2(4).Value = vbChecked
    End If
End Sub

Private Sub Opt3_2_Click(Index As Integer)
    If Opt3_2(Index).Value = True And Chk2(6).Value = vbUnchecked Then
       Chk2(6).Value = vbChecked
    End If
End Sub

Private Sub Opt3_3_Click(Index As Integer)
    If Opt3_3(Index).Value = True And Chk2(8).Value = vbUnchecked Then
       Chk2(8).Value = vbChecked
    End If
End Sub

Private Sub Opt3_4_Click(Index As Integer)
    If Opt3_4(Index).Value = True And Chk2(10).Value = vbUnchecked Then
       Chk2(10).Value = vbChecked
    End If
End Sub

Private Sub Opt3_5_Click(Index As Integer)
    If Opt3_5(Index).Value = True And Chk2(12).Value = vbUnchecked Then
       Chk2(12).Value = vbChecked
    End If
End Sub

Private Sub Opt3_6_Click(Index As Integer)
    If Opt3_6(Index).Value = True And Chk2(14).Value = vbUnchecked Then
       Chk2(14).Value = vbChecked
    End If
End Sub

Private Sub Opt4s_Click(Index As Integer)
    If Opt4s(Index).Value = True Then
       Chk2(47).Value = vbChecked
       Frame5(9).BackColor = &H8000000F
       'Added by Lydia 2018/04/19
       'Modified by Lydia 2018/05/23 若新申請案進度已發文，不可勾選"提申前"。
       'If Val(strCase(9)) > 0 And Index = 1 And m_TCT20 <> "2" Then
       '    MsgBox "已有申請日，不可改成提申前 ! ", vbCritical
       If Val(n_CP27) > 0 And Index = 1 And m_TCT20 <> "2" Then
           MsgBox "新申請案進度已發文，不可改成提申前 ! ", vbCritical
       'end 2018/05/23
           Opt4s(Index).Value = 0
           If m_TCT20 <> "" Then
               Opt4s(Val(m_TCT20) - 1).Value = 1
           End If
       End If
    End If
End Sub

Private Sub Opt5_Click(Index As Integer)
    If Opt5(Index).Value = True Then
       If Index < 7 And txtData(23).Text <> "" Then txtData(23).Text = ""
       Frame3.BackColor = &H8000000F
    End If
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
    TextInverse txtData(Index)
End Sub

'Modified by Lydia 2021/09/27 改成Form 2.0
'Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      'Modified by Lydia 2023/05/30 +6相似案號
      Case 2, 6
          KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txtData_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   txtData(Index).ToolTipText = txtData(Index).Text
End Sub

'Added by Lydia 2021/09/27 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtData_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtData(Index)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
       Case 2   '欲翻譯此案件者
          If txtData(Index).Text <> "" Then
              'Modified by Lydia 2025/03/13 新增國外翻譯社
              'For intI = 41 To 44
              '     Chk2(intI).Value = vbUnchecked
              'Next
              For Each oChk In Chk27
                  oChk.Value = vbUnchecked
              Next
              'end 2025/03/13
              txtData(47).Text = ""
              'Modified by Lydia 2019/08/13 自2019年8月15日起實施，亦即自當日起交稿案件一律以調整後費率計算; 並且取消折扣案件之限制
              'If m_TCT27kind <> "" And InStr(m_TCT27kind, txtData(Index).Text) = 0 Then
              If m_TCT27kind <> "" And InStr(m_TCT27kind, txtData(Index).Text) = 0 And strSrvDate(1) < "20190815" Then
                  'Modified by Lydia 2018/09/26 +相似度
                  MsgBox IIf(InStr(m_TCT27kind, "A") = 0, "此案為固定報價、有折扣或有相似度，" & vbCrLf, "") & "欲翻譯此案件者，請輸入" & m_TCT27kind & " !", vbExclamation
                  GoTo ExceptRun
              End If
              If txtData(Index).Text <> "A" And txtData(Index).Text <> "B" Then
                  MsgBox "欲翻譯此案件者，請輸入A或B !", vbExclamation
                  GoTo ExceptRun
              End If
          End If
       Case 6   '相似案號
          If txtData(Index).Text <> "" Then
             'Modified by Lydia 2023/05/30 改成要輸入全部案號; ex.P-131591
             'If Len(Trim(txtData(Index))) <> 6 Then
             '   txtData(Index).Text = Right(String(6, "0") & Trim(txtData(Index)), 6)
             'End If
             'strExc(0) = GetPrjName("FCP-" & txtData(Index).Text & "-0-00")
             Call ChgCaseNo(txtData(Index), strExc)
             If (strExc(1) = "P" Or strExc(1) = "FCP") And Len(strExc(2)) = 6 Then
                strExc(0) = GetPrjName(strExc(1) & "-" & strExc(2) & "-" & strExc(3) & "-" & strExc(4))
             Else
                strExc(0) = ""
             End If
             'end 2023/05/30
             If strExc(0) = "" Then
                MsgBox "相似案號不存在專利基本檔 !", vbExclamation
                GoTo ExceptRun
             'Added by Lydia 2023/05/30
             ElseIf strExc(1) & strExc(2) = strCase(1) & strCase(2) Then
                MsgBox "相似案號不可輸入本案 !", vbExclamation
                GoTo ExceptRun
             'end 2023/05/30
             End If
             txtData(Index).Text = strExc(1) & strExc(2) & strExc(3) & strExc(4) 'Added by Lydia 2023/05/30
             Chk2(1).Value = 1
             'Added by Lydia 2018/07/12 上傳相似比對結果檔案,新案翻譯才需要
             'Modified by Lydia 2018/08/09 工程師約定8/13上線
             If strSrvDate(1) >= "20180813" And m_TF01 <> "" And m_TF01pty = "201" Then
                  CmdFile.Visible = True
             End If
             'end 2018/07/12
          End If
       Case 7   '相似內容%
          If txtData(Index).Text <> "" Then
             If txtData(6).Text = "" Then
                MsgBox "請輸入相似案號 !", vbExclamation
                txtData(6).SetFocus
                Txtdata_GotFocus 6
                Cancel = True
             ElseIf Val(txtData(Index).Text) < 0 Then
                MsgBox "相似內容%請輸入數字 !", vbExclamation
                txtData(7).SetFocus
                Txtdata_GotFocus 7
                Cancel = True
             End If
          End If
       Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19 '(說明書)技術領域~符號說明的位置和建議內容
          If txtData(Index).Text <> "" Then
              If Chk2(Index - 4).Value = vbUnchecked Then Chk2(Index - 4).Value = vbChecked
          'Added by Lydia 2018/08/07 如果說明書頁籤之內容不完整的位置敘述刪除,將會自動清空同一項的核取項和選擇項(缺or須修正) ex.FCP-59317發明內容 標題 缺or須修正 那項無法取消
          ElseIf InStr("08,10,12,14,16,18", Format(Index, "00")) > 0 Then
              Chk2(Index - 4).Value = vbUnchecked
              Select Case Index
                    Case 8
                          Opt3_1(0).Value = 0: Opt3_1(1).Value = False
                    Case 10
                          Opt3_2(0).Value = 0: Opt3_2(1).Value = False
                    Case 12
                          Opt3_3(0).Value = 0: Opt3_3(1).Value = False
                    Case 14
                          Opt3_4(0).Value = 0: Opt3_4(1).Value = False
                    Case 16
                          Opt3_5(0).Value = 0: Opt3_5(1).Value = False
                    Case 18
                          Opt3_6(0).Value = 0: Opt3_6(1).Value = False
              End Select
          'end 2018/08/07
          End If
       Case 20  '缺摘要建議內容
          If txtData(Index).Text <> "" Then
             If Chk2(16).Value = vbUnchecked Then Chk2(16).Value = vbChecked
             If Chk2(17).Value = vbUnchecked Then Chk2(17).Value = vbChecked
          End If
       Case 21, 22, 24 '缺頁,(說明書)其他問題,(申請專利範圍)項號錯誤
          If txtData(Index).Text <> "" Then
             If Chk2(Index - 3).Value = vbUnchecked Then Chk2(Index - 3).Value = vbChecked
          End If
       Case 25, 26 '依附關係錯誤
          If txtData(Index).Text <> "" Then
             If Chk2(22).Value = vbUnchecked Then Chk2(22).Value = vbChecked
          End If
       Case 27, 28 '依附關係不明確
          If txtData(Index).Text <> "" Then
             If Chk2(23).Value = vbUnchecked Then Chk2(23).Value = vbChecked
          End If
       Case 29, 30 '多附多 'Memo by Lydia 2018/04/17 原"不當依附"
          If txtData(Index).Text <> "" Then
             If Chk2(24).Value = vbUnchecked Then Chk2(24).Value = vbChecked
          End If
       Case 31 '引用記載形式
          If txtData(Index).Text <> "" Then
             If Chk2(25).Value = vbUnchecked Then Chk2(25).Value = vbChecked
          End If
       Case 32, 33, 34, 35 '標的不一致
          If txtData(Index).Text <> "" Then
             If Chk2(26).Value = vbUnchecked Then Chk2(26).Value = vbChecked
          End If
       Case 36, 37, 38 '不予專利
          If txtData(Index).Text <> "" Then
             If Chk2(27).Value = vbUnchecked Then Chk2(27).Value = vbChecked
          End If
       Case 39, 40 '混雜式請求項,(申請專利範圍)其他問題
          If txtData(Index).Text <> "" Then
             If Chk2(Index - 11).Value = vbUnchecked Then Chk2(Index - 11).Value = vbChecked
          End If
       Case 41, 42, 44, 45 '建議指定代表圖,缺圖,格式不符說明,不完整圖內容
          If txtData(Index).Text <> "" Then
             If Chk2(Index - 10).Value = vbUnchecked Then Chk2(Index - 10).Value = vbChecked
          End If
       Case 43  '格式不符圖內容
          If txtData(Index).Text <> "" Then
             If Chk2(34).Value = vbUnchecked Then Chk2(34).Value = vbChecked
          End If
       Case 46  '(圖式)其他問題
          If txtData(Index).Text <> "" Then
             If Chk2(40).Value = vbUnchecked Then Chk2(40).Value = vbChecked
          End If
       'Added by Lydia 2018/04/18
       Case 47 '其他指定翻譯
          If Trim(txtData(Index).Text) <> "" Then
            'Modified by Lydia 2025/03/13 新增國外翻譯社
            'Chk2(41).Value = vbUnchecked
            'Chk2(42).Value = vbUnchecked
            'Chk2(43).Value = vbUnchecked
            'Chk2(44).Value = vbChecked
            For Each oChk In Chk27
               If oChk.Index = 0 Then
                  oChk.Value = vbChecked
               Else
                  oChk.Value = vbUnchecked
               End If
            Next
            'end 2025/03/13
            strExc(1) = Right(Trim(UCase(txtData(Index).Text)), 1)
            If InStr("A,B", strExc(1)) = 0 Then
                MsgBox "其他指定翻譯請在名稱後面加A或B !", vbExclamation
                GoTo ExceptRun
            'Added by Lydia 2018/06/01 排除命名人員本人
            ElseIf InStr(Trim(UCase(txtData(Index).Text)), lblData(8).Caption) > 0 Then
                MsgBox "命名人員欲翻譯此案件，請填上一欄位 !", vbExclamation
                txtData(Index).Text = ""
                Chk2(44).Value = vbUnchecked
                Cancel = True
            'end 2018/06/01
            Else
                'Added by Lydia 2018/09/21 判斷在職員工名稱
                If Trim(txtData(47).Text) <> Trim(m_TCTList(12).fiOldData) Then
                    'Modified by Lydia 2019/05/31 含外翻人員
                    'strExc(2) = GetPrjSalesNM_2(Mid(txtData(47).Text, 1, Len(txtData(47).Text) - 1))
                    strExc(2) = GetPrjSalesNM_2(Mid(txtData(47).Text, 1, Len(txtData(47).Text) - 1), , , , True)
                    If strExc(2) = "" Then
                        MsgBox "請輸入在職員工名稱 !", vbExclamation
                       GoTo ExceptRun
                    End If
                End If
                'end 2018/09/21
                'Modified by Lydia 2019/08/13 自2019年8月15日起實施，亦即自當日起交稿案件一律以調整後費率計算; 並且取消折扣案件之限制
                'If m_TCT27kind <> "" And InStr(m_TCT27kind, strExc(1)) = 0 Then
                If m_TCT27kind <> "" And InStr(m_TCT27kind, strExc(1)) = 0 And strSrvDate(1) < "20190815" Then
                    'Modified by Lydia 2018/09/26 +相似度
                    MsgBox IIf(InStr(m_TCT27kind, "A") = 0, "此案為固定報價、有折扣或有相似度", "") & "欲翻譯此案件者，請輸入" & m_TCT27kind & " !", vbExclamation
                    GoTo ExceptRun
                End If
            End If
          End If
   End Select
    
   'Remove by Lydai 2021/05/06 直接用TextBox.MaxLength控制長度
   'If CheckLengthIsOK(Me.txtData(Index).Text, Me.txtData(Index).MaxLength) = False Then
   '   Cancel = True
   'End If
   'end 2021/05/06
   
   Exit Sub
   
ExceptRun:
   txtData(Index).SetFocus
   Txtdata_GotFocus Index
   Cancel = True
End Sub
'Modified by Lydia 2018/12/21 +bolJump 是否跳過檢查
Private Function TxtValidate1(Optional ByVal bolJump As Boolean = False) As Boolean
Dim tInx As Integer  '頁籤
Dim inB As Integer
Dim tmpB As Boolean
Dim tmpArr As Variant 'Added by Lydia 2018/07/12
Dim dbTfRate As Double, bolIsHigher As Boolean  'Added by Lydia 2021/07/29 判斷翻譯費折扣率＞30%
Dim strTmp As String 'Added by Lydia 2025/03/13

   TxtValidate1 = False
   
'案件名稱-檢查
   tInx = 0
   If Trim(txtData(3)) = "" Then
      MsgBox "請輸入中文案件名稱 !", vbExclamation
      SSTab1.Tab = tInx
      txtData(3).SetFocus
      Txtdata_GotFocus 3
      Exit Function
   End If
   
   'Modified by Lydia 2018/12/21 + bolJump
   If Frame1.Visible = True And bolJump = False Then
      inB = 0
      For Each oOpt In Opt1
         If oOpt.Value = True Then inB = oOpt.Index + 1
      Next
      If inB = 0 Then
         MsgBox "請輸入設計案屬性 !", vbExclamation
         SSTab1.Tab = tInx
         Frame1.BackColor = &HC0FFC0
         Exit Function
      End If
   End If
   
   If Chk2(1).Value = vbChecked And (Trim(txtData(6)) = "" Or Trim(txtData(7)) = "") Then
      MsgBox "請輸入相似案號和相似內容% !", vbExclamation
      SSTab1.Tab = tInx
      inB = 6
      If Trim(txtData(6)) <> "" Then inB = 7
      txtData(inB).SetFocus
      Txtdata_GotFocus inB
      Exit Function
   End If
   'Added by Lydia 2024/05/31 相似案號檢查; 發現FCP-071046的相似案號只有「FCP071020」
   If Trim(txtData(6)) <> "" Then
      Call Txtdata_Validate(6, tmpB)
      If tmpB = True Then
         SSTab1.Tab = tInx
         inB = 6
         txtData(inB).SetFocus
         Txtdata_GotFocus inB
         Exit Function
      End If
   End If
   'end 2024/05/31
   
   'Added by Lydia 2018/07/12 檢查相似比對結果檔案,新案翻譯(未發文,未分案,待比對)才需要
   'Modified by Lydia 2018/08/09 工程師約定8/13上線
    If strSrvDate(1) >= "20180813" And Chk2(1).Value = vbChecked And m_TF01 <> "" And m_TF01pty = "201" And Val(m_TF01cp27) = 0 _
              And (m_TF29 = "Y" Or m_TF01cp14 = "" Or (m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) > 0)) Then
        strExc(1) = Dir(strResPath & "\" & strCase(1) & strCase(2) & "*.res.doc*")
        'Added by Lydia 2018/09/19 開放PDF (ex.FCP-59599)
        If strExc(1) = "" Then
             strExc(1) = Dir(strResPath & "\" & strCase(1) & strCase(2) & "*.res.pdf")
        End If
        'end 2018/09/19
        'Modified by Lydia 2018/09/19 改判斷
        'If (strExc(1) = "" Or m_strSaveFiles <> "") And m_TCT10 = strUserNum Then
        '    If m_strSaveFiles = "" Then
        If (strExc(1) = "" Or m_strSaveFiles <> "") Then
            'Modified by Lydia 2018/11/26 以前一畫面的員工編號判斷
            'If m_strSaveFiles = "" And m_TCT10 = strUserNum Then
            If m_strSaveFiles = "" And m_TCT10 = m_UserNo Then
        'end 2018/09/19
                'Modified by Lydia 2018/09/19 開放PDF
                'If MsgBox("是否上傳相似比對結果檔案(*.RES.doc/docx)？", vbCritical + vbYesNo + vbDefaultButton1) = vbYes Then
                If MsgBox("是否上傳相似比對結果檔案(*.RES.doc/docx 或 *.RES.PDF)？", vbCritical + vbYesNo + vbDefaultButton1) = vbYes Then
                    CmdFile.Visible = True
                    CmdFile.SetFocus
                    Exit Function
                End If
            'Modified by Lydia 2018/09/19 改判斷
            'Else
            ElseIf m_strSaveFiles <> "" Then
                tmpArr = Empty
                tmpArr = Split(m_strSaveFiles, "&")
                strExc(2) = ""
                For intI = 0 To UBound(tmpArr)
                    strExc(3) = Trim(tmpArr(intI))
                    If strExc(3) <> "" Then
                        If InStr(strExc(3), " (") > 0 Then strExc(3) = RTrim(Mid(strExc(3), 1, InStr(strExc(3), " (")))
                        If Dir(strExc(3)) <> "" Then '檔案路徑正確
                             If InStr(strExc(3), "\") > 0 Then strExc(3) = Mid(strExc(3), InStrRev(strExc(3), "\") + 1)
                             'Modified by Lydia 2018/09/19 開放PDF (ex.FCP-59599)
                             'Modifed by Lydia 2018/09/27  開放可上傳多個檔案
                             'If UCase(strExc(3)) <> strCase(1) & strCase(2) & ".RES.DOC" And UCase(strExc(3)) <> strCase(1) & strCase(2) & ".RES.DOCX" And UCase(strExc(3)) <> strCase(1) & strCase(2) & ".RES.PDF" _
                                 And UCase(strExc(3)) <> strCase(1) & Val(strCase(2)) & ".RES.DOC" And UCase(strExc(3)) <> strCase(1) & Val(strCase(2)) & ".RES.DOCX" And UCase(strExc(3)) <> strCase(1) & Val(strCase(2)) & ".RES.PDF" Then
                              If (Left(UCase(strExc(3)), Len(strCase(1) & strCase(2))) <> strCase(1) & strCase(2) And Left(UCase(strExc(3)), Len(strCase(1) & Val(strCase(2)))) <> strCase(1) & Val(strCase(2))) _
                                 Or (InStr(".RES.DOC;.RES.PDF", Right(UCase(strExc(3)), 8)) = 0 And Right(UCase(strExc(3)), 9) <> ".RES.DOCX") Then
                                 strExc(2) = strExc(2) & tmpArr(intI) & vbCrLf
                             End If
                        Else
                             strExc(2) = strExc(2) & tmpArr(intI) & vbCrLf
                        End If
                    End If
                Next intI
                If strExc(2) <> "" Then
                    'Modified by Lydia 2018/09/19 開放PDF
                    'MsgBox "下列檔案不符合命名規則(ex.FCP012345.RES.Doc)：" & strExc(2), vbCritical
                    MsgBox "下列檔案不符合命名規則" & vbCrLf & "(ex.FCP012345.RES.Doc 或 FCP012345.RES.PDF)：" & strExc(2), vbCritical
                    Exit Function
                End If
            End If
        End If
    End If
    'end 2018/07/12

   
   If bolJump = False Then 'Added by Lydia 2018/12/21 + bolJump
        'Added by Lydia 2023/01/18 命名作業不可新增告代和主動修正
        If strNotBList <> "" And (Chk2(45).Value = vbChecked Or Chk2(47).Value = vbChecked) Then
           If InStr(strNotBList, ",") > 0 Then
              strExc(9) = Mid(strNotBList, 1, InStr(strNotBList, ",") - 1)
           Else
              strExc(9) = strNotBList
           End If
           strExc(10) = Pub_GetITS01Type(strExc(9))
           'Added by Lydia 2025/10/02 針對Spruson & Ferguson (Asia)Pte Ltd (Y21071)之「Y01 命名作業」進行細部設定 (僅阻擋新案「235核對中說格式」或「209檢視中說」案件的「告代」與「主動修正」收文，並跳出說明提醒，但不阻擋其他如「201新案翻譯」、「210  製作中說」新案的「告代」與「主動修正」收文）)
           strExc(2) = ""
           If InStr("Y21071000,", m_PA75) > 0 And InStr("201,210", m_TF01pty) > 0 Then
              strExc(2) = "B"
           End If
           'end 2025/10/02
           If strExc(2) = "" Then 'Added by Lydia 2025/10/02
              If PUB_GetITStoList(Me.Name, strExc(10), strExc(9), False, False, , , "Y01") = True Then
              End If
              SSTab1.Tab = tInx
              Chk2(46).Value = vbChecked
              Exit Function
           End If 'Added by Lydia 2025/10/02
        End If
        'end 2023/01/18
      
        If Chk2(45).Value = vbUnchecked And Chk2(46).Value = vbUnchecked And Chk2(47).Value = vbUnchecked Then
           MsgBox "請輸入是否收文主動修正和告代 !", vbExclamation
           SSTab1.Tab = tInx
           Frame5(8).BackColor = &HC0FFC0
           Exit Function
        'Modified by Lydia 2018/03/01 +chk2(48)
        ElseIf Chk2(46).Value = vbChecked And (Chk2(45).Value = vbChecked Or Chk2(47).Value = vbChecked Or Chk2(48).Value = vbChecked) Then
           MsgBox "勾選不需收文，不會收文主動修正和告代 !", vbExclamation
           SSTab1.Tab = tInx
           Frame5(8).BackColor = &HC0FFC0
           Exit Function
        'Modified by Lydia 2018/04/18 檢查何時告代
        'ElseIf Chk2(47).Value = vbChecked Then
        End If
        If Chk2(47).Value = vbChecked Then
        'end 2018/04/18
             inB = 0
             For Each oOpt In Opt4s
                If oOpt.Value = True Then inB = oOpt.Index + 1
             Next
             If inB = 0 Then
                 MsgBox "請輸入何時告代 !", vbExclamation
                 SSTab1.Tab = tInx
                 Frame5(9).BackColor = &HC0FFFF
                 Exit Function
             End If
        'Added by Lydia 2018/04/19 檢查收文狀態
        ElseIf m_TCT20 <> "" Then
             'Modified by Lydia 2022/04/28 改成共用模組
             'If frm090902_2.ChkCPisExist(strCase, "901", strCP901, strExc(2)) = True Then
             If PUB_ChkBCPisExist(strCase, "901", strCP901, strExc(2)) = True Then
                 If Val(strExc(2)) > 0 Then
                     MsgBox "請通知承辦人員到卷宗區搬移檔案後，才可取消 !", vbCritical
                     Chk2(47).Value = vbChecked
                     SSTab1.Tab = tInx
                     Frame5(9).BackColor = &HC0FFFF
                     Exit Function
                 End If
             End If
        End If
        'Added by Lydia 2018/04/18 檢查何時主動修正
        If Chk2(45).Value = vbChecked Then
             inB = 0
             For Each oOpt In Opt4s2
                If oOpt.Value = True Then inB = oOpt.Index + 1
             Next
             If inB = 0 Then
                 MsgBox "請輸入何時主動修正 !", vbExclamation
                 SSTab1.Tab = tInx
                 Frame5(10).BackColor = &HC0FFFF
                 Exit Function
             End If
        'Added by Lydia 2018/04/19 檢查收文狀態
        ElseIf m_TCT117 <> "" Then
             'Modified by Lydia 2022/04/28 改成共用模組
             'If frm090902_2.ChkCPisExist(strCase, "203", strCP203, strExc(2)) = True Then
             If PUB_ChkBCPisExist(strCase, "203", strCP203, strExc(2)) = True Then
                 If Val(strExc(2)) > 0 Then
                     MsgBox "請通知承辦人員到卷宗區搬移檔案後，才可取消 !", vbCritical
                     Chk2(45).Value = vbChecked
                     SSTab1.Tab = tInx
                     Frame5(10).BackColor = &HC0FFFF
                     Exit Function
                 End If
             End If
        End If
        'end 2018/04/18
        
        inB = 0
        For Each oOpt In Opt5
           If oOpt.Value = True Then inB = oOpt.Index + 1
        Next
        If inB = 0 Then
           MsgBox "請輸入案件類別 !", vbExclamation
           SSTab1.Tab = tInx
           Frame3.BackColor = &HC0FFC0
           Exit Function
        ElseIf inB = 8 And Trim(txtData(23)) = "" Then
           MsgBox "請輸入其他類別內容 !", vbExclamation
           SSTab1.Tab = tInx
           txtData(23).SetFocus
           Txtdata_GotFocus 23
           Exit Function
        End If
   End If 'end 2018/12/21
   
   '欲翻譯此案件者/指定翻譯
   tmpB = False
   Txtdata_Validate 2, tmpB
   If tmpB = True Then
      SSTab1.Tab = tInx
      txtData(2).SetFocus
      Txtdata_GotFocus 2
      Exit Function
   End If
   'Modified by Lydia 2025/03/13 新增國外翻譯社
   'inB = 0
   'For intI = 41 To 44
   '     If Chk2(intI).Value = vbChecked Then inB = intI - 40
   'Next
   'If inB = 4 And Trim(txtData(47).Text) = "" Then
   inB = -1
   For Each oChk In Chk27
      If oChk.Value = vbChecked Then
         inB = oChk.Index
      End If
   Next
   If inB = 0 And Trim(txtData(47).Text) = "" Then
   'end 2025/03/13
      MsgBox "請輸入其他翻譯 !", vbExclamation
      SSTab1.Tab = tInx
      txtData(47).SetFocus
      Txtdata_GotFocus 47
      Exit Function
   End If
   If inB > 0 And Trim(txtData(2).Text) <> "" Then
      MsgBox "欲翻譯此案件者和指定翻譯不可同時輸入 !", vbExclamation
      'Modified by Lydia 2025/03/13 新增國外翻譯社
      'For intI = 41 To 44
      '     Chk2(intI).Value = vbUnchecked
      'Next
      For Each oChk In Chk27
         oChk.Value = vbUnchecked
      Next
      'end 2025/03/13
      SSTab1.Tab = tInx
      txtData(2).SetFocus
      Txtdata_GotFocus 2
      Exit Function
   End If
   'Added by Lydia 2018/04/18 其他指定翻譯：所內員工
    Txtdata_Validate 47, tmpB
    If tmpB = True Then
       SSTab1.Tab = tInx
       txtData(47).SetFocus
       Txtdata_GotFocus 47
       Exit Function
    End If
   'end 2018/04/18
       
    'Added by Lydia 2018/07/12 檢查欲翻譯人員
    If m_TF01pty = "201" Then
        strExc(1) = m_TCTList(11).fiOldData
        'Modified by Lydia 2025/03/13 新增國外翻譯社
        'strExc(2) = txtData(2).Text
        'If Chk2(41).Value = vbChecked Then
        '     strExc(2) = "1"
        'ElseIf Chk2(42).Value = vbChecked Then
        '     strExc(2) = "2"
        'ElseIf Chk2(43).Value = vbChecked Then
        '     strExc(2) = "3"
        'End If
        strTmp = txtData(2).Text
        For Each oChk In Chk27
           If oChk.Value = vbChecked Then
              If oChk.Index = 0 Then
                 strTmp = "Z"  '其他:指定為Z
              Else
                 strTmp = oChk.Index '依序
              End If
           End If
        Next
        'end 2025/03/13
        'Added by Lydia 2021/07/29 增加判斷其他
        If m_TCTList(12).fiOldData <> "" Then strExc(1) = m_TCTList(12).fiOldData
        If txtData(47).Text <> "" Then strExc(2) = txtData(47).Text
        'end 2021/07/29
        'Modified by Lydia 2025/03/13
        'If strExc(2) <> strExc(1) Then
        If strTmp <> strExc(1) Then
             If Val(m_TF01cp27) > 0 Then
                 MsgBox "新案翻譯已發文，不可變更欲翻譯人員！", vbCritical
                 Exit Function
             ElseIf m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) = 0 Then
                 MsgBox "新案翻譯已分案，不可變更欲翻譯人員！", vbCritical
                 Exit Function
             End If
             'Added Lydia 2025/04/23 外專翻譯分案承辦人不得為翻譯社及外譯人員 --公告1120419-05
             'Modified by Lydia 2025/07/01 增加例外案件設定InStr(m_str所內譯例外,  strCase(1) & strCase(2) & strCase(3) & strCase(4)) = 0 And
             If InStr(m_str所內譯例外, strCase(1) & strCase(2) & strCase(3) & strCase(4)) = 0 And (InStr(m_str所內譯, m_PA26) > 0 Or InStr(m_str所內譯, m_PA75) > 0) Then
                 If Val(strTmp) > 0 Then
                     'Modified by Lydia 2025/06/05 「BASF集團公司為申請人的所有專利案件」改為「本案所有」
                     MsgBox "本案所有相關翻譯事宜（201新案翻譯/927其他翻譯）皆須由本所工程師翻譯/處理，不得委外。", vbExclamation + vbOKOnly
                     Exit Function
                 End If
             Else
             'end 2025/04/23
                'Added by Lydia 2021/07/29 判斷翻譯費折扣率
                dbTfRate = PUB_GetTransFeeRate(strCase(1), strCase(2), strCase(3), strCase(4), , bolIsHigher, True)
             End If 'Added by Lydia 2025/04/23
             '控制翻譯費折扣率＞30%客戶案件之承辦人只能為所內人員上班譯編號。
             If dbTfRate > 30 Then
                 'Modified by Lydia 2025/03/13 新增國外翻譯社
                 'If Chk2(41).Value = 1 Or Chk2(42).Value = 1 Or Chk2(43).Value = 1 Or UCase(txtData(2)) = "A" Or UCase(Right(txtData(47), 1)) = "A" Then
                 If Val(strTmp) > 0 Or UCase(txtData(2)) = "A" Or UCase(Right(txtData(47), 1)) = "A" Then
                     MsgBox "該案件之承辦人只能為所內人員上班譯編號！", vbExclamation
                     Exit Function
                 End If
             ElseIf bolIsHigher = True Then  '折扣率＞30%但是例外控制的客戶
                  '不受限
             End If
             'end 2021/07/29

        End If
    End If
    'end 2018/07/12
             
'說明書-檢查
   tInx = 1
   tmpB = False
   For inB = 8 To 22
      Txtdata_Validate inB, tmpB
      If tmpB = True Then
         SSTab1.Tab = tInx
         txtData(inB).SetFocus
         Txtdata_GotFocus inB
         Exit Function
      End If
   Next
   
'申請專利範圍-檢查
   tInx = 2
   For inB = 23 To 46
      Txtdata_Validate inB, tmpB
      If tmpB = True Then
         SSTab1.Tab = tInx
         txtData(inB).SetFocus
         Txtdata_GotFocus inB
         Exit Function
      End If
   Next
   
   'Added by Lydia 2018/10/18 提醒工程師檢查是否彩圖提申
    bUpdCP64 = False
    If Chk2(49).Value = vbUnchecked And m_PA63 = "Y" And (m_TCT01cp64 = "" _
            Or (m_TCT01cp64 <> "" And InStr(m_TCT01cp64, "命名-黑白圖提申") = 0)) Then
       'Modified by Lydia 2024/09/26 debug: 工程師組別strCase(5)>>strCase(7)
       If strSrvDate(1) < "20230501" Or (strSrvDate(1) >= "20230501" And strCase(7) = "3") Then  'Added by Lydia 2023/04/26 (電子電機、化學、機械 三組)直接以彩圖製作成ori版本提申(自112年5月1日起實施); 日文組仍維持原程式=>需判斷是否以彩圖提申
             inB = MsgBox("本案客戶有提供彩圖，請判斷是否以彩圖提申" & vbCrLf & "是：彩圖提申" & vbCrLf & "否：黑白圖提申" & vbCrLf & "取消：回到命名作業再重新判斷", vbInformation + vbYesNoCancel + vbDefaultButton3)
             If inB = 2 Then 'Cancel
                 Exit Function
             ElseIf inB = 7 Then 'No
                  bUpdCP64 = True
             Else
                  Chk2(49).Value = vbChecked
             End If
       End If 'Added by Lydia 2023/04/26
    End If
   'end 2018/10/18
   
   'Added by Lydia 2020/02/21 檢查「名稱有特殊字」
   If strCase(1) = "P" Or strCase(1) = "FCP" Then
       If Pub_GetPA174toFile("2", strCase(1), strCase(2), strCase(3), strCase(4), Me, frm100101_M_1) = True Then
           strExc(1) = "Y"
       Else
           strExc(1) = "N"
       End If
       If ChkPA174.Value = vbUnchecked And strExc(1) = "Y" Then
           If MsgBox("原始檔區已有案件名稱Word檔，請問是否取消「名稱有特殊字」？", vbInformation + vbYesNo + vbDefaultButton2, "檢查資料") = vbNo Then
               Exit Function
           End If
       End If
       If ChkPA174.Value = vbChecked And strExc(1) = "N" Then
           If MsgBox("原始檔區沒有案件名稱Word檔，請問是否繼續作業？", vbInformation + vbYesNo + vbDefaultButton2, "檢查資料") = vbNo Then
               Exit Function
           End If
       End If
       '當「名稱有特殊字」有勾選，並且有修改案件名稱，將原始檔之維護word檔自動打開，並彈訊息提醒。
       If ChkPA174.Value = vbChecked And bolAskPA174 = False Then  '不用再次彈訊息
           If m_PA05 & m_PA06 <> txtData(3) & txtData(4) Then
               MsgBox "名稱有特殊字，案件名稱有修改，請一併修改案件名稱Word檔。", vbInformation, "檢查資料"
               Call ProcPA174toFile("Y")
               Exit Function
           End If
       End If
   End If
   'end 2020/02/21
   
   'Added by Lydia 2023/03/10 FMP大陸新藥發明專利權期限補償控管
   If Chk2(51).Visible = True Then
      If Chk2(51).Value = 1 And Opt5(0).Value = False Then
          MsgBox "專利權期間延長相關的案件類別請選擇" & Opt5(0).Caption, vbExclamation, "大陸新藥發明專利權期限補償控管"
          SSTab1.Tab = 0
          Exit Function
      End If
   End If
   'end 2023/03/10
   
   'Added by Lydia 2024/10/17 英文組的彩圖依照PA63設定--from Phoebe ; ex.FCP-72571  'Added by Lydia 2023/04/26 (電子電機、化學、機械 三組)直接以彩圖製作成ori版本提申(自112年5月1日起實施); 日文組仍維持原程式=>需判斷是否以彩圖提申
   If strSrvDate(1) >= "20230501" And strCase(7) <> "3" And m_PA63 = "Y" And Chk2(49).Value = vbUnchecked Then
       MsgBox "承辦人員於接洽單註記有彩圖提申！", vbCritical + vbOKOnly
       SSTab1.Tab = 0
       Exit Function
   End If
    
   'Added by Lydia 2021/09/27 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/09/27
   
   '清除畫面所有欄位的跳行符號
   PUB_FilterFormText Me
   
   TxtValidate1 = True
End Function

'Added by Lydai 2017/12/27 外文本
Private Sub cmdOpen_Click()
Dim hLocalFile As Long 'Added by Lydia 2018/06/21

On Error GoTo ErrHand01 'Added by Lydia 2018/03/23 無權限的錯誤要改訊息
    'Added by Lydia 2020/01/20 開啟[原始檔區]
    If InStr(cmdOpen.Caption, "原始檔") > 0 Then
        If PUB_CheckFormExist("frm100101_M") Then
            MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
        Else
            If cmdOpen.Tag = "" Then
                MsgBox strCase(1) & "-" & strCase(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
            Else
                strExc(1) = ""
                frm100101_M.m_strKey = cmdOpen.Tag '多筆總收文號
                frm100101_M.SetParent Me
                If frm100101_M.QueryData = True Then
                   frm100101_M.Show
                   Me.Hide
                End If
            End If
        End If
    Else
    'end 2020/01/20
        'Modified by Lydia 2018/05/09 +系統別
        'Modifiede by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
        'strExc(1) = Pub_GetFCPcaseFilePath(strCase(2), , strCase(1))
        'If Dir(strExc(1) & "\*.*") <> "" Then
        '     'Modified by Lydia 2018/06/21 用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
        '     'SHELL "Explorer.exe " & strExc(1), vbNormalFocus  '開啟案件資料夾
        '     ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
        'Else
        '     MsgBox lblData(6).Caption & "在" & strExc(1) & "的資料夾不存在或無檔案!", vbInformation
        'End If
        strExc(1) = ""
        'end 2021/12/06
    End If 'Added by Lydia 2020/01/20
    
'Added by Lydia 2018/03/23
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         '全部錯誤訊息統一
         MsgBox "無法讀取" & strExc(1) & "，請通知電腦中心！", vbCritical
         Resume Next
    End If
'end 2018/03/23
End Sub

'Added by Lydia 2018/04/18
Private Sub Opt4s2_Click(Index As Integer)
    If Opt4s2(Index).Value = True Then
       Chk2(45).Value = vbChecked
       Frame5(10).BackColor = &H8000000F
       'Modified by Lydia 2018/05/23 若新申請案進度已發文，不可勾選"提申前"。
       'If Val(strCase(9)) > 0 And Index = 1 And m_TCT117 <> "2" Then
       '    MsgBox "已有申請日，不可改成提申前 ! ", vbCritical
       If Val(n_CP27) > 0 And Index = 1 And m_TCT117 <> "2" Then
           MsgBox "新申請案進度已發文，不可改成提申前 ! ", vbCritical
       'end 2018/05/23
           Opt4s2(Index).Value = 0
           If m_TCT117 <> "" Then
               Opt4s2(Val(m_TCT117) - 1).Value = 1
           End If
           Exit Sub
       End If
       
       'Added by Lydia 2018/05/15 檢查是否有承辦收的A類主動修正
       If Index = 1 And m_CP203 <> "" Then
           MsgBox "已依客戶其它指示收文提申前主動修正 !", vbCritical
           Chk2(45).Value = vbUnchecked
           Exit Sub
       End If
       'end 2018/05/15
    End If
End Sub

'Added by Lydia 2018/07/12 上傳相似比對結果檔案
Private Sub cmdFile_Click()
   
   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.lblCaseNo = strCase(1) & "-" & strCase(2) & "-" & strCase(3) & "-" & strCase(4)
   frm090801_8.Label4.Visible = False
   frm090801_8.bolNotPDF = True
   frm090801_8.Show vbModal
End Sub

'Added by Lydia 2020/02/21
Private Sub CmdPA174_Click()
    Call ProcPA174toFile("N")
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟/維護FCP0xxxxx.新案性質.案件名稱.doc
Private Sub ProcPA174toFile(ByVal pKind As String)
Dim strKind As String

    If ChkPA174.Value = vbUnchecked Then
        MsgBox "請先勾選「有特殊字」!", vbInformation + vbOKOnly, Me.Caption
    Else
        If pKind = "Y" Then 'bolAskPA174
            strKind = "3"
        Else
            strKind = "1"
        End If
        If Pub_GetPA174toFile(strKind, strCase(1), strCase(2), strCase(3), strCase(4), Me, frm100101_M_1) = True Then
        End If
    End If
    
End Sub

'Added by Lydia 2020/02/21
Public Sub PubShowNextData()
   '原始檔Word檔維護，上傳後直接進入存檔
   If bolAskPA174 = True Then
        Call cmdok_Click(cmdState) '確定->存檔
   End If
End Sub

