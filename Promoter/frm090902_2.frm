VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090902_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "待確認"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔(&S)"
      Height          =   360
      Index           =   3
      Left            =   6360
      TabIndex        =   198
      Top             =   15
      Visible         =   0   'False
      Width           =   860
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "外文本(&P)"
      Height          =   300
      Left            =   120
      TabIndex        =   192
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "退回(&B)"
      Height          =   360
      Index           =   0
      Left            =   5400
      TabIndex        =   177
      Top             =   15
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認(&A)"
      Height          =   360
      Index           =   1
      Left            =   6357
      TabIndex        =   178
      Top             =   15
      Width           =   1160
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Left            =   7615
      TabIndex        =   184
      Top             =   15
      Width           =   1160
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6612
      Left            =   120
      TabIndex        =   41
      Top             =   420
      Width           =   8652
      _ExtentX        =   15261
      _ExtentY        =   11663
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   4057
      TabCaption(0)   =   "案件名稱"
      TabPicture(0)   =   "frm090902_2.frx":0000
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
      Tab(0).Control(25)=   "Label1(7)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(46)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblData(12)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblData(13)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblData(15)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblData(16)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label5(2)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(54)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblCMboth"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label1(6)"
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
      Tab(0).Control(47)=   "Frame6"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Chk2(51)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "說明書"
      TabPicture(1)   =   "frm090902_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5(12)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "申請專利範圍 ＆ 圖示"
      TabPicture(2)   =   "frm090902_2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5(7)"
      Tab(2).Control(1)=   "Frame5(6)"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox Chk2 
         Caption         =   "專利權期間延長相關"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   241
         Top             =   2070
         Width           =   2175
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame6"
         Height          =   220
         Left            =   90
         TabIndex        =   237
         Top             =   2670
         Width           =   1005
         Begin VB.CheckBox ChkPA174 
            Caption         =   "有特殊字"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   0
            TabIndex        =   238
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   270
         Style           =   1  '圖片外觀
         TabIndex        =   236
         Top             =   2910
         Width           =   840
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame5"
         Height          =   5850
         Index           =   12
         Left            =   -74850
         TabIndex        =   202
         Top             =   450
         Width           =   8355
         Begin VB.CheckBox Chk2 
            Caption         =   "其它問題"
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   91
            Top             =   5490
            Width           =   1035
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "缺頁"
            Height          =   255
            Index           =   18
            Left            =   360
            TabIndex        =   89
            Top             =   5205
            Width           =   855
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "建議內容"
            Height          =   255
            Index           =   17
            Left            =   840
            TabIndex        =   87
            Top             =   4890
            Width           =   1095
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "缺摘要"
            Height          =   255
            Index           =   16
            Left            =   360
            TabIndex        =   86
            Top             =   4650
            Width           =   1215
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame5"
            Height          =   660
            Index           =   5
            Left            =   660
            TabIndex        =   229
            Top             =   3960
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   14
               Left            =   960
               TabIndex        =   80
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_6 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   81
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_6 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   82
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   15
               Left            =   960
               TabIndex        =   84
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   18
               Left            =   4200
               TabIndex        =   83
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
               TabIndex        =   85
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
               TabIndex        =   233
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   32
               Left            =   3720
               TabIndex        =   232
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   33
               Left            =   1800
               TabIndex        =   231
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   34
               Left            =   3480
               TabIndex        =   230
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
            Left            =   660
            TabIndex        =   224
            Top             =   3276
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   12
               Left            =   960
               TabIndex        =   74
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_5 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   75
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_5 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   76
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   13
               Left            =   960
               TabIndex        =   78
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   16
               Left            =   4200
               TabIndex        =   77
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
               TabIndex        =   79
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
               TabIndex        =   228
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   27
               Left            =   3720
               TabIndex        =   227
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   28
               Left            =   1800
               TabIndex        =   226
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   30
               Left            =   3480
               TabIndex        =   225
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
            Left            =   660
            TabIndex        =   218
            Top             =   2592
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   10
               Left            =   960
               TabIndex        =   68
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_4 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   69
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_4 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   70
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   11
               Left            =   960
               TabIndex        =   72
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   14
               Left            =   4200
               TabIndex        =   71
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
               TabIndex        =   73
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
               TabIndex        =   223
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   21
               Left            =   3720
               TabIndex        =   222
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   22
               Left            =   1800
               TabIndex        =   221
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   23
               Left            =   3480
               TabIndex        =   220
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "說明"
               Height          =   255
               Index           =   35
               Left            =   360
               TabIndex        =   219
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
            Left            =   660
            TabIndex        =   213
            Top             =   1908
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   9
               Left            =   960
               TabIndex        =   66
               Top             =   320
               Width           =   1095
            End
            Begin VB.OptionButton Opt3_3 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   64
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton Opt3_3 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   63
               Top             =   0
               Width           =   615
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   8
               Left            =   960
               TabIndex        =   62
               Top             =   0
               Width           =   735
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   13
               Left            =   2040
               TabIndex        =   67
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
               TabIndex        =   65
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
               TabIndex        =   217
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   17
               Left            =   1800
               TabIndex        =   216
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   18
               Left            =   3720
               TabIndex        =   215
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "發明內容"
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   214
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
            Left            =   660
            TabIndex        =   208
            Top             =   1224
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   7
               Left            =   960
               TabIndex        =   60
               Top             =   320
               Width           =   1095
            End
            Begin VB.OptionButton Opt3_2 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   58
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton Opt3_2 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   57
               Top             =   0
               Width           =   615
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   180
               Index           =   6
               Left            =   960
               TabIndex        =   56
               Top             =   0
               Width           =   735
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   11
               Left            =   2040
               TabIndex        =   61
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
               TabIndex        =   59
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
               TabIndex        =   212
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   13
               Left            =   1800
               TabIndex        =   211
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   14
               Left            =   3720
               TabIndex        =   210
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "先前技術"
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   209
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
            Left            =   660
            TabIndex        =   203
            Top             =   540
            Width           =   7400
            Begin VB.CheckBox Chk2 
               Caption         =   "標題"
               Height          =   255
               Index           =   4
               Left            =   960
               TabIndex        =   50
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton Opt3_1 
               Caption         =   "缺"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   51
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Opt3_1 
               Caption         =   "須修正"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   52
               Top             =   0
               Width           =   855
            End
            Begin VB.CheckBox Chk2 
               Caption         =   "建議內容"
               Height          =   255
               Index           =   5
               Left            =   960
               TabIndex        =   54
               Top             =   320
               Width           =   1095
            End
            Begin MSForms.TextBox txtData 
               Height          =   285
               Index           =   8
               Left            =   4200
               TabIndex        =   53
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
               TabIndex        =   55
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
               TabIndex        =   207
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "位置"
               Height          =   255
               Index           =   10
               Left            =   3720
               TabIndex        =   206
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "("
               Height          =   255
               Index           =   11
               Left            =   1800
               TabIndex        =   205
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   ")，"
               Height          =   255
               Index           =   29
               Left            =   3480
               TabIndex        =   204
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "內容不完整"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   49
            Top             =   260
            Width           =   1215
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "說明書"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   1215
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   22
            Left            =   1080
            TabIndex        =   92
            Top             =   5490
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
            Left            =   1920
            TabIndex        =   90
            Top             =   5190
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
            Left            =   1920
            TabIndex        =   88
            Top             =   4890
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
            Left            =   2460
            TabIndex        =   235
            Top             =   260
            Width           =   5715
         End
         Begin VB.Label Label1 
            Caption         =   "頁數："
            Height          =   255
            Index           =   36
            Left            =   1320
            TabIndex        =   234
            Top             =   5205
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Height          =   2580
         Index           =   8
         Left            =   6240
         TabIndex        =   194
         Top             =   3345
         Width           =   2175
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Height          =   840
            Index           =   9
            Left            =   480
            TabIndex        =   196
            Top             =   1440
            Width           =   1575
            Begin VB.OptionButton Opt4s 
               Caption         =   "提申後告代"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   90
               Width           =   1335
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
               Caption         =   "當日告代"
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   1215
            End
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
            Caption         =   "不需收文"
            Height          =   255
            Index           =   46
            Left            =   240
            TabIndex        =   21
            Top             =   2280
            Width           =   1620
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
               Caption         =   "提申前"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Opt4s2 
               Caption         =   "提申後"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   14
               Top             =   0
               Width           =   975
            End
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   4656
         TabIndex        =   39
         Text            =   "Combo1"
         Top             =   6228
         Width           =   2655
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "列印(&P)"
         Height          =   330
         Index           =   2
         Left            =   7536
         TabIndex        =   40
         Top             =   6204
         Width           =   860
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Height          =   2145
         Index           =   7
         Left            =   -74880
         TabIndex        =   182
         Top             =   4140
         Width           =   8415
         Begin VB.CheckBox Chk2 
            Caption         =   "其它問題"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   142
            Top             =   1760
            Width           =   1095
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "用途說明"
            Height          =   255
            Index           =   39
            Left            =   4920
            TabIndex        =   149
            Top             =   1500
            Width           =   1695
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "色彩"
            Height          =   255
            Index           =   38
            Left            =   4920
            TabIndex        =   148
            Top             =   1215
            Width           =   975
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不主張設計的部分"
            Height          =   255
            Index           =   37
            Left            =   4920
            TabIndex        =   147
            Top             =   930
            Width           =   1935
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "超過一個實施例"
            Height          =   255
            Index           =   36
            Left            =   4920
            TabIndex        =   146
            Top             =   645
            Width           =   1695
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不完整：圖"
            Height          =   255
            Index           =   35
            Left            =   4920
            TabIndex        =   144
            Top             =   360
            Width           =   1275
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "格式不符（圖"
            Height          =   255
            Index           =   34
            Left            =   480
            TabIndex        =   139
            Top             =   960
            Width           =   1425
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "彩圖"
            Height          =   255
            Index           =   33
            Left            =   1320
            TabIndex        =   137
            Top             =   660
            Width           =   720
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "缺圖（　　　　）圖"
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   136
            Top             =   660
            Width           =   1995
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "建議指定代表圖：圖"
            Height          =   255
            Index           =   31
            Left            =   480
            TabIndex        =   134
            Top             =   360
            Width           =   2025
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "圖式"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   133
            Top             =   120
            Width           =   975
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   46
            Left            =   1200
            TabIndex        =   143
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
            Index           =   45
            Left            =   6240
            TabIndex        =   145
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
            Index           =   44
            Left            =   2280
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   138
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
            TabIndex        =   135
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
            TabIndex        =   183
            Top             =   1240
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Height          =   3795
         Index           =   6
         Left            =   -74880
         TabIndex        =   166
         Top             =   360
         Width           =   8415
         Begin VB.CheckBox Chk2 
            Caption         =   "其它問題"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   131
            Top             =   3405
            Width           =   1095
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "混雜式請求項：請求項"
            Height          =   255
            Index           =   28
            Left            =   600
            TabIndex        =   129
            Top             =   3135
            Width           =   2175
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "不予專利（請求項"
            Height          =   255
            Index           =   27
            Left            =   600
            TabIndex        =   125
            Top             =   2820
            Width           =   1815
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "標的不一致"
            Height          =   255
            Index           =   26
            Left            =   600
            TabIndex        =   120
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "屬於引用記載形式之獨立項：請求項"
            Height          =   255
            Index           =   25
            Left            =   600
            TabIndex        =   118
            Top             =   1605
            Width           =   3315
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "多附多（附屬項"
            Height          =   255
            Index           =   24
            Left            =   600
            TabIndex        =   115
            Top             =   1320
            Width           =   1725
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "依附關係不明確（附屬項"
            Height          =   255
            Index           =   23
            Left            =   600
            TabIndex        =   112
            Top             =   1020
            Width           =   2340
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "依附關係錯誤（附屬項"
            Height          =   255
            Index           =   22
            Left            =   600
            TabIndex        =   109
            Top             =   720
            Width           =   2205
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "項號錯誤：請求項"
            Height          =   255
            Index           =   21
            Left            =   600
            TabIndex        =   107
            Top             =   420
            Width           =   1815
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "申請專利範圍"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   106
            Top             =   120
            Width           =   1575
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   40
            Left            =   1200
            TabIndex        =   132
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
            TabIndex        =   130
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   124
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
            Index           =   33
            Left            =   4080
            TabIndex        =   122
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
            Index           =   31
            Left            =   3960
            TabIndex        =   119
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
            TabIndex        =   117
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
            Index           =   28
            Left            =   6240
            TabIndex        =   114
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
            Index           =   26
            Left            =   6120
            TabIndex        =   111
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
            Index           =   24
            Left            =   2400
            TabIndex        =   108
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
            TabIndex        =   181
            Top             =   2813
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，法條"
            Height          =   255
            Index           =   39
            Left            =   5640
            TabIndex        =   180
            Top             =   2820
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "，標的："
            Height          =   255
            Index           =   38
            Left            =   3720
            TabIndex        =   179
            Top             =   2820
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "，標的："
            Height          =   255
            Index           =   52
            Left            =   3960
            TabIndex        =   176
            Top             =   2520
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "被依附之請求項"
            Height          =   255
            Index           =   51
            Left            =   960
            TabIndex        =   175
            Top             =   2520
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "，標的："
            Height          =   255
            Index           =   50
            Left            =   3360
            TabIndex        =   174
            Top             =   2220
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "附屬項"
            Height          =   255
            Index           =   49
            Left            =   960
            TabIndex        =   173
            Top             =   2220
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   47
            Left            =   7365
            TabIndex        =   172
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，應依附於請求項"
            Height          =   255
            Index           =   45
            Left            =   4095
            TabIndex        =   171
            Top             =   1320
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   43
            Left            =   7980
            TabIndex        =   170
            Top             =   1020
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，應依附於請求項"
            Height          =   255
            Index           =   42
            Left            =   4695
            TabIndex        =   169
            Top             =   1020
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "）"
            Height          =   255
            Index           =   41
            Left            =   7860
            TabIndex        =   168
            Top             =   735
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "，應依附於請求項"
            Height          =   255
            Index           =   40
            Left            =   4575
            TabIndex        =   167
            Top             =   735
            Width           =   1560
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   1068
         Left            =   120
         TabIndex        =   163
         Top             =   5145
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
         Begin VB.CheckBox Chk27 
            Caption         =   "捷恩凱"
            Height          =   255
            Index           =   2
            Left            =   4848
            TabIndex        =   38
            Top             =   72
            Visible         =   0   'False
            Width           =   975
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
            Caption         =   "其他"
            Height          =   255
            Index           =   0
            Left            =   3912
            TabIndex        =   35
            Top             =   480
            Width           =   705
         End
         Begin MSForms.TextBox txtData 
            Height          =   285
            Index           =   47
            Left            =   4680
            TabIndex        =   36
            Top             =   480
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
            TabIndex        =   199
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "是否欲翻譯此案件者：　　　   (A:下班翻譯 B.上班翻譯)"
            Height          =   255
            Index           =   25
            Left            =   880
            TabIndex        =   165
            Top             =   165
            Width           =   4575
         End
         Begin VB.Label Label1 
            Caption         =   "指定翻譯："
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   164
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Height          =   720
         Left            =   120
         TabIndex        =   161
         Top             =   4425
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
            TabIndex        =   162
            Top             =   160
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   800
         Left            =   120
         TabIndex        =   160
         Top             =   3630
         Width           =   5895
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   201
            Top             =   130
            Width           =   4605
            Begin VB.CheckBox Chk2 
               Caption         =   "有序列表"
               Height          =   255
               Index           =   50
               Left            =   3150
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
            Top             =   225
            Width           =   900
         End
         Begin VB.CheckBox Chk2 
            Caption         =   "本案說明書內容與"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   450
            Width           =   1770
         End
         Begin VB.Label Label4 
            Caption         =   "%相同"
            Height          =   252
            Index           =   6
            Left            =   4380
            TabIndex        =   240
            Top             =   456
            Width           =   540
         End
         Begin VB.Label Label4 
            Caption         =   "之內容"
            Height          =   228
            Index           =   4
            Left            =   3180
            TabIndex        =   239
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
         TabIndex        =   159
         Top             =   3345
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
            TabIndex        =   187
            Top             =   0
            Width           =   975
         End
      End
      Begin MSForms.TextBox txtData 
         Height          =   285
         Index           =   5
         Left            =   1590
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   3000
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
         Top             =   2670
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
         Top             =   2355
         Width           =   7005
         VariousPropertyBits=   671105051
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "(英)："
         Height          =   255
         Index           =   6
         Left            =   1140
         TabIndex        =   157
         Top             =   2670
         Width           =   495
      End
      Begin VB.Label lblCMboth 
         AutoSize        =   -1  'True
         Caption         =   "lblCMboth"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   200
         Top             =   2130
         Width           =   2685
      End
      Begin VB.Label Label1 
         Caption         =   "P.S.其他指定翻譯請在名稱後面加A或B"
         ForeColor       =   &H00FF00FF&
         Height          =   252
         Index           =   54
         Left            =   216
         TabIndex        =   197
         Top             =   6252
         Width           =   3252
      End
      Begin VB.Label Label5 
         Caption         =   "命名人員："
         Height          =   180
         Index           =   2
         Left            =   6120
         TabIndex        =   193
         Top             =   825
         Width           =   900
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   16
         Left            =   4010
         TabIndex        =   191
         Top             =   825
         Width           =   855
         ForeColor       =   255
         BackColor       =   -2147483645
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
         TabIndex        =   190
         Top             =   1125
         Width           =   1215
         ForeColor       =   255
         BackColor       =   -2147483645
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
         TabIndex        =   189
         Top             =   520
         Width           =   555
         BackColor       =   -2147483645
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
         TabIndex        =   188
         Top             =   520
         Width           =   855
         BackColor       =   -2147483645
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
         Left            =   3576
         TabIndex        =   186
         Top             =   6276
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
         Left            =   1140
         TabIndex        =   158
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱　 (中)："
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   156
         Top             =   2370
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "案件性質："
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   154
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "譯畢期限："
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   153
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "總收文號："
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   152
         Top             =   1125
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   4
         Left            =   4010
         TabIndex        =   151
         Top             =   1125
         Width           =   1800
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "3175;459"
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
         TabIndex        =   150
         Top             =   1485
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   5
         Left            =   4010
         TabIndex        =   105
         Top             =   1485
         Width           =   855
         BackColor       =   -2147483645
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
         TabIndex        =   104
         Top             =   1800
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   6
         Left            =   4010
         TabIndex        =   103
         Top             =   1800
         Width           =   1455
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "2566;459"
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
         TabIndex        =   102
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   3
         Left            =   6120
         TabIndex        =   101
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "急件，請於　   　　　　　　　前譯畢名稱"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   100
         Top             =   540
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "國　　籍："
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   99
         Top             =   1485
         Width           =   1095
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   10
         Left            =   7320
         TabIndex        =   98
         Top             =   1485
         Width           =   1215
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "2143;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   8
         Left            =   7080
         TabIndex        =   97
         Top             =   840
         Width           =   795
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "1402;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   9
         Left            =   7080
         TabIndex        =   96
         Top             =   1125
         Width           =   795
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "1402;459"
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
         TabIndex        =   95
         Top             =   1485
         Width           =   1095
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   2
         Left            =   1400
         TabIndex        =   94
         Top             =   1485
         Width           =   1200
         BackColor       =   -2147483645
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
         TabIndex        =   93
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   3
         Left            =   1395
         TabIndex        =   47
         Top             =   1800
         Width           =   1560
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "2752;459"
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
         TabIndex        =   46
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "本所期限："
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   45
         Top             =   825
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   0
         Left            =   1200
         TabIndex        =   44
         Top             =   540
         Width           =   855
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "1508;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   1
         Left            =   1200
         TabIndex        =   43
         Top             =   825
         Width           =   855
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "1508;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   260
         Index           =   11
         Left            =   7095
         TabIndex        =   42
         Top             =   1800
         Width           =   1440
         BackColor       =   -2147483645
         VariousPropertyBits=   27
         Size            =   "2540;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "frm090902_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Lydia 2021/09/27 改成Form2.0 ; lblData(index)、txtData(index)
'Created by Lydia 2017/11/14 外專新案未命名區-待確認明細
Option Explicit
Private Const m_FS As Integer = 16  '輸入欄位對應Table欄位的起始位置
'Modified by Lydia 2023/02/16 改成共用常數m_FE=>TF_TCT, m_NotFS=>TF_TCTnotFS
'Private Const m_FE As Integer = 119 '輸入欄位對應Table欄位的終止位置
'Private Const m_NotFS As String = "112,113,114,115" 'Added by Lydia 2018/03/01 排除不修改的欄位
'end 2023/02/16
Private Const m_Frame5 As Integer = 12 'Added by Lydia 2019/09/19 Frame5的index

Dim strPrinter As String
Dim m_PrevForm As Form '前一畫面
Dim m_iStiu As String '狀態：M-主管確認, Q-卷宗區進入
Dim oLbl As Control
Dim oTxt As Control
Dim oChk As CheckBox
Dim oOpt As OptionButton
Dim intLeaveKind As Integer

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
'Modified by Lydia 2020/02/17 +12-名稱有特殊字(PA174)
Dim strCase(1 To 12) As String '1~4本所案號pa01~pa04,5-專利種類pa08,6-申請國家pa09,7-分案組別pa150,8-設計案屬性pa158
Dim m_TCT01 As String  '收文號=PK
Dim m_TCT01cp10 As String 'Added by Lydia 2024/02/05 新案收文號之案件性質
Dim m_TCT01cp06 As String, m_TCT01cp142 As String, m_TCT01cp164 As String 'Added by Lydia 2025/08/29 新案收文號之本所期限CP06、指定送件日CP142、指定日期方式CP164
Dim m_TCT04 As String  '工程師主管
Dim m_TCT04chk As String 'Added by Lydia 2023/05/31
Dim s_TCT04m As String 'Added by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
Dim m_TCT05 As String  '工程師主管-確認日期
Dim m_TCT07 As String  '工程師主任
Dim m_TCT10 As String  '命名人員編號
Dim m_TCT14 As String 'Added by Lydia 2019/09/19 重送確認記錄
Dim m_TCT20 As String 'Added by Lydai 2018/04/19 何時告代
Dim m_TCT117 As String 'Added by Lydai 2018/04/19 何時主動修正
Dim m_PA05 As String, m_PA06 As String  'Added by Lydia 2018/04/20 專利檔中文、英文名稱
Dim bolEmail As Boolean '是否發Email通知
Dim strTempA As String  '暫存列印字串
Dim bolChgText As Boolean '是否處理完暫存列印字串
'Added by Lydia 2017/12/15
Dim m_TCT08 As String  '工程師主任-確認日期
Dim m_WList As String '請假的主管
Dim bolUpdMan As Boolean '是否代主管確認
Dim bolUpdMan2 As Boolean '是否代主任確認
'end 2017/12/15

'Added by Lydia 2017/11/24 列印模組最後放在主管確認,因為卷宗區用主管確認的表單來看
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 600, ciColGap = 150
Private Const ciPrtX = 700, ciPrtY = 700 '表格內起始X,Y位置
Private Const m_TTop = 600, m_Tbottom = 16000   '表格下邊界
Private Const m_MidCol = 9000 '收文主動修正和告代欄位
Dim iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim m_dblTitleHeight As Double '抬頭
Dim m_TBWidth As Double '表格寬
Dim m_TCT27kind As String '欲翻譯此案件者可輸入的選項
Private Const sChked = "◎" '核取項
Private Const sUnchked = "○" '非核取項
Dim strCP203 As String, strCP901 As String '主動修正和告代的收文號
Dim m_UserSt16 As String  'Added by Lydia 2017/12/29 傳入員工編號的工程師組別
Dim rInx As Integer '讀取次數
Dim n_CP118 As String, n_CP27 As String 'Added by Lydia 2018/04/18 新申請案：是否電子送件、發文日
Dim tCP06 As String, tCP27 As String 'Added by Lydia 2018/04/18 新案翻譯：所限、發文日
Dim m_UserList  As String 'Added by Lydia 2018/04/19 主管確認後,可修改人員
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
'Added by Lydia 2019/09/02
Dim bAgainTrans As Boolean '是否重送命名記錄
Dim bolAgain As Boolean '是否啟用重送命名記錄的功能
'Dim bolAdd203 As Boolean, bolAdd901 As Boolean '詢問是否產生主動修正或告代
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
Dim m_strTFAcon As String 'Added by Lydia 2025/10/15

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

Public Sub SetParent(ByRef fm As Form, ByVal pCase As String, ByVal pNo As String, ByVal pUser As String, Optional ByVal pType As String = "Q", Optional ByVal pList As String = "")
   Set m_PrevForm = fm
   m_TCT01 = pNo
   m_UserNo = pUser
   m_iStiu = pType
   m_WList = pList
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
   For Each oOpt In opt1
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
   cmdFile.Visible = False
   m_strSaveFiles = ""
   
   'Added by Lydia 2018/10/22
   lblCMboth.Caption = ""
   lblCMboth.Tag = ""
   
   CmdOpen.Tag = "" 'Added by Lydia 2020/01/20
   
   'Added by Lydia 2020/02/17
   ChkPA174.Value = vbUnchecked
   bolAskPA174 = False
   cmdState = -1
   Frame6.Enabled = True
End Sub

Private Sub TxtEnabled()
   'Remove by Lydia 2018/04/19 開放從卷宗區點選命名記錄(RCD.Menu)查看資料時，若使用者為命名人員或該人員的上級主管可修改資料後存檔
'   For Each oTxt In txtData
'      oTxt.Locked = True
'   Next
'   'Option和CheckBox選項無值,則Enabled=False
'   For Each oChk In Chk2
'      If oChk.Value = False Then
'         oChk.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt1
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt3_1
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt3_2
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt3_3
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt3_4
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt3_5
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt3_6
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt4s
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   For Each oOpt In Opt5
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   'Added by Lydia 2018/04/18
'   For Each oOpt In Opt4s2
'      If oOpt.Value = False Then
'         oOpt.Enabled = False
'      End If
'   Next
'   'end 2018/04/18
'end 2018/04/19
   
   cmdOK(3).Visible = False
   '主管只修改一案兩請和相似案,有其他問題退回命名人員
   If m_iStiu = "M" Then
      Chk2(0).Enabled = True
      Chk2(1).Enabled = True
      'Remove by Lydia 2018/04/19
      'txtData(6).Locked = False
      'txtData(7).Locked = False
      'Added by Lydia 2020/04/27 若命名人員是主管本人,直接進入主管確認階段,沒有退回功能; (取消四組主管不能分案給自己的限制)
      'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
      'If m_TCT04 = m_TCT10 And m_UserNo = m_TCT04 Then
      If (m_TCT04 = m_TCT10 And m_UserNo = m_TCT04) Or (s_TCT04m = m_TCT10 And m_UserNo = s_TCT04m) Then
          cmdOK(0).Visible = False
      End If
      'end 2020/04/27
   '來自卷宗區的使用者,只有看的權限
   ElseIf m_iStiu = "Q" Then
      cmdOK(0).Visible = False
      cmdOK(1).Visible = False
      'Added by Lydia 2018/04/19 開放從卷宗區點選命名記錄(RCD.Menu)查看資料時，若使用者為命名人員或該人員的上級主管可修改資料後存檔
      If InStr(m_UserList, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
         'Added by Lydia 2023/05/31
         If m_TCT04chk = "" Then
            MsgBox "本案處於新案二次認領階段，暫時無法修改！", vbInformation, "新案認領管制"
         Else
         'end 2023/05/31
            cmdOK(3).Visible = True
         End If 'Added by Lydia 2023/05/31
      End If
      'end 2018/04/19
   End If
   
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
      If Frame5(8).Enabled = True And strNotBList <> "" And (Index = 45 Or Index = 47) Then
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
   
   'Modified by Lydia 2018/04/19 卷宗區(RCD.menu)->存檔
   'If m_iStiu = "M" And cmdOK(0).Enabled = True And cmdOK(1).Enabled = True Then
   If (m_iStiu = "M" And cmdOK(0).Enabled = True And cmdOK(1).Enabled = True) Or cmdOK(3).Visible = True Then
      If CheckDataDiff(bolCheck) Then
         If bolCheck = True Then
            If MsgBox("資料有異動，確定離開嗎?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
      End If
   End If
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bDiff As Boolean
Dim tmpBol As Boolean 'Added by Lydia 2022/05/30

   cmdState = Index 'Added by Lydia 2020/02/17
   
   Select Case Index
      Case 0  '退回
         intLeaveKind = 0
         '先發信,後更新日期
         If TxtValidate1 = True Then
            If CheckDataDiff(bDiff) = True Then
            End If
            'Pub_Send_CFPdg = False 'Mark by Lydia 2022/05/30
            '郵件主旨:(急件！) FCP-00000新案命名(完成期限：本所期限，譯畢期限：急件才顯示)
            strExc(1) = "退回 - " & strCase(1) & "-" & Val(strCase(2)) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "")
            strExc(1) = strExc(1) & "新案命名"
            'Modified by Lydia 2018/12/17
            'If lblData(1).Caption <> "" Or lblData(12).Caption <> "" Then
            If lblData(12).Caption <> "" Then
               strExc(1) = strExc(1) & "("
               'Remove by Lydia 2018/12/17 將完成期限刪除以免工程師誤判命名deadline
               'If lblData(1).Caption <> "" Then
               '   strExc(1) = strExc(1) & "完成期限：" & lblData(1).Caption & IIf(lblData(12).Caption <> "", "，", "")
               'End If
               'end 2018/12/17
               If lblData(12).Caption <> "" Then
                  strExc(1) = strExc(1) & "譯畢期限：" & lblData(12).Caption & " " & lblData(13).Caption
               End If
               strExc(1) = strExc(1) & ")"
            End If
            '郵件-內文
            strExc(2) = "智權人員姓名：" + lblData(9).Caption + vbCrLf + _
                  "本所案號：" + lblData(6).Caption + vbCrLf + _
                  "專利種類：" + lblData(2).Caption + vbTab + vbTab + vbTab + "案件性質：" + lblData(15).Caption + vbCrLf + _
                  "中說類型：" + lblData(3).Caption + vbCrLf + _
                   "案件名稱(中)：" & txtData(3).Text & vbCrLf + _
                   "案件名稱(英)：" & txtData(4).Text & vbCrLf + _
                   "案件名稱(日)：" & txtData(5).Text & vbCrLf + _
                   vbCrLf + "退回原因："
            'Modified by Lydia 2022/05/30 改用frm880019
            'frm880005.txtEmail(0) = m_TCT10
            'frm880005.txtEmail(0).Tag = Me.Name
            'frm880005.txtEmail(1) = strExc(1)
            'frm880005.txtEmail(2) = strExc(2)
            'frm880005.Show vbModal
            '''改成共用變數控制
            'If Pub_Send_CFPdg = False Then
            '   Exit Sub
            'End If
            frm880019.txtReceiver = m_TCT10
            frm880019.txtSubject = strExc(1)
            frm880019.txtContent = strExc(2)
            frm880019.cmdAttach.Visible = False
            frm880019.SetParent Me
            frm880019.Show vbModal
            tmpBol = frm880019.m_bolDone '是否傳送成功
            Unload frm880019
            If tmpBol = False Then
                 MsgBox "送信失敗，請重新Email !", vbCritical
                 Exit Sub
            End If
            'end 2022/05/30
            
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
      Case 1  '確認
         If TxtValidate1 = True Then
            'Added by Lydia 2017/12/27
            If Trim(txtData(3)) = "待命名" And Trim(txtData(4)) = "" Then
                 MsgBox "案件名稱中文為待命名並且英文為空白時，不可確認!", vbExclamation
                 Exit Sub
            End If
            'end 2017/12/27

            'Added by Lydia 2021/05/06 命名作業若有指定翻譯人員，查詢該認領人員之新案翻譯未上完稿日案件
            If m_TF01pty = "201" And m_TF01cp14 = "" And Val(m_TF01cp27) = 0 And (txtData(2) <> "" Or txtData(47) <> "") Then
                 'Modified by Lydia 2025/10/15
                 'If txtData(2) = "A" Or txtData(2) = "B" Then
                 '    strExc(2) = m_TCT10
                 'Else
                 '    strExc(2) = GetPrjSalesNM_2(Mid(txtData(47).Text, 1, Len(txtData(47).Text) - 1), , , , True)
                 'End If
                 strExc(4) = GetTCT27val(strExc(2))
                 'end 2025/10/15
                 strExc(4) = Pub_GetEngEP09List(strExc(2))
                 If strExc(4) <> "" Then
                     If MsgBox(GetStaffName(strExc(2)) & " 尚未完稿案件：" & strExc(4) & vbCrLf & vbCrLf & "是否繼續確認？", vbInformation + vbYesNo + vbDefaultButton2, "翻譯人員檢查") = vbNo Then
                         Exit Sub
                     End If
                 End If
            End If
            'end 2021/05/06
            
            'Added by Lydia 2017/12/15 是否代確認
            If m_WList <> "" And InStr(m_WList, m_TCT04) > 0 And m_UserNo <> m_TCT04 And m_UserNo = strUserNum Then
               If m_TCT14 = "" Then 'Added by Lydia 2019/09/24 重送命名流程,預設不代主管確認
                     'Modified by Lydia 2017/12/27 不限主管請假 => 拿掉 "工程師主管請假"
                     If MsgBox("是否一併進行主管確認？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                          bolUpdMan = True '代主管確認
                     End If
               End If
            'Added by Lydia 2024/03/18 外專機械設計組人員異動調整程式：針對內專主管直接進行代主管確認
            'Modified by Lydia 2024/03/20 判斷主任與主管不同組=直接進行代主管確認; 機械組案件101,102由外專工程師(可能是電子組或化學組), 103由內專工程師翻譯
            'Modified by Lydia 2024/04/18 外專機械組工程師暫時由外專其他組主管帶領; ex.FCP-071584
            'ElseIf m_TCT07 = m_UserNo And Mid(m_TCT07, 4, 1) = "9" And m_UserNo = strUserNum Then
            ElseIf m_TCT07 = m_UserNo And m_UserNo = strUserNum And (Mid(m_TCT07, 4, 1) = "9" Or (PUB_GetStaffST16(m_TCT10) = "4") And PUB_GetStaffST16(m_TCT07) <> "4") Then
               bolUpdMan = True '代主管確認
            'end 2024/03/18
            Else
                  If m_UserNo = m_TCT04 And Val(m_TCT05) = 0 And m_TCT07 <> "" And Val(m_TCT08) = 0 Then
                        If MsgBox("分案主任未確認，是否一併進行確認？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                             bolUpdMan2 = True '代主任確認
                        End If
                  End If
            End If
            'end 2017/12/15
            
            'Added by Lydia 2025/10/15 檢查命名作業有輸入翻譯人員(包含內翻或外翻)，同時也有其他人員從"認翻譯作業"進行認翻譯，彈選項
            m_strTFAcon = ""
            If m_TF01pty = "201" And m_TF01cp14 = "" And Val(m_TF01cp27) = 0 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) Then
                strExc(4) = GetTCT27val(strExc(1))
                If strExc(4) <> "" Then
                   '來自認翻譯作業frm060122
                   strExc(0) = "select st01,st02,decode(tfa05,'A','下班','上班') as tfa05 from transfeeassign,staff where tfa01='" & m_TF01 & "' and tfa06 is null and tfa04=st01(+) "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      RsTemp.MoveFirst
                      strExc(0) = "": m_strTFAcon = ""
                      Do While Not RsTemp.EOF
                         strExc(0) = strExc(0) & "、" & RsTemp.Fields("st02") & "-" & RsTemp.Fields("tfa05")
                         m_strTFAcon = m_strTFAcon & ";" & RsTemp.Fields("st01")
                         RsTemp.MoveNext
                      Loop
                      strExc(0) = Mid(strExc(0), 2)
                      m_strTFAcon = Mid(m_strTFAcon, 2)
                   End If
                   '命名作業的指定翻譯人員
JumpToReInput:
                   If m_strTFAcon <> "" And strExc(4) <> "" Then
                      strSql = InputBox("命名作業指定翻譯人員：" & strExc(4) & vbCrLf & "認翻譯人員：" & strExc(0) & vbCrLf & vbCrLf & _
                               "請輸入選項1或2：" & vbCrLf & "　　　1=確認" & strExc(0) & "為翻譯人員，" & vbCrLf & "　　　　清空命名作業的翻譯人員，" & vbCrLf & _
                               "　　　2=確認" & strExc(4) & "為翻譯人員，" & vbCrLf & "　　　　刪除認領翻譯人員，" & vbCrLf & "　　　　同時發送email通知認領人員", "認翻譯人員", "")
                      If strSql = "" Then
                         Exit Sub
                      Else
                         If strSql = "1" Then '選擇:來自認翻譯作業frm060122
                            For Each oChk In Chk27
                               oChk.Value = vbUnchecked
                            Next
                            txtData(2) = ""
                            txtData(47) = ""
                            m_strTFAcon = ""
                         ElseIf strSql = "2" Then
                            '選擇:命名作業指定翻譯人員
                         Else
                            GoTo JumpToReInput
                         End If
                      End If
                   End If
                End If
            End If
            'end 2025/10/15
            
            'Move by Lydia 2025/10/15 從'Added by Lydia 2017/12/15 是否代確認的上方移下來
            If CheckDataDiff(bDiff) = True Then
            End If
            
            'Added by Lydia 2019/09/19 因為提申前修改收文(重送命名流程) , 無法用之前判斷修改前/後有差異來處理
                                                     '所以先參考重送記錄,再判斷是否產生新的收文
            'Mark by Lydia 2019/09/23 先保留
'            bolAdd203 = True: bolAdd901 = True
'            If bolAgain = True And m_TCT14 <> "" And InStr(m_TCT14, "變更") > 0 And (m_UserNo = m_TCT04 Or bolUpdMan = True) Then   '有重送記錄(主管確認階段)
'               If Chk2(45).Value = vbChecked Then
'                   If ChkCPisExist(strCase, "203", , , "2") = True Then
'                       If MsgBox("本案已有發文主動修正203的進度，是否要另外收文？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
'                          bolAdd203 = False
'                       End If
'                   End If
'               End If
'               If Chk2(47).Value = vbChecked Then
'                   If ChkCPisExist(strCase, "901", , , "2") = True Then
'                       If MsgBox("本案已有發文告代901的進度，是否要另外收文？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
'                          bolAdd901 = False
'                       End If
'                   End If
'               End If
'            End If
            
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
      Case 2  '列印
         '先存檔,後列印
         If m_iStiu = "M" Then
            If TxtValidate1 = True Then
               If CheckDataDiff(bDiff) = True Then
                  If bDiff = True Then
                     If OnSaveData(Index) = False Then Exit Sub
                     '存檔後,重讀DB
                     ClearForm
                     If ReadData = True Then
                     End If
                  End If
               End If
            PUB_PrintTCTcon m_TCT01, Me.Combo1.Text, strPrinter
            End If
         Else '非主管,無權限修改
            PUB_PrintTCTcon m_TCT01, Me.Combo1.Text, strPrinter
         End If
       'Added by Lydia 2018/04/19
       Case 3 '卷宗區->存檔
         If TxtValidate1 = True Then
            If Trim(txtData(3)) = "待命名" And Trim(txtData(4)) = "" Then
                 MsgBox "案件名稱中文為待命名並且英文為空白時，不可存檔!", vbExclamation
                 Exit Sub
            End If
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
                End If
                
                'Added by Lydia 2019/09/02 命名人員可以從卷宗區修改命名記錄，重新送回命名區請主管重新確認(重送命名流程)
                'Remove by Lydia 2019/09/17 修改內容不要用PDF，直接在email內文寫上修改前後。
'                If bolAgain = True And bAgainTrans = True Then '發email
'                    '附件
'                    PUB_PrintTCTcon m_TCT01, Me.Combo1.Text, strPrinter, "$$" & strCase(1) & strCase(2) & IIf(strCase(3) & strCase(4) = "000", "", strCase(3) & strCase(4)) & "修改後"
'                    Sleep 1000
'                    strExc(1) = ""
'                    strExc(2) = Dir(App.path & "\$$" & strCase(1) & strCase(2) & "*修改*.pdf")
'                    Do While strExc(2) <> ""
'                        strExc(1) = strExc(1) & "*" & App.path & "\" & strExc(2)
'                        strExc(2) = Dir()
'                    Loop
'                    '收件者
'                    strExc(3) = m_TCT04
'                    If m_TCT07 <> "" And m_TCT07 <> m_UserNo Then
'                        strExc(3) = strExc(3) & ";" & m_TCT07
'                    End If
'
'                    '主旨
'                    strExc(4) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) = "000", "", "-" & strCase(3) & "-" & strCase(4)) & "送回命名區，請主管重新確認"
'
'                    PUB_SendMail strUserNum, strExc(3), "", strExc(4), vbCrLf & "請參考附件", , IIf(strExc(1) = "", "", Mid(strExc(1), 2)), , , , strUserNum
'                End If
    
            End If
            Unload Me
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
   Frame6.BackColor = &H8000000F 'Added by Lydia 2020/02/21
   
   SSTab1.Tab = 0

   If m_iStiu = "Q" Or Val(m_TCT05) > 0 Then
         Me.Caption = "案件命名-已確認"
   End If
   
   'Added by Lydia 2018/07/12
   m_GrpManList = Pub_GetSt16Man(True) '所有工程師主管(含F編號)
   strResPath = Pub_GetSpecMan("FCP相似比對結果暫存")
   'Added by Lydia 2022/10/12 配合特殊情況之指定職代，增加判斷m_TCT04
   If InStr(m_GrpManList, m_TCT04) = 0 Then m_GrpManList = m_GrpManList & "," & m_TCT04
           
   'Added by Lydia 2019/09/02 是否啟用重送命名記錄的功能
   bolAgain = True
   'Remark by Lydia 2019/09/17 修改內容不要用PDF，直接在內文寫上修改前後。
   'If bolAgain = True Then
   '    '清除暫存PDF
   '    PUB_KillTempFile "$$*修改*.pdf"
   'End If
       
   Chk2(51).Top = 3345   'Added by Lydia 2023/03/10
   
   'Added by Lydia 2025/03/13 以後原本的4會改成Z
   If strSrvDate(1) < "20250314" Then Chk27(4).Visible = False
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PUB_SendMailCache 'Added by Lydia 2017/12/15
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set frm090902_2 = Nothing
   If TypeName(m_PrevForm) <> "Nothing" Then
        m_PrevForm.Show
        If m_PrevForm.Name = "frm090902" Then
            m_PrevForm.doQuery False
        'Added by Lydia 2019/09/02 重新查詢卷宗區
        ElseIf m_PrevForm.Name = "frm100101_L" Then
            m_PrevForm.ReadAttachFile
        End If
   End If
   Set m_PrevForm = Nothing
End Sub

'Modified by Lydia 2018/04/26 +是否檢查中英文名稱bChk
Public Function ReadData(Optional bChk As Boolean = False) As Boolean
Dim rsRD As New ADODB.Recordset
  
    '先清除案件名稱資料檔串列
    ClearTCTFieldList
    bAgainTrans = False 'Added by Lydia 2019/09/02
    
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
       'Added by Lydia 2020/02/17 基本檔：名稱有特殊字
       If strCase(12) = "Y" Then
          ChkPA174.Value = vbChecked
       End If
       
        'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地，外文本->原始檔區
        If PUB_ChkCPExist(strCase, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            CmdOpen.Caption = Replace(CmdOpen.Caption, "外文本", "原始檔")
            CmdOpen.Tag = strExc(1)
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
    'Added by Lydia 2017/12/29 主管確認：檢查專利檔和命名記錄檔的工程師組別
    'Modified by Lydia 2024/02/27 外專機械設計組人員異動調整程式：新案認領組別，請取消機械設計組，案件歸電子組; 由Wilson代機械組主管T1
    'If rInx = 0 And m_iStiu = "M" And strCase(7) <> m_UserSt16 Then
    'Mark by Lydia 2024/03/20 原因:機械組案件101,102由外專工程師(可能是電子組或化學組), 103由內專工程師翻譯
    'If rInx = 0 And m_iStiu = "M" And strCase(7) <> m_UserSt16 And strCase(7) <> "1" And strCase(7) <> "4" Then
    '     MsgBox "命名記錄檔的工程師組別和專利基本檔不一致，請通知程序人員到新案建檔設定工程師組別! "
    '     cmdOK(0).Enabled = False
    '     cmdOK(1).Enabled = False
    '     cmdOK(2).Enabled = False
    '     cmdOpen.Enabled = False
    'End If
    'end 2024/03/20
    rInx = rInx + 1
    'end 2017/12/29
    
    'Added by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
    s_TCT04m = Pub_GetFCPGrpMan(m_UserSt16) '以命名人員之工程師組別取得主管
    
    'Added by Lydia 2018/04/20 檢查基本檔和命名記錄的名稱
    'Modified by Lydia 2018/04/26 櫃台現在會幫忙Key英文名稱,所以改成分開檢查(主管確認預設不檢查)
'    If m_PA05 & m_PA06 <> "待命名" And m_PA05 & m_PA06 <> txtData(3).Text & txtData(4).Text Then
'        If MsgBox("命名作業的中英文名稱和專利基本檔不一致，是否代入專利基本檔的中英文名稱？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
'             txtData(3).Text = m_PA05
'             txtData(4).Text = m_PA06
'        End If
'    End If
    If bChk = True Then
       If Val(strCase(9)) = 0 Then  'Added by Lydia 2024/10/23 debug: 提申後，除案件名稱外，其餘皆可修改。ex.FCP-072369被命名工程師代入專利基本檔的中文名稱
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
       End If  'Added by Lydia 2024/10/23
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
    'Modified by Lydia 2018/04/19 +第2~5級管制人st52~st55(主管)
    'Modified by Lydia 2018/10/18 另外抓PA63,CP64
    'Modified by Lydia 2022/11/30 +抓該工程師組別的主管
    'Str01 = "SELECT A.*,s1.st02 as TCT10n,s1.st52||','||s1.st53||','||s1.st54||','||s1.st55 st5255,pa63,cp64 " & _
               "FROM TRANSCASETITLE A,staff s1,caseprogress,patent " & _
               "WHERE TCT01='" & m_TCT01 & "' and tct10=s1.st01(+) and tct10=s1.st01(+) and tct01=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
    'Modified by Lydia 2023/01/18 +pa75,pa26,pa27,pa28,pa29,pa30
    'Modified by Lydia 2023/03/10 + CP10,PA176
    'Modified by Lydia 2025/08/29 + CP06,CP142,CP164
    Str01 = "SELECT A.*,s1.st02 as TCT10n,s1.st52||','||s1.st53||','||s1.st54||','||s1.st55 st5255,oman,pa63,cp64,pa75,pa26,pa27,pa28,pa29,pa30,cp10,PA176,CP06,CP142,CP164 " & _
               "FROM TRANSCASETITLE A,staff s1,caseprogress,patent, setspecman " & _
               "WHERE TCT01='" & m_TCT01 & "' and tct10=s1.st01(+) and tct10=s1.st01(+) and tct01=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
               "and decode(s1.st16,'1','T','2','R','3','S','4','T1',s1.st16)=OCODE(+) "
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, Str01)
    If intA = 1 Then
        With rsA
           'Added by Lydia 2018/10/18
           m_PA63 = "" & rsA.Fields("pa63")  '客戶有提供彩圖
           m_TCT01cp64 = "" & rsA.Fields("cp64") '新案收文號的進度備註
           'end 2018/10/18
           m_TCT01cp10 = "" & rsA.Fields("cp10") 'Added by Lydia 2024/02/05 新案收文號之案件性質
           'Added by Lydia 2025/08/29
           m_TCT01cp06 = "" & rsA.Fields("cp06")  '本所期限
           m_TCT01cp164 = "" & rsA.Fields("cp164") '指定日期方式
           m_TCT01cp142 = "" & rsA.Fields("cp142") '指定送件日
           'end 2025/08/29
           
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
           'Added by Lydia 2022/11/30 改判斷為各組主管
           If "" & .Fields("oman") <> "" And "" & .Fields("TCT04") <> "" & .Fields("oman") Then
              m_TCT04 = "" & .Fields("oman")
           Else
           'end 2022/11/30
              m_TCT04 = "" & .Fields("TCT04")
           End If     'Added by Lydia 2022/11/30
           m_TCT04chk = "" & .Fields("TCT04") 'Added by Lydia 2023/05/31 目前空白表示處於第2次認領階段
           
           m_TCT05 = "" & .Fields("TCT05")
           '工程師主任
           m_TCT07 = "" & .Fields("TCT07")
           m_TCT08 = "" & .Fields("TCT08")            'Added by Lydia 2017/12/15 主任-確認日期
           '命名人員
           m_TCT10 = Trim("" & .Fields("TCT10"))
           lblData(8).Caption = Trim("" & .Fields("TCT10n")) 'Added by Lydia 2018/03/16 顯示命名人員
           m_TCT14 = "" & .Fields("TCT14")   'Added by Lydia 2019/09/19 重送確認記錄
           'Added by Lydia 2018/04/19 記錄告代和主動修正的狀態
           m_TCT20 = "" & .Fields("TCT20")
           m_TCT117 = "" & .Fields("TCT117")
           'Added by Lydia 2018/04/19 主管確認後,可修改人員
           m_UserList = "" & .Fields("st5255")
           If m_TCT07 <> "" Then m_UserList = IIf(InStr(m_UserList, m_TCT07) = 0, m_UserList & "," & m_TCT07, m_UserList)
           m_UserList = m_TCT10 & "," & Replace(m_UserList, ",,", ",")
           'Added by Lydia 2022/10/12 配合特殊情況之指定職代，增加判斷m_TCT04
           If InStr(m_UserList, m_TCT04) = 0 Then m_UserList = m_UserList & "," & m_TCT04
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
             m_TF20 = "" & rsA.Fields("tf20")
         Else
             m_TF20 = ""
         End If
         m_TF29 = "" & rsA.Fields("tf29") '待比對
         'end 2018/07/12
    End If
    'end 2018/06/01
    
    Set rsA = Nothing
    
    'Modified by Lydia 2017/12/27 主管可改全部
    'Call TxtEnabled
    'Modified by Lydia 2018/04/19 改卷宗區可修改
    'If m_iStiu <> "M" Then Call TxtEnabled
    Call TxtEnabled
    
    'Added by Lydia 2018/06/01 限新案翻譯才可輸入欲翻譯人員
    'Modified by Lydia 2018/08/21 日文組在命名作業認領經主管確認後，其他人員不可再認領，當命名完成後不可修改認領區塊，其他歐美組比照辦理
    'If m_TF01pty <> "201" Then
    'Modifeid by Lydia 2018/09/20 開放工程師主管可修改 (ex.FCP-59428)
    'If m_TF01pty <> "201" Or (m_iStiu = "Q" And Pub_StrUserSt03 <> "M51") Then
    If m_TF01pty <> "201" Or (m_iStiu = "Q" And InStr(m_GrpManList, strUserNum) = 0 And Pub_StrUserSt03 <> "M51") Then
        Frame4.Enabled = False
    Else
        Frame4.Enabled = True
    End If
    'end 2018/06/01
    
    'Added by Lydia 2019/09/19 與Phoebe,Sharon, 99025蘇韋寧討論: 提申前的「說明書」和「申請專利範圍 ＆ 圖示」都已經確認好了,
                                                                                                     '所以提申後只開放修改相似度,其他有修改Title或收文,請工程師與程序聯繫,避免工程師修改內容未註明而程序未能知悉。
    'Modified by Lydia 2019/10/01 9/23與Owen,Phoebe,Sharon討論:
                                            '1.命名後，在中說發文前可以修改命名記錄
                                            '2.提申後，除案件名稱外，其餘皆可修改。
                                            '3.只修改相似案號和相似度，不需重送命名。
    'If bolAgain = True And m_iStiu = "Q" And Val(strCase(9)) > 0 Then
    If bolAgain = True And m_iStiu = "Q" Then
        'Remove by Lydia 2020/09/04
        'If (m_TF01 <> "" And m_TF01cp27 <> "") Or (m_TF01 = "" And Val(strCase(9)) > 0) Then '有中說發文前 or 無中說提申後
    'end 2019/10/01
        '    txtData(3).Locked = True
        '    txtData(4).Locked = True
        '    txtData(5).Locked = True
        '    Frame1.Enabled = False  '設計案屬性
        '    Frame3.Enabled = False  '案件類別
        '    Frame4.Enabled = False  '認翻譯
        '    For intI = 0 To m_Frame5
        '        Frame5(intI).Enabled = False
        '    Next intI
        ''Added by Lydia 2019/10/01
        'End If
        'end 2020/09/04
        If Val(strCase(9)) > 0 Then '提申後，除案件名稱外，其餘皆可修改。
            txtData(3).Locked = True
            txtData(4).Locked = True
            txtData(5).Locked = True
            'Added by Lydia 2020/09/04 除案件名稱外，其餘皆可修改。
            Frame1.Enabled = True  '設計案屬性
            Frame3.Enabled = True  '案件類別
            Frame4.Enabled = True  '認翻譯
            For intI = 0 To m_Frame5
                Frame5(intI).Enabled = True
            Next intI
            'end 2020/09/04
            Frame6.Enabled = False 'Added by Lydia 2020/02/21 「名稱有特殊字」比照辦理
        End If
        'end 2019/10/01
    End If
    'end 2019/09/19
    
    'Added by Lydia 2018/10/01 工程師勾選”需收文主動修正”、”需收文告代”其中一項，若需修改或刪除，需有主管同意(工程師請主管進行修改)
    'Move by Lydia 2020/09/04 從上面移下來
    If m_iStiu = "Q" And Pub_StrUserSt03 <> "M51" Then
       If (InStr(m_UserList, strUserNum) = 0 Or strUserNum = m_TCT10) Then   '直屬主管或主管
            Frame5(8).Enabled = False
       Else
            Frame5(8).Enabled = True
       End If
       'Added by Lydia 2019/09/02 開放重送命名記錄=>可修改
       If bolAgain = True And InStr(m_UserList, strUserNum) > 0 Then
           Frame5(8).Enabled = True
       End If
    End If
    'end 2018/10/01
    'end 2020/09/04
    
    'Added by Lydia 2018/07/12 上傳相似比對結果檔案,新案翻譯才需要
    'Modified by Lydia 2018/08/09 工程師約定8/13上線
    'Modified by Lydia 2018/12/20 相似案號或相似度有值,就顯示
    'If strSrvDate(1) >= "20180813" And m_TF01 <> "" And m_TF01pty = "201" And txtData(6) <> "" And txtData(7) <> "" _
        And (m_iStiu = "M" Or cmdOK(3).Visible = True) Then
    If strSrvDate(1) >= "20180813" And m_TF01 <> "" And m_TF01pty = "201" And (txtData(6) <> "" Or txtData(7) <> "") _
        And (m_iStiu = "M" Or cmdOK(3).Visible = True) Then
        cmdFile.Visible = True
    End If
    
    'Added by Lydia 2023/01/18 命名作業不可新增告代和主動修正
    strNotBList = Pub_GetITSforHandle(strCase(1) & strCase(2) & strCase(3) & strCase(4), m_PA75, m_PA26 & "," & m_PA27 & "," & m_PA28 & "," & m_PA29 & "," & m_PA30)
    
    'Added by Lydia 2024/10/17 英文組的彩圖依照PA63設定--from Phoebe ; ex.FCP-72571 'Added by Lydia 2023/04/26 (電子電機、化學、機械 三組)直接以彩圖製作成ori版本提申(自112年5月1日起實施); 日文組仍維持原程式=>需判斷是否以彩圖提申
    If strSrvDate(1) >= "20230501" And m_iStiu = "M" And strCase(7) <> "3" And m_PA63 = "Y" Then
        Chk2(49).Value = vbChecked
    End If

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
bolChgText = False

On Error GoTo Err01
   Select Case nIndex
       Case 16 '案件名稱(中)
           If nType = "S" Then
              If nData <> "" Then
                 txtData(3).Text = nData
              '案件名稱(中)預設'待命名'
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
              strTempA = "案件名稱(中)：" & nData
              bolChgText = True
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
              strTempA = String(4, "　") & "(英)：" & nData
              bolChgText = True
           End If
       Case 18 '設計案屬性
           If nType = "S" Then
              If Val(nData) >= 1 And Val(nData) <= 4 Then
                 opt1(Val(nData) - 1).Value = 1
              End If
              If nMax > 0 Then
                 opt1(0).Tag = "TCT" & Format(nIndex, "00")
              End If
           ElseIf nType = "U" Then
              intM = 0
              For Each oOpt In opt1
                 If oOpt.Value = True Then intM = oOpt.Index + 1
              Next
              SetTCTFieldNewData opt1(0).Tag, IIf(intM = 0, Empty, intM)
           ElseIf nType = "W" Then
              strTempA = "設計案屬性："
              For Each oOpt In opt1
                 If nData <> "" And Val(nData) = oOpt.Index + 1 Then
                    strTempA = strTempA & "　" & sChked & oOpt.Caption
                 Else
                    strTempA = strTempA & "　" & sUnchked & oOpt.Caption
                 End If
              Next
              bolChgText = True
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
                   strTempA = ""
                   If nData = "Y" Then
                       strTempA = strTempA & "Y需收文主動修正|"
                   Else
                       strTempA = strTempA & "N需收文主動修正|"
                       If nData = "N" Then strTempA = "A" & strTempA
                   End If
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
                    For Each oOpt In Opt4s2
                       If nData <> "" And Val(nData) = oOpt.Index + 1 Then
                          strTempA = strTempA & "　Y" & oOpt.Caption & "|"
                       Else
                          strTempA = strTempA & "　N" & oOpt.Caption & "|"
                       End If
                    Next
                    '改為符號
                    strTempA = Replace(strTempA, "Y", sChked)
                    strTempA = Replace(strTempA, "N", sUnchked)
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
                   If nData = "Y" Then
                       strTempA = strTempA & "　Y不請款|"
                   Else
                       'Modified by Lydia 2018/04/18
                       'strTempA = strTempA
                       strTempA = strTempA & "　N不請款|"
                   End If
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
              If nData <> "" Then
                    strTempA = strTempA & "Y需收文告代|"
              Else
                    strTempA = strTempA & "N需收文告代|"
              End If
              For Each oOpt In Opt4s
                 If nData <> "" And Val(nData) = oOpt.Index + 1 Then
                    strTempA = strTempA & "　Y" & oOpt.Caption & "|"
                 Else
                    strTempA = strTempA & "　N" & oOpt.Caption & "|"
                 End If
              Next
              If bAgainTrans = False Then 'Added by Lydia 2019/09/17 判斷不是在讀取email內文
                  If Left(strTempA, 1) = "A" Then
                     strTempA = strTempA & "Y不需收文" & "|"
                     strTempA = Mid(strTempA, 2)
                  Else
                     strTempA = strTempA & "N不需收文" & "|"
                  End If
              End If
              '改為符號
              strTempA = Replace(strTempA, "Y", sChked)
              strTempA = Replace(strTempA, "N", sUnchked)
              bolChgText = True
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
              strTempA = "可一案兩請"
              If nData = "Y" Then
                 strTempA = sChked & strTempA
              Else
                 strTempA = sUnchked & strTempA
              End If
              'Remove by Lydia 2018/04/18 和彩圖提申(TCT118)同一行
              'bolChgText = True
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
              strTempA = "本案說明書內容與"
              If nData = "Y" Then
                 strTempA = sChked & strTempA
              Else
                 strTempA = sUnchked & strTempA
              End If
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
              If Trim(nData) <> "" Then
                 Call ChgCaseNo(nData, strExc)
                 strTempA = strTempA & strExc(1) & "-" & strExc(2) & IIf(strExc(3) & strExc(4) <> "000", strExc(3) & "-" & strExc(4), "")
              Else
                 'Modified by Lydia 2023/05/30
                 'strTempA = strTempA & "FCP-" & String(6, " ")
                 strTempA = strTempA & strCase(1) & "-" & String(6, " ")
              End If
              strTempA = strTempA & "之內容"
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
              If Trim(nData) <> "" Then
                 strTempA = strTempA & nData
              Else
                 strTempA = strTempA & String(2, " ")
              End If
              strTempA = strTempA & "%相同"
              bolChgText = True
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
              If nData <> "" Then
                  strTempA = "案件類別：" & Opt5(Val(nData) - 1).Caption
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & " " & nData
              End If
              bolChgText = True
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
              strTempA = "欲翻譯此案件者："
              If Right(nData, 1) = "A" Then
                 strTempA = strTempA & GetStaffName(Mid(nData, 1, InStr(nData, "-") - 1)) & "-下班翻譯"
              ElseIf Right(nData, 1) = "B" Then
                 strTempA = strTempA & GetStaffName(Mid(nData, 1, InStr(nData, "-") - 1)) & "-上班翻譯"
              ElseIf nData <> "" Then
                 strTempA = strTempA & Chk2(40 + Val(nData)).Caption
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "-" & nData
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = sChked & "說明書"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　" & sChked & "內容不完整"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "技術領域標題"
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "(" & Opt3_1(Val(nData) - 1).Caption & ")"
              End If
              strTempA = strTempA & IIf(strTempA <> "", "，位置：", "")
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "技術領域建議內容："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "先前技術標題"
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "(" & Opt3_2(Val(nData) - 1).Caption & ")"
              End If
              strTempA = strTempA & IIf(strTempA <> "", "，位置：", "")
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "先前技術建議內容："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "發明內容標題"
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "(" & Opt3_3(Val(nData) - 1).Caption & ")"
              End If
              strTempA = strTempA & IIf(strTempA <> "", "，位置：", "")
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "發明內容建議內容："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "圖式簡單說明標題"
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "(" & Opt3_4(Val(nData) - 1).Caption & ")"
              End If
              strTempA = strTempA & IIf(strTempA <> "", "，位置：", "")
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "圖式簡單說明建議內容："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "實施方式標題"
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "(" & Opt3_5(Val(nData) - 1).Caption & ")"
              End If
              strTempA = strTempA & IIf(strTempA <> "", "，位置：", "")
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "實施方式建議內容："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "符號說明標題"
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "(" & Opt3_6(Val(nData) - 1).Caption & ")"
              End If
              strTempA = strTempA & IIf(strTempA <> "", "，位置：", "")
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "符號說明建議內容："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　" & sChked & "缺摘要："
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "缺摘要建議內容："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　" & sChked & "缺頁"
              End If
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
              If nData <> "" Then
                 strTempA = strTempA & "　頁數：" & nData
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　" & sChked & "其它問題："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = sChked & "申請專利範圍"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "項號錯誤：請求項:"
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "依附關係錯誤(附屬項:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　")
                 strTempA = strTempA & "，應依附於請求項:"
              End If
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
              If strTempA <> "" Then
                  strTempA = strTempA & IIf(nData <> "", nData, "　") & ")"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "依附關係不明確(附屬項:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & "，應依附於請求項:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & ")"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "多附多(附屬項:" 'Memo by Lydia 2018/04/17 原"不當依附"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & "，應依附於請求項:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & ")"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "屬於引用記載形式之獨立項：請求項:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & ")"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "標的不一致"
              End If
              bolChgText = True
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
              strTempA = "　　　附屬項:"
              If nData = "" Then
                 strTempA = strTempA & "　，標的:"
              Else
                 strTempA = strTempA & nData & "，標的:"
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = "　　　被依附之請求項:"
              If nData = "" Then
                 strTempA = strTempA & "　，標的:"
              Else
                 strTempA = strTempA & nData & "，標的:"
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "不予專利(請求項:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & "，標的:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & "，法條:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & ")"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "混雜式請求項：請求項:"
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　" & sChked & "其它問題："
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = sChked & " 圖式"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "建議指定代表圖：圖"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & nData
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "缺圖"
              End If
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
              If strTempA <> "" And nData = "Y" Then
                 strTempA = strTempA & "(彩圖)"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & "　圖:" & nData
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "格式不符"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & "(圖:" & IIf(nData <> "", nData, "　") & "，說明:"
              End If
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
              If strTempA <> "" Then
                 strTempA = strTempA & IIf(nData <> "", nData, "　") & ")"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　" & sChked & "其它問題:"
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "不完整：圖:"
              End If
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
              strTempA = strTempA & nData
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "超過一個實施例"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "不主張設計的部分"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "色彩"
              End If
              bolChgText = True
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
              strTempA = ""
              If nData = "Y" Then
                 strTempA = "　　" & sChked & "用途說明"
              End If
              bolChgText = True
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
              If nData = "Y" Then
                 strTempA = strTempA & "　　　" & sChked & "彩圖提申"
              Else
                 strTempA = strTempA & "　　　" & sUnchked & "彩圖提申"
              End If
              bolChgText = True
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
              If nData = "Y" Then
                 strTempA = strTempA & "　　　" & sChked & "有序列表"
              Else
                 strTempA = strTempA & "　　　" & sUnchked & "有序列表"
              End If
              bolChgText = True
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
              If Chk2(51).Visible = True And Opt5(0).Value = True Then
                 If Chk2(51).Value = vbChecked Then
                    strM01 = "Y"
                 Else
                    strM01 = "N"
                 End If
                 SetTCTFieldNewData Chk2(51).Tag, strM01
              End If
           ElseIf nType = "W" Then
              If Chk2(51).Visible = True Then
                  If nData = "Y" Then
                     strTempA = strTempA & "　" & sChked & "專利權期間延長相關"
                  Else
                     strTempA = strTempA & "　" & sUnchked & "專利權期間延長相關"
                  End If
                  bolChgText = True
              End If
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
Dim strDefDir As String, strNewFileName As String
Dim strTmp As String
Dim strExSql As String
Dim nIndex As Integer
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim strPK As String
Dim strCP48 As String, strCP16 As String, strCP17  As String, strCP18  As String 'Added by Lydia 2018/04/18
Dim bUpd203 As Boolean, bUpd901 As Boolean 'Added by Lydia 2018/04/19 是否有變更主動修正和告代
Dim bUpdPA As Boolean 'Added by Lydia 2018/04/19 是否變更專利基本檔
Dim strTo As String 'Added by Lydia 2018/04/20
Dim dCP33 As Double, dCP34 As Double 'Added by Lydia 2018/05/10標準價和底價
Dim strCP26 As String 'Added by Lydia 2018/05/10 是否算案件
Dim tmpArr As Variant 'Added by Lydia 2018/09/27
Dim strDataBefore As String, strDataAfter As String 'Added by Lydia 2019/09/17 修改前和修改後的資料
Dim tmpArr1 As Variant, intP As Integer 'Added by Lydia 2019/09/17
Dim bUpdTCT21 As Boolean, bUpdTCT118 As Boolean 'Added by Lydia 2019/12/02 記錄一案兩請TCT21,彩圖提申TCT118
Dim strPA16 As String 'Added by Lydia 2020/03/27
Dim strCP06 As String 'Added by Lydia 2021/08/23 本所期限
Dim strCP20 As String 'Added by Lydia 2022/05/03
Dim strMailCp09 As String, strMailOld As String, strMailSub As String, strMailCont As String 'Added by Lydia 2023/05/10
Dim strCP06Old As String, strCP48Old As String 'Added by Lydia 2023/11/30 原本告代或主修的本所期限和承辦期限

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
                      'Modified by Lydia 2019/04/15  案號5碼補到6碼
                      'If Pub_FtpPutTyping2(strExc(3), strResPath & "/" & Mid(strExc(3), InStrRev(strExc(3), "\") + 1)) = False Then
                      strExc(2) = Mid(strExc(3), InStrRev(strExc(3), "\") + 1)
                      If Mid(UCase(strExc(2)), 1, Len(strCase(1) & Val(strCase(2)))) = UCase(strCase(1) & Val(strCase(2))) Then
                          strExc(2) = strCase(1) & strCase(2) & Mid(strExc(2), Len(strCase(1) & Val(strCase(2))) + 1)
                      End If
                      If Pub_FtpPutTyping2(strExc(3), strResPath & "/" & strExc(2)) = False Then
                      'end 2019/04/15
                          Exit Function
                      End If
                End If
      'Added by Lydia 2018/09/27
          End If
      Next nIndex
      'end 2018/09/27
   End If
   'end 2018/07/12
   
   'Added by Lydia 2020/03/27 目前案件准/駁
   'Modified by Lydia 2022/05/03 +FCP-067004
   'Modified by Lydia 2024/05/28 改成模組
   'If InStr("FCP062174000,FCP067004000", strCase(1) & strCase(2) & strCase(3) & strCase(4)) > 0 Then
   If PUB_GetCP20forSpec(strCase(1), strCase(2), strCase(3), strCase(4), "") = "N" Then
   'end 2024/05/28
       strExc(0) = "select pa16 from patent where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           strPA16 = "" & RsTemp.Fields("pa16")
       End If
   End If
   'end 2020/03/27
   
   OnSaveData = False
   ' 更新輸入欄位
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
         'Added by Lydia 2018/04/19 是否有變更主動修正和告代
         If InStr("TCT19,TCT116,TCT117", m_TCTList(nIndex).fiName) > 0 Then
              bUpd203 = True
         ElseIf InStr("TCT20,", m_TCTList(nIndex).fiName) > 0 Then
              bUpd901 = True
         ElseIf InStr("TCT16,TCT17,TCT18", m_TCTList(nIndex).fiName) > 0 Then '是否變更專利基本檔
              bUpdPA = True
         'Added by Lydia 2019/12/02 記錄一案兩請TCT21,彩圖提申TCT118
         ElseIf InStr("TCT21,", m_TCTList(nIndex).fiName) > 0 Then
              bUpdTCT21 = True
         ElseIf InStr("TCT118,", m_TCTList(nIndex).fiName) > 0 Then
              bUpdTCT118 = True
         'end 2019/12/02
         End If   'end 2018/04/19
                  
         'Added by Lydia 2019/09/02 命名人員可以從卷宗區修改命名記錄，重新送回命名區請主管重新確認(重送命名流程)
         '排除相似度、相似案的修改
         'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
         'If iType = "3" And bolAgain = True And InStr("TCT23,TCT24", m_TCTList(nIndex).fiName) = 0 And m_UserNo <> m_TCT04 Then
         If iType = "3" And bolAgain = True And InStr("TCT23,TCT24", m_TCTList(nIndex).fiName) = 0 And InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) = 0 Then
         'If iType = "3" And bolAgain = True And m_UserNo <> m_TCT04 Then 'Mark by 2019/09/17 修改相似度要重送命名 (2019/10/01保留)
             bAgainTrans = True
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
 
    'Added by Lydia 2019/09/02 命名人員可以從卷宗區修改命名記錄，重新送回命名區請主管重新確認(重送命名流程)
    'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
    'If iType = "3" And InStr(strExSql, "TCT") > 0 And bolAgain = True And bAgainTrans = True And m_UserNo <> m_TCT04 Then
    If iType = "3" And InStr(strExSql, "TCT") > 0 And bolAgain = True And bAgainTrans = True And InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) = 0 Then
        If MsgBox("修改命名記錄需要送回命名區請主管重新確認，" & vbCrLf & "請問是否繼續作業？", vbInformation + vbYesNo + vbDefaultButton2, "") = vbNo Then
            bAgainTrans = False
            Exit Function
        Else
            '先輸出原本PDF檔案
            'Modified by Lydia 2019/09/18 改成記錄修改前和修改後的資料,直接在email內文寫上修改前後
            'PUB_PrintTCTcon m_TCT01, Me.Combo1.Text, strPrinter, "$$" & strCase(1) & strCase(2) & IIf(strCase(3) & strCase(4) = "000", "", strCase(3) & strCase(4)) & "修改前"
            strDataBefore = ProcDataWord1 '修改前
            strDataAfter = ProcDataWord2 '修改後
            '經過逐行比對,只列出有異動的資料內容
            strExc(1) = "": strExc(2) = ""
            tmpArr = Empty: tmpArr1 = Empty
            tmpArr = Split(strDataBefore, vbCrLf)
            tmpArr1 = Split(strDataAfter, vbCrLf)
            If UBound(tmpArr1) > UBound(tmpArr) Then
                intI = UBound(tmpArr1)
            Else
                intI = UBound(tmpArr)
            End If
            For nIndex = 0 To intI
                 If nIndex <= UBound(tmpArr) Then '修改前=>strExc(1)
                     strExc(5) = tmpArr(nIndex)
                 Else
                     strExc(5) = ""
                 End If
                 If nIndex <= UBound(tmpArr1) Then '修改後=>strExc(2)
                     strExc(6) = tmpArr1(nIndex)
                 Else
                     strExc(6) = ""
                 End If
                 If strExc(5) <> strExc(6) Then
                      '主動修正和告代：先合併為一個字串，後逐行列印
                     If strExc(5) <> "" And InStr(strExc(5), "不請款|") > 0 Then
                         strExc(1) = strExc(1) & vbCrLf & IIf(Left(strExc(5), 1) = "A", sChked & "不需收文主動修正和告代", Replace(strExc(5), "|", ""))
                         strExc(2) = strExc(2) & vbCrLf & IIf(Left(strExc(6), 1) = "A", sChked & "不需收文主動修正和告代", Replace(strExc(6), "|", ""))
                        intP = InStr(strExc(1), "不請款")
                        If intP > 0 Then strExc(1) = Mid(strExc(1), 1, intP + 2) & vbCrLf & Mid(strExc(1), intP + 3)
                        intP = InStr(strExc(2), "不請款")
                        If intP > 0 Then strExc(2) = Mid(strExc(2), 1, intP + 2) & vbCrLf & Mid(strExc(2), intP + 3)
                     Else
                       strExc(1) = strExc(1) & vbCrLf & strExc(5)
                       strExc(2) = strExc(2) & vbCrLf & strExc(6)
                     End If
                 End If
            Next nIndex
            '收件者
            strExc(3) = m_TCT04
            strExc(5) = "" 'CC
            If m_TCT07 <> "" And m_TCT07 <> m_UserNo Then
                strExc(3) = strExc(3) & ";" & m_TCT07
            'Modified by Lydia 2023/02/19 命名人員之主管ST52修改也要CC命名人員; ex.FCP-68960
            'ElseIf m_UserNo <> m_TCT10 Then '主任修改: CC給命名人員
            End If
            If m_UserNo <> m_TCT10 Then
            'end 2023/02/19
                strExc(5) = m_TCT10
            End If
            'Added by Lydia 2023/02/19 通知命名修改的內容Email，請cc程序人員。
            If Val(n_CP27) = 0 Or (Val(n_CP27) > 0 And (InStr(UCase(strExSql), "TCT23") > 0 Or InStr(UCase(strExSql), "TCT24") > 0)) Then
               '提申後之工程師的修改: 相似度%有異動，請cc程序人員及Sharon
               strExc(8) = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4))
               If strExc(8) <> "" Then
                  strExc(5) = strExc(5) & IIf(strExc(5) <> "", ";", "") & strExc(8)
               End If
               If Val(n_CP27) > 0 Then
                  strExc(8) = Pub_GetSpecMan("M")
                  If strExc(8) <> "" Then strExc(5) = strExc(5) & IIf(strExc(5) <> "", ";", "") & strExc(8)
               End If
            End If
            'end 2023/02/19
            
            '主旨
            strExc(4) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) = "000", "", "-" & strCase(3) & "-" & strCase(4)) & "命名區資料異動，請主管至未命名區\待確認-重新確認"
            strExc(1) = ChgSQL(IIf(Trim(Replace(strExc(1), vbCrLf, "")) <> "", "修改前：" & strExc(1) & vbCrLf & vbCrLf & String(40, "=") & vbCrLf & vbCrLf, "") & _
                                         "修改後：" & strExc(2))
                        
            strExc(0) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                " values( '" & strUserNum & "','" & strExc(3) & "',to_char(sysdate,'yyyymmdd')" & _
                ",to_char(sysdate,'hh24miss'),'" & strExc(4) & "','" & strExc(1) & "'," & CNULL(strExc(5)) & ")"
            cnnConnection.Execute strExc(0)
            'end 2019/09/17
            
            '還原記錄
            strExSql = strExSql & ", TCT05=NULL, TCT06=NULL, TCT13=NULL" '主管確認日期,時間
            '判斷是否為主任
            If m_UserNo = m_TCT07 Then
               strExSql = strExSql & ",TCT08=" & strSrvDate(1) & ", TCT09=" & Mid(Format(ServerTime, "000000"), 1, 4)
            Else '命名人員
               strExSql = strExSql & ",TCT08=NULL, TCT09=NULL ,TCT11=" & strSrvDate(1) & ", TCT12=" & Mid(Format(ServerTime, "000000"), 1, 4)
            End If
            '重送確認記錄：只保留最新記錄，過去記錄請查詢DML_log => 系統日(操作者員工編號,變更修正/告代)
            strExc(1) = ""
            If bUpd203 = True Then strExc(1) = strExc(1) & ",變更修正"
            If bUpd901 = True Then strExc(1) = strExc(1) & ",變更告代"
            'Added by Lydia 2019/12/02 記錄一案兩請TCT21,彩圖提申TCT118
            If bUpdTCT21 = True Then strExc(1) = strExc(1) & ",一案兩請"
            If bUpdTCT118 = True Then strExc(1) = strExc(1) & ",彩圖提申"
            'end 2019/12/02
            strExSql = strExSql & ",TCT14='" & ChangeTStringToTDateString(strSrvDate(1)) & " " & Left(Format(ServerTime, "00:00:00"), 5) & " (" & strUserNum & strExc(1) & ");' "
        End If
    End If
    'end 2019/09/02
         
   '抓時間
   strExc(1) = Mid(Format(ServerTime, "000000"), 1, 4)

   '確認
   If iType = 1 Then
      'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
      'If m_UserNo = m_TCT04 And Val(m_TCT05) = 0 Then
      If InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 And Val(m_TCT05) = 0 Then
         strExSql = strExSql & IIf(bDifference = True, ", ", " ") & "TCT05=" & strSrvDate(1) & ", TCT06=" & Val(strExc(1)) & ", TCT13=" & CNULL(strUserNum)
         'Added by Lydia 2017/12/15 代主任確認
         If bolUpdMan2 = True Then
              strExSql = strExSql & ",TCT08=" & strSrvDate(1) & ", TCT09=" & Val(strExc(1))
         End If
         'end 2017/12/15
      Else
         strExSql = strExSql & IIf(bDifference = True, ", ", " ") & "TCT08=" & strSrvDate(1) & ", TCT09=" & Val(strExc(1))
         'Added by Lydia 2017/12/15 代主管確認
         If bolUpdMan = True Then
              strExSql = strExSql & ",TCT05=" & strSrvDate(1) & ", TCT06=" & Val(strExc(1)) & ", TCT13=" & CNULL(strUserNum)
         End If
         'end 2017/12/15
      End If
   ElseIf iType = 0 Then
   '退回
      'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
      'If m_UserNo = m_TCT04 Then
      If InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Then
         strExc(0) = "TCT05=NULL, TCT06=NULL, TCT08=NULL, TCT09= NULL,TCT13=NULL "
      Else
         strExc(0) = "TCT08=NULL, TCT09= NULL"
      End If
      strExSql = strExSql & IIf(bDifference = True, ", ", " ") & strExc(0) & ", TCT11=NULL, TCT12=NULL "
   End If
   
   'Added by Lydia 2018/09/20 判斷沒有變更,清空字串
   If InStr(strExSql, "TCT") = 0 Then
       strExSql = ""
   Else
   'end 2018/09/20
       strExSql = strExSql & " WHERE TCT01 = '" & m_TCT01 & "' "
   End If
   
    cnnConnection.BeginTrans
       'Modified by Lydia 2017/12/15　+代主管確認
       'If m_UserNo = m_TCT04 And Val(m_TCT05) = 0 And iType = 1 Then
       'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
       'If Val(m_TCT05) = 0 And iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True) Then
       If Val(m_TCT05) = 0 And iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) Then
            '新增卷宗區(RCD.menu)
            strExc(0) = "select count(*) cnt from casepaperpdf where cpp01='" & m_TCT01 & "' and instr(cpp02,'" & FCP命名記錄 & "') > 0 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp(0) = 0 Then
                  'Modified by Lydia 2025/11/17 cpp02改成本所案號.案件性質.RCD.Menu
                  'strExc(0) = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                      " values('" & m_TCT01 & "','" & m_TCT01 & "." & FCP命名記錄 & "',0,'" & strUserNum & "'," & strSrvDate(1) & "," & strExc(1) & "00" & "," & strSrvDate(1) & "," & strExc(1) & "00" & ",'Y')"
                  strExc(0) = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                      " values('" & m_TCT01 & "','" & strCase(1) & strCase(2) & IIf(strCase(3) & strCase(4) = "000", "", strCase(3) & strCase(4)) & "." & m_TCT01cp10 & "." & FCP命名記錄 & "',0,'" & strUserNum & "'," & strSrvDate(1) & "," & strExc(1) & "00" & "," & strSrvDate(1) & "," & strExc(1) & "00" & ",'Y')"
                  cnnConnection.Execute strExc(0), intI
               'Added by Lydia 2018/04/20 更新時間
               Else
                    strExc(0) = "update casepaperpdf set cpp08=" & strSrvDate(1) & ",cpp09=" & Format(ServerTime, "000000") & _
                                     "  where cpp01='" & m_TCT01 & "' and instr(cpp02,'" & FCP命名記錄 & "') > 0 "
                    cnnConnection.Execute strExc(0), intI
               'end 2018/04/20
               End If
            End If
            
            'Move by Lydia 2018/07/13 移到下面(一案兩請和彩圖提申通知)
            
            'Added by Lydia 2018/04/25 因為新申請案接洽單可直接收回代902和主動修正203,所以命名作業的主管確認自動掛承辦人=命名人員並且上已分案
            'Modified by Lydia 2022/04/06 增加收文A類901告代
            'Modified by Lydia 2022/04/28 增加: 加速審查422,高速審查431
            'Modified by Lydia 2023/05/10 保留FCP案的告代901和主動修正203；因FCP新案急件重新認領，修改進度檔若有提申後告代、主動修正再發一次mail通知舊和新承辦人之事。
            'strExc(4) = "Update caseprogress set cp14='" & m_TCT10 & "', cp122='Y' where " & ChgCaseprogress(strCase(1) & strCase(2) & strCase(3) & strCase(4)) & _
                              " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in ('902','203','901','422','431') "
            strExc(4) = "Update caseprogress set cp14='" & m_TCT10 & "', cp122='Y' where " & ChgCaseprogress(strCase(1) & strCase(2) & strCase(3) & strCase(4)) & _
                              " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in (" & GetAddStr(IIf(strCase(1) = "FCP", Replace(TCTforCP14, "203,901,", ""), TCTforCP14)) & ") "
            cnnConnection.Execute strExc(4), intI
            'end 2018/04/25
            'Added by Lydia 2024/02/05 FMP案外觀設計申請（103）若命名完成，請將 撰稿 （210）的承辦人，update成命名工程師 --- Phoebe
            If strCase(1) = "P" And m_TCT01cp10 = "103" Then
               strExc(4) = "Update caseprogress set cp14='" & m_TCT10 & "', cp122='Y' where " & ChgCaseprogress(strCase(1) & strCase(2) & strCase(3) & strCase(4)) & _
                                 " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in ('210') "
               cnnConnection.Execute strExc(4), intI
            End If
            'end 2024/02/05
            
            'Added by Lydia 2023/05/10 FCP案急件新案重新認領的告代901和主動修正203
            strExc(0) = "select cp09,cp10,cp06,cp48,cp14,cp64 from caseprogress where " & ChgCaseprogress(strCase(1) & strCase(2) & strCase(3) & strCase(4)) & _
                             "  and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in ('203','901') "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                   strExc(5) = IIf("" & RsTemp.Fields("cp10") = "901", "告代", "主動修正")
                   If strCase(1) = "FCP" And Val(strCase(9)) > 0 And "" & RsTemp.Fields("cp14") <> m_TCT10 And "" & RsTemp.Fields("cp14") <> "" And InStr("" & RsTemp.Fields("cp64"), "提申後" & strExc(5)) > 0 Then
                       If InStr(strMailCp09 & ",", "" & RsTemp.Fields("cp09")) = 0 Then
                           strMailCp09 = strMailCp09 & "," & RsTemp.Fields("cp09")
                           If InStr(strMailOld & ";", "" & RsTemp.Fields("cp14")) = 0 Then
                              strMailOld = strMailOld & "';" & RsTemp.Fields("cp14")
                           End If
                           strMailSub = strMailSub & "、提申後" & strExc(5)
                           strMailCont = strMailCont & "  提申後" & IIf(Len(strExc(5)) = 2, "　　", "") & "　　承辦期限：" & ChangeWStringToTDateString("" & RsTemp.Fields("cp48")) & "　本所期限：" & ChangeWStringToTDateString("" & RsTemp.Fields("cp06")) & vbCrLf
                       End If
                   End If
                   strExc(4) = "Update caseprogress set cp14='" & m_TCT10 & "', cp122='Y' where " & ChgCaseprogress(strCase(1) & strCase(2) & strCase(3) & strCase(4)) & _
                                    " and cp158=0 and cp159=0 and cp09='" & RsTemp.Fields("cp09") & "' "
                   cnnConnection.Execute strExc(4), intI
                   RsTemp.MoveNext
               Loop
            End If
            'end 2023/05/10
            
            'Added by Lydia 2018/08/06 設計案(103,125)命名作業確定時，有製作外文提申本242進度承辦人請自動帶命名人員及自動分案上"Y"。
            If strCase(5) = "3" Then
                'Added by Lydia 2025/08/29 目前"（242）製作提申外文本"未掛任何期限，會導致工程師在期限彈跳視窗無法彈出的情況，故請增加此性質的本所期限和承辦期限
                strExc(2) = ""
                If m_TCT01cp06 <> "" Then
                  '本所期限=新案提申的本所期限，承辦期限=本所期限往前-5個工作天
                  '若新案有指定送件日當天or之前=>本所期限=新案提申的本所期限(本所期限若超過指定送件日, 則以指定送件日當本所期限)，承辦期限=指定送件日往前-5個工作天
                   strExc(0) = m_TCT01cp06
                   If (m_TCT01cp164 = "1" Or m_TCT01cp164 = "2") And m_TCT01cp06 > m_TCT01cp142 And m_TCT01cp142 <> "" Then
                     strExc(0) = m_TCT01cp142
                   End If
                   If strExc(0) <= strSrvDate(1) Then
                      strExc(0) = strSrvDate(1)
                   Else
                      strExc(1) = CompWorkDay(6, strExc(0), 1)
                   End If
                   If strExc(1) <= strSrvDate(1) Then strExc(1) = strSrvDate(1)
                   strExc(2) = ", CP06=" & strExc(0) & ", CP48=" & strExc(1)
                End If
               'end 2025/08/29
               'Modified by Lydia 2025/08/29 預設本所期限和承辦期限 strExc(2)
                strExc(4) = "Update caseprogress set cp14='" & m_TCT10 & "' " & strExc(2) & ", cp122='Y' where " & ChgCaseprogress(strCase(1) & strCase(2) & strCase(3) & strCase(4)) & _
                                  " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 ='242' "
                cnnConnection.Execute strExc(4), intI
                '有製作中說210進度承辦人請自動帶命名人員
                'Mark by Lydia 2018/08/07 因怕程序人員看到已掛承辦人就忘了要去分案，故請將製作中說210進度承辦人自動帶命名人員改在分案時
                'strExc(4) = "Update caseprogress set cp14='" & m_TCT10 & "' where " & ChgCaseprogress(strCase(1) & strCase(2) & strCase(3) & strCase(4)) & _
                                  " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 ='210' "
                'cnnConnection.Execute strExc(4), intI
                'end 2018/08/07
            End If
            'end 2018/08/06
       End If
       
       'Move by Lydia 2022/04/28 從「提申後從命名系統修改專利名稱則有欄位註記」上方移過來
       'Modified by Lydia 2019/09/02 修改命名記錄，不用重新跑命名流程
       'Added by  Lydia 2018/04/19 卷宗區(RCD.menu)->存檔,更新修改日期
       'If bDifference = True And iType = 3 Then
       If bDifference = True And iType = 3 And bAgainTrans = False Then
            strExc(0) = "update casepaperpdf set cpp08=" & strSrvDate(1) & ",cpp09=" & Format(ServerTime, "000000") & _
                             "  where cpp01='" & m_TCT01 & "' and instr(cpp02,'" & FCP命名記錄 & "') > 0 "
            cnnConnection.Execute strExc(0), intI
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strExSql '維護log (命名記錄)
            Pub_SeekTbLog strExSql, , , , Me.Caption & "(" & Me.Name & ")"
       'Added by Lydia 2019/09/02 修改命名記錄，需要重新跑命名流程=>刪除卷宗區.Menu記錄
       ElseIf iType = 3 And bAgainTrans = True And bolAgain = True Then
            strExc(0) = "delete from casepaperpdf where cpp01='" & m_TCT01 & "' and instr(cpp02,'" & FCP命名記錄 & "') > 0 "
            cnnConnection.Execute strExc(0), intI
       'end 2019/09/02
       End If
        
       'Move by Lydia 2022/04/28 從「外專日文組新案無卷命名email設定」上方移過來
       'Modified by Lydia 2019/09/19
       'If strExSql <> "" Then cnnConnection.Execute strExSql, intI '更新命名記錄 'Modified by Lydia 2018/09/20 +判斷非空白
       If strExSql <> "" Then
           If bolAgain = True And bAgainTrans = True Then
                'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                'Pub_SeekTbLog strExSql, , True '若是有重送命名流程, 記錄log
                'Modified by Lydia 2025/10/30 改用模組判斷
                'Pub_SeekTbLog strExSql, , True, , Me.Caption & "(" & Me.Name & ")"
                Pub_SeekTbLog strExSql, , PUB_FilterSeekSQL(strExSql), , Me.Caption & "(" & Me.Name & ")"
           End If
           cnnConnection.Execute strExSql, intI
       End If
       
       'Move by Lydia 2018/04/19 從新增卷宗區(RCD.menu)下方移下來，並且重新整理
       If iType = 1 Or iType = 3 Then
            '更新基本檔
            'Modified by Lydia 2018/08/20 debug 主管確認
            'If iType = 1 Or (iType = 3 And bUpdPA = True) Then
            'Modified by Lydia 2019/09/24 限主管確認或主管修改存檔
            'If (iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) Or (iType = 3 And bUpdPA = True) Then
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If (iType = 1 Or iType = 3) And (m_UserNo = m_TCT04 Or bolUpdMan = True) Then
            If (iType = 1 Or iType = 3) And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) Then
                strExc(2) = ""
                If Frame1.Visible = True Then
                   intI = 0
                   For Each oOpt In opt1
                      If oOpt.Value = True Then intI = oOpt.Index + 1
                   Next
                   If intI > 0 Then
                      strExc(2) = ", pa158=" & CNULL("" & intI, True)
                   End If
                End If
                'Modifeid by Lydia 2018/06/26 去除單引號
                'strExc(0) = "update patent set pa05=" & CNULL(txtData(3).Text) & ", pa06=" & CNULL(txtData(4).Text) & strExc(2) & _
                            " where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' "
                'Modified by Lydia 2020/02/17  存檔「名稱有特殊字」
                'strExc(0) = "update patent set pa05=" & CNULL(ChgSQL(txtData(3).Text)) & ", pa06=" & CNULL(ChgSQL(txtData(4).Text)) & strExc(2) & _
                            " where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' "
                strExc(0) = "update patent set pa05=" & CNULL(ChgSQL(txtData(3).Text)) & ", pa06=" & CNULL(ChgSQL(txtData(4).Text)) & strExc(2)
                If strCase(12) <> IIf(ChkPA174.Value = 0, "", "Y") Then
                    strExc(0) = strExc(0) & " , pa174=" & CNULL(IIf(ChkPA174.Value = 0, "", "Y"))
                End If
                'Added by Lydia 2021/04/09 有序列表
                strExc(0) = strExc(0) & ", pa175=" & CNULL(IIf(Chk2(50).Value = 0, "", "Y"))
                'Added by Lydia 2023/03/10 FMP大陸新藥發明專利權期限補償控管
                If Chk2(51).Visible = True And Opt5(0).Value = True Then
                    strExc(0) = strExc(0) & ", PA176=" & CNULL(IIf(Chk2(51).Value = 1, "Y", "N"))
                End If
                'end 2023/03/10
                'Added by Lydia 2024/03/20 外專機械設計組人員異動調整程式：原因:機械組案件101,102由外專工程師(可能是電子組或化學組), 103由內專工程師翻譯
                'Modified by Lydia 2024/04/11 排除日文組案件 +strCase(7) <> "3"
                If Opt5(5).Value = True And strCase(7) <> "3" Then
                   strExc(0) = strExc(0) & ", PA150='4' "
                   strCase(7) = "4"
                End If
                'end 2024/03/20
                strExc(0) = strExc(0) & " where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' "
                'end 2020/02/17
                'Modified by Lydia 2018/10/19 +詳細記錄
                'Pub_SeekTbLog strExc(0)
                'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                'Pub_SeekTbLog strExc(0), , True
                'Modified by Lydia 2025/10/30 改用模組判斷
                'Pub_SeekTbLog strExc(0), , True, , Me.Caption & "(" & Me.Name & ")"
                Pub_SeekTbLog strExc(0), , PUB_FilterSeekSQL(strExc(0)), , Me.Caption & "(" & Me.Name & ")"
                cnnConnection.Execute strExc(0), intI '記錄log
            End If
            
            'Added by Lydia 2019/09/23 (主管確認)用重送確認記錄判斷
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If bolAgain = True And m_TCT14 <> "" And InStr(m_TCT14, "變更修正") > 0 And iType = 1 And Val(m_TCT05) = 0 And (m_UserNo = m_TCT04 Or bolUpdMan = True) Then
            If bolAgain = True And m_TCT14 <> "" And InStr(m_TCT14, "變更修正") > 0 And iType = 1 And Val(m_TCT05) = 0 _
                And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) Then
                   bUpd203 = True
            End If
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If bolAgain = True And m_TCT14 <> "" And InStr(m_TCT14, "變更告代") > 0 And iType = 1 And Val(m_TCT05) = 0 And (m_UserNo = m_TCT04 Or bolUpdMan = True) Then
            If bolAgain = True And m_TCT14 <> "" And InStr(m_TCT14, "變更告代") > 0 And iType = 1 And Val(m_TCT05) = 0 _
                 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) Then
                   bUpd901 = True
            End If
            'Added by Lydia 2019/10/04 第一次主管確認,預設要判斷收文
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If m_TCT14 = "" And iType = 1 And Val(m_TCT05) = 0 And (m_UserNo = m_TCT04 Or bolUpdMan = True) Then
            If m_TCT14 = "" And iType = 1 And Val(m_TCT05) = 0 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) Then
                   bUpd203 = True
                   bUpd901 = True
            End If
            
            '主動修正203
            'Modified by Lydia 2018/08/20 debug 主管確認才產生B類收文(ex.FCP-59439)
            'If iType = 1 Or (iType = 3 And bUpd203 = True) Then
            'Modified by Lydia 2019/09/24 改判斷
            'If (iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) Or (iType = 3 And bUpd203 = True) Then
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If (iType = 1 Or iType = 3) And (m_UserNo = m_TCT04 Or bolUpdMan = True) And bUpd203 = True Then
            If (iType = 1 Or iType = 3) And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) And bUpd203 = True Then
                  'Modified by Lydia 2022/04/28 改成共用模組
                  'If ChkCPisExist(strCase, "203", strCP203, strExc(3)) = True Then
                  'Moddified by Lydia 2023/05/10 回傳承辦
                  'If PUB_ChkBCPisExist(strCase, "203", strCP203, strExc(3)) = True Then
                  If PUB_ChkBCPisExist(strCase, "203", strCP203, strExc(3), , , strExc(4)) = True Then
                     'Added by Lydia 2023/11/30 抓原本的本所期限和承辦期限
                     strCP06Old = "": strCP48Old = ""
                     strExc(0) = "select cp06,cp48 from caseprogress where cp09='" & strCP203 & "' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strCP06Old = "" & RsTemp.Fields("cp06")
                        strCP48Old = "" & RsTemp.Fields("cp48")
                     End If
                     'end 2023/11/30
                  End If
                  '收文
                  If Chk2(45).Value = vbChecked Then
                        '承辦期限
                        'Modified by Lydia 2022/04/28 FMP預設承辦期限和所限
                        'If strCase(1) = "FCP" Then 'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
                        If strCase(1) = "FCP" Or strCase(1) = "P" Then
                            'Modified by Ldia 2021/08/23 +本所期限
                            'Modified by Lydia 2022/08/04 改成共用模組
                            'strCP48 = GetCP48(IIf(Opt4s2(0).Value = True, "1", "2"), "203", strCP06)
                            strCP48 = PUB_GetTCTbCP48(IIf(Opt4s2(0).Value = True, "1", "2"), strCase(1), strCase(2), strCase(3), strCase(4), strCase(6), strCase(9), lblData(5).Caption, lblData(1).Caption, "203", strCP06, tCP06, tCP27)
                        Else 'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
                            strCP48 = ""
                            strCP06 = "" 'Added by Lydia 2021/08/23
                        End If
                        'end 2018/05/10

                        'Added by Lydia 2022/05/03 是否向客戶收款
                        strCP20 = PUB_GetCP20(strCase(1), "203")
                        'Modified by Lydia 2024/05/28 改成模組
                        ''FCP-062174審定前不收費控制: 判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
                        'If strPA16 = "" And InStr("FCP062174000", strCase(1) & strCase(2) & strCase(3) & strCase(4)) > 0 Then
                        '    strCP20 = "N"
                        'End If
                        '' FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
                        'If strPA16 <> "1" And InStr("FCP067004000", strCase(1) & strCase(2) & strCase(3) & strCase(4)) > 0 Then
                        If PUB_GetCP20forSpec(strCase(1), strCase(2), strCase(3), strCase(4), strPA16) = "N" Then
                        'end 2024/05/28
                            strCP20 = "N"
                        End If
                        'end 2022/05/03
                        
                        '計算費用
                        'Modified by Lydia 2022/05/03 統一用OnUpdateFee
                        'Call OnUpdateFee("203", IIf(Chk2(48).Value = vbChecked, "N", ""), strCP16, strCP17, strCP18)
                        'Added by Lydia 2020/03/27 FCP-062174審定前不收費控制: 判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
                        'If strPA16 = "" And InStr("FCP062174000", strCase(1) & strCase(2) & strCase(3) & strCase(4)) > 0 Then
                        '     strCP16 = "": strCP17 = "": strCP18 = ""
                        'End If
                        ''end 2020/03/27
                        Call OnUpdateFee("203", strCP20, strCP16, strCP17, strCP18)
                        'end 2022/05/03
                        'Added by Lydia 2018/05/10 標準價和底價
                        If ClsPDGetCaseLowPrice(strCase(1), strCase(6), "203", dCP33, dCP34) = 1 Then
                        End If
                        
                        'Modified by Lydia 2019/09/19 +人工判斷
                        'If strCP203 = "" Then
                        'Mark by Lydia 2019/09/23 保留
                        'If strCP203 = "" And bolAdd203 = True Then
                        If strCP203 = "" Then
                            strPK = AutoNo("B", 6)
                            'Modified by Lydia 2025/06/25
                            'Pub_SetPAIsCase strCase(1), "203", strCP26 'Added by Lydia 2018/05/10 是否算案件數
                             If PUB_GetCPMbyCP10(strCase(1), "203", "cpm05") = "N" Then
                                strCP26 = "N"
                             End If
                             'end 2025/06/25
                             
                            'Add By Sindy 2021/6/18 非智慧局期限，要掛本所期限
                            'Remove by Lydia 2021/08/23 改用GetCP48取得strCP06
                            'Call GetPrjState6HM(strCase(1), "203", "cpm34", strExc(0))
                            'strExc(6) = "" '本所期限
                            'If Val(strCP48) > 0 And strExc(0) = "N" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                            '   strExc(6) = PUB_GetFCPOurDeadline(DBDATE(strCP48), , , , "N")
                            'End If
                            ''2021/6/18 END
                            'end 2021/08/23
                            
                            'Modified by Lydia 2018/04/25 +已分案CP122=Y
                            'Modified by Lydia 2018/05/10 +cp26,cp33,cp34
                            'Modified by Lydia 2018/08/22 +CP118 (若新案為電子送件,則主動修正也設為電子送件; ex.FC-58733)
                            'Modified by Lydia 2019/01/29 cp05收文日改為系統日(ex.FCP-60080的B類主動修正和告代之承辦期限:因為新案收文在107/12/11，但是ORI在108/1/19才來，所以1/21命名作業產生B類收文的期限為107/12/11+5或15個工作天；目前決定類似案例採人工修改期限。)
                            'Modify By Sindy 2021/6/18 + ,cp06
                            'Modified by Lydia 2021/08/23 strexc(6)=>strCP06 (GetCP48取得)
                            'Modified by Lydia 2022/05/03 CP20改用模組 IIf(Chk2(48).Value = vbChecked, "N", "")=> strCP20
                            strExc(0) = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp16,cp17,cp18,cp79,cp48,cp64,cp122,cp26,cp33,cp34,cp118,cp06) " & _
                                     "select cp01,cp02,cp03,cp04," & strSrvDate(1) & ", '" & strPK & "','203',cp11,cp12,cp13," & CNULL(m_TCT10) & "," & CNULL(strCP20) & _
                                     ", " & CNULL(strCP16, True) & ", " & CNULL(strCP17, True) & ", " & CNULL(strCP18, True) & ", " & CNULL(strCP16, True) & ", " & CNULL(strCP48, True) & _
                                     ", '" & ChangeWStringToWDateString(strSrvDate(1)) & IIf(Opt4s2(0).Value = True, " 命名-提申後主動修正;", " 命名-提申前主動修正;") & "', 'Y' " & _
                                     ", " & CNULL(strCP26) & ", " & dCP33 & ", " & dCP34 & ", " & CNULL(IIf(n_CP118 <> "", "Y", "")) & ", " & CNULL(strCP06, True) & " from caseprogress where cp09='" & m_TCT01 & "' "
                            cnnConnection.Execute strExc(0), intI
                            
                        'Modified by Lydia 2019/09/19 +判斷有未發文的收文才修改
                        ElseIf strCP203 <> "" Then
                            strExc(1) = ""
                            If Opt4s2(0).Value = True Then
                                strExc(2) = "1"
                                'Modified by Lydia 2021/11/04 本所期限cp06也要更新
                                strExc(1) = ", cp06 = " & CNULL(strCP06, True) & ", cp48 = " & CNULL(strCP48, True) & ", cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 命名-提申後主動修正;'||cp64 "
                            ElseIf Opt4s2(1).Value = True Then
                                strExc(2) = "2"
                                'Modified by Lydia 2021/11/04 本所期限cp06也要更新
                                strExc(1) = ", cp06 = " & CNULL(strCP06, True) & ", cp48 = " & CNULL(strCP48, True) & ", cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 命名-提申前主動修正;'||cp64 "
                            End If
                            'Added by Lydia 2023/05/10 FCP案急件新案重新認領的告代901和主動修正203
                            If strCase(1) = "FCP" And Val(strCase(9)) > 0 And strExc(4) <> m_TCT10 And strExc(4) <> "" And InStr(strExc(1), "提申後主動修正") > 0 Then
                                 If InStr(strMailCp09 & ",", strCP203) = 0 Then
                                     strMailCp09 = strMailCp09 & "," & strCP203
                                     If InStr(strMailOld, strExc(4)) = 0 Then
                                        strMailOld = strMailOld & "';" & strExc(4)
                                     End If
                                     strMailSub = strMailSub & "、提申後主動修正"
                                     strMailCont = strMailCont & "  提申後主動修正　　承辦期限：" & ChangeWStringToTDateString(strCP48) & "　本所期限：" & ChangeWStringToTDateString(strCP06) & vbCrLf
                                 End If
                            End If
                            'end 2023/05/10
                            
                            'Modified by Lydia 2019/09/24 重送命名流程,預設改備註
                            'If m_TCT117 = strExc(2) Then strExc(1) = "" '沒有變更,不改變備註
                            'Modified by Lydia 2023/11/30 判斷沒有改期限才清限;ex.FCP-70655分別在11/14,11/30跑命名
                            'If m_TCT117 = strExc(2) And m_TCT14 = "" Then strExc(1) = ""
                            If m_TCT117 = strExc(2) And m_TCT14 = "" And strCP06Old = strCP06 And strCP48Old = strCP48 Then strExc(1) = ""
                            
                            'Modified by Lydia 2018/04/25 +已分案CP122=Y
                            strExc(0) = "update caseprogress set cp14='" & m_TCT10 & "', cp20='" & IIf(Chk2(48).Value = vbChecked, "N", "") & "' " & _
                                              ", cp16 = " & CNULL(strCP16, True) & ", cp17 = " & CNULL(strCP17, True) & " ,cp18 = " & CNULL(strCP18, True) & _
                                              strExc(1) & ", CP122='Y' where cp09='" & strCP203 & "'  and cp158=0 "
                                  'Modified by Lydia 2019/09/24 +Update
                            'cnnConnection.Execute strExc(0), intI
                            cnnConnection.Execute "begin user_data.user_enabled:=1; " & strExc(0) & "; end ;", intI
                        End If
                  '刪除收文
                  ElseIf strCP203 <> "" Then
                        strExc(0) = "insert into DataDeleteRecord(dd01,dd02,dd03,dd04,dd14,dd15,dd16,dd17,dd18,dd19,dd20,dd21,dd22,dd23,dd24,dd25,dd26,dd27,dd28) " & _
                                          "select cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,cp05,cp13,cp16,cp17,cp60,'" & strUserNum & "','工程師命名-刪除收文',cp66,cp65,'" & strSrvDate(1) & "',mno " & _
                                          "from caseprogress,(SELECT '" & strCP203 & "' fno,(MAX(DD28)+1) mno FROM DATADELETERECORD) x " & _
                                          "where cp09='" & strCP203 & "' and cp09=fno "
                        cnnConnection.Execute strExc(0), intI
                        
                        strExc(0) = "delete from caseprogress where cp09='" & strCP203 & "' and cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' "
                        'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                        'Pub_SeekTbLog strExc(0)
                        Pub_SeekTbLog strExc(0), , , , Me.Caption & "(" & Me.Name & ")"
                        cnnConnection.Execute strExc(0), intI
                  End If
            End If
            '告代901
            'Modified by Lydia 2018/08/20 debug 主管確認才產生B類收文(ex.FCP-59439)
            'If iType = 1 Or (iType = 3 And bUpd901 = True) Then
            'Modified by Lydia 2019/09/24 改判斷
            'If (iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) Or (iType = 3 And bUpd901 = True) Then
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If (iType = 1 Or iType = 3) And (m_UserNo = m_TCT04 Or bolUpdMan = True) And bUpd901 = True Then
            If (iType = 1 Or iType = 3) And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True) And bUpd901 = True Then
                  'Modified by Lydia 2022/04/28 改成共用模組
                  'If ChkCPisExist(strCase, "901", strCP901, strExc(3)) = True Then
                  'Modified by Lydia 2023/05/10 回傳承辦
                  'If PUB_ChkBCPisExist(strCase, "901", strCP901, strExc(3)) = True Then
                  If PUB_ChkBCPisExist(strCase, "901", strCP901, strExc(3), , , strExc(4)) = True Then
                     'Added by Lydia 2023/11/30 抓原本的本所期限和承辦期限
                     strCP06Old = "": strCP48Old = ""
                     strExc(0) = "select cp06,cp48 from caseprogress where cp09='" & strCP901 & "' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strCP06Old = "" & RsTemp.Fields("cp06")
                        strCP48Old = "" & RsTemp.Fields("cp48")
                     End If
                     'end 2023/11/30
                  End If
                  '收文
                  If Chk2(47).Value = vbChecked Then
                        '承辦期限
                        'Modified by Lydia 2022/04/28 FMP預設承辦期限和所限
                        'If strCase(1) = "FCP" Then 'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
                        If strCase(1) = "FCP" Or strCase(1) = "P" Then
                            'Modified by Ldia 2021/08/23 +本所期限
                            'Modified by Lydia 2022/08/04 改成共用模組
                            'strCP48 = GetCP48(IIf(Opt4s(0).Value = True, "1", IIf(Opt4s(1).Value = True, "2", "3")), "901", strCP06)
                            strCP48 = PUB_GetTCTbCP48(IIf(Opt4s(0).Value = True, "1", "2"), strCase(1), strCase(2), strCase(3), strCase(4), strCase(6), strCase(9), lblData(5).Caption, lblData(1).Caption, "901", strCP06, tCP06, tCP27)
                        Else 'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
                            strCP48 = ""
                            strCP06 = "" 'Added by Lydia 2021/11/04
                        End If
                        'end 2018/05/10
                        
                        'Added by Lydia 2022/05/03 是否向客戶收款
                        strCP20 = PUB_GetCP20(strCase(1), "901")
                        'Modified by Lydia 2024/05/28 改成模組
                        ''FCP-062174審定前不收費控制: 判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
                        'If strPA16 = "" And InStr("FCP062174000", strCase(1) & strCase(2) & strCase(3) & strCase(4)) > 0 Then
                        '    strCP20 = "N"
                        'End If
                        '' FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
                        'If strPA16 <> "1" And InStr("FCP067004000", strCase(1) & strCase(2) & strCase(3) & strCase(4)) > 0 Then
                        If PUB_GetCP20forSpec(strCase(1), strCase(2), strCase(3), strCase(4), strPA16) = "N" Then
                        'end 2024/05/28
                            strCP20 = "N"
                        End If
                        'end 2022/05/03
                        
                        '計算費用
                        'Modified by Lydia 2022/05/03 統一用OnUpdateFee
                        'Call OnUpdateFee("901", "", strCP16, strCP17, strCP18)
                        ''Added by Lydia 2020/03/27 FCP-062174審定前不收費控制: 判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
                        'If strPA16 = "" And InStr("FCP062174000", strCase(1) & strCase(2) & strCase(3) & strCase(4)) > 0 Then
                        '     strCP16 = "": strCP17 = "": strCP18 = ""
                        'End If
                        'end 2020/03/27
                        Call OnUpdateFee("901", strCP20, strCP16, strCP17, strCP18)
                        'end 2022/05/03
                        'Added by Lydia 2018/05/10 標準價和底價
                        If ClsPDGetCaseLowPrice(strCase(1), strCase(6), "901", dCP33, dCP34) = 1 Then
                        End If

                        'Modified by Lydia 2019/09/19 +人工判斷
                        'If strCP901 = "" Then
                        'Mark by Lydia 2019/09/23 保留
                        'If strCP901 = "" And bolAdd901 = True Then
                        If strCP901 = "" Then
                            strPK = AutoNo("B", 6)
                            'Modified by Lydia 2025/06/25
                            'Pub_SetPAIsCase strCase(1), "901", strCP26 'Added by Lydia 2018/05/10 是否算案件數
                            If PUB_GetCPMbyCP10(strCase(1), "901", "cpm05") = "N" Then
                               strCP26 = "N"
                            End If
                            'end 2025/06/25
                            
                            'Add By Sindy 2021/6/18 非智慧局期限，要掛本所期限
                            'Remove by Lydia 2021/08/23 改用GetCP48取得strCP06
                            'Call GetPrjState6HM(strCase(1), "901", "cpm34", strExc(0))
                            'strExc(6) = "" '本所期限
                            'If Val(strCP48) > 0 And strExc(0) = "N" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                            '   strExc(6) = PUB_GetFCPOurDeadline(DBDATE(strCP48), , , , "N")
                            'End If
                            ''2021/6/18 END
                            
                            'Modified by Lydia 2018/04/25 +已分案CP122=Y
                            'Modified by Lydia 2018/05/10 +cp26,cp33,cp34
                            'Modified by Lydia 2018/06/26 告代預設為不請款
                            'strExc(0) = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp16,cp17,cp18,cp79,cp48,cp64,CP122,cp26,cp33,cp34) " & _
                                     "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','901',cp11,cp12,cp13," & CNULL(m_TCT10) & "," & CNULL(IIf(Chk2(48).Value = vbChecked, "N", "")) & _
                                     ", " & CNULL(strCP16, True) & ", " & CNULL(strCP17, True) & ", " & CNULL(strCP18, True) & ", " & CNULL(strCP16, True) & ", " & CNULL(strCP48, True) & _
                                     ", '" & ChangeWStringToWDateString(strSrvDate(1)) & IIf(Opt4s(0).Value = True, " 命名-提申後告代;", IIf(Opt4s(1).Value = True, " 命名-提申前告代;", " 命名-當日告代;")) & "', 'Y' " & _
                                     ", " & CNULL(strCP26) & ", " & dCP33 & ", " & dCP34 & " from caseprogress where cp09='" & m_TCT01 & "' "
                            'Modified by Lydia 2019/01/29 cp05收文日改為系統日(ex.FCP-60080的B類主動修正和告代之承辦期限:因為新案收文在107/12/11，但是ORI在108/1/19才來，所以1/21命名作業產生B類收文的期限為107/12/11+5或15個工作天；目前決定類似案例採人工修改期限。)
                            'Modify By Sindy 2021/6/18 + ,cp06
                            'Modified by Lydia 2021/08/23 strexc(6)=>strCP06 (GetCP48取得)
                            'Modified by Lydia 2022/05/03 CP20改用模組 PUB_GetCP20(strCase(1), "901")=> strCP20
                            strExc(0) = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp16,cp17,cp18,cp79,cp48,cp64,CP122,cp26,cp33,cp34,cp06) " & _
                                     "select cp01,cp02,cp03,cp04," & strSrvDate(1) & " , '" & strPK & "','901',cp11,cp12,cp13," & CNULL(m_TCT10) & "," & CNULL(strCP20) & _
                                     ", " & CNULL(strCP16, True) & ", " & CNULL(strCP17, True) & ", " & CNULL(strCP18, True) & ", " & CNULL(strCP16, True) & ", " & CNULL(strCP48, True) & _
                                     ", '" & ChangeWStringToWDateString(strSrvDate(1)) & IIf(Opt4s(0).Value = True, " 命名-提申後告代;", IIf(Opt4s(1).Value = True, " 命名-提申前告代;", " 命名-當日告代;")) & "', 'Y' " & _
                                     ", " & CNULL(strCP26) & ", " & dCP33 & ", " & dCP34 & ", " & CNULL(strCP06, True) & " from caseprogress where cp09='" & m_TCT01 & "' "
                            cnnConnection.Execute strExc(0), intI
                        
                        'Modified by Lydia 2019/09/19 +判斷有未發文的收文才修改
                        ElseIf strCP901 <> "" Then
                            strExc(1) = ""
                            If Opt4s(0).Value = True Then
                                strExc(2) = "1"
                                'Modified by Lydia 2021/11/04 本所期限cp06也要更新
                                strExc(1) = ", cp06 = " & CNULL(strCP06, True) & ", cp48 = " & CNULL(strCP48, True) & ", cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 命名-提申後告代;'||cp64 "
                            ElseIf Opt4s(1).Value = True Then
                                strExc(2) = "2"
                                'Modified by Lydia 2021/11/04 本所期限cp06也要更新
                                strExc(1) = ", cp06 = " & CNULL(strCP06, True) & ", cp48 = " & CNULL(strCP48, True) & ", cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 命名-提申前告代;'||cp64 "
                            ElseIf Opt4s(2).Value = True Then
                                strExc(2) = "3"
                                'Modified by Lydia 2022/04/28 + ,cp06 = CP66
                                strExc(1) = ",cp06 = CP66 , cp48 = cp66 , cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 命名-當日告代;'||cp64 "
                            End If
                            'Added by Lydia 2023/05/10 FCP案急件新案重新認領的告代901和主動修正203
                            If strCase(1) = "FCP" And Val(strCase(9)) > 0 And strExc(4) <> m_TCT10 And strExc(4) <> "" And InStr(strExc(1), "提申後告代") > 0 Then
                                 If InStr(strMailCp09 & ",", strCP901) = 0 Then
                                     strMailCp09 = strMailCp09 & "," & strCP901
                                     If InStr(strMailOld, strExc(4)) = 0 Then
                                        strMailOld = strMailOld & "';" & strExc(4)
                                     End If
                                     strMailSub = strMailSub & "、提申後告代"
                                     strMailCont = strMailCont & "  提申後告代　　　　承辦期限：" & ChangeWStringToTDateString(strCP48) & "　本所期限：" & ChangeWStringToTDateString(strCP06) & vbCrLf
                                 End If
                            End If
                            'end 2023/05/10
                            
                            'Modified by Lydia 2019/09/24 重送命名流程,預設改備註
                            'If m_TCT20 = strExc(2) Then strExc(1) = "" '沒有變更,不改變備註
                            'Modified by Lydia 2023/11/30 判斷沒有改期限才清限;ex.FCP-70655分別在11/14,11/30跑命名
                            'If m_TCT117 = strExc(2) And m_TCT14 = "" Then strExc(1) = ""
                            If m_TCT117 = strExc(2) And m_TCT14 = "" And strCP06Old = strCP06 And strCP48Old = strCP48 Then strExc(1) = ""
                            
                            'Modified by Lydia 2018/04/25 +已分案CP122=Y
                            strExc(0) = "update caseprogress set cp14='" & m_TCT10 & "' " & _
                                              ", cp16 = " & CNULL(strCP16, True) & ", cp17 = " & CNULL(strCP17, True) & " ,cp18 = " & CNULL(strCP18, True) & _
                                              strExc(1) & ", CP122='Y'  where cp09='" & strCP901 & "'  and cp158=0 "
                            'Modified by Lydia 2019/09/24 +Update
                            'cnnConnection.Execute strExc(0), intI
                            cnnConnection.Execute "begin user_data.user_enabled:=1; " & strExc(0) & "; end ;", intI
                        End If
                  '刪除收文
                  ElseIf strCP901 <> "" Then
                        strExc(0) = "insert into DataDeleteRecord(dd01,dd02,dd03,dd04,dd14,dd15,dd16,dd17,dd18,dd19,dd20,dd21,dd22,dd23,dd24,dd25,dd26,dd27,dd28) " & _
                                          "select cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,cp05,cp13,cp16,cp17,cp60,'" & strUserNum & "','工程師命名-刪除收文',cp66,cp65,'" & strSrvDate(1) & "',mno " & _
                                          "from caseprogress,(SELECT '" & strCP901 & "' fno,(MAX(DD28)+1) mno FROM DATADELETERECORD) x " & _
                                          "where cp09='" & strCP901 & "' and cp09=fno "
                        cnnConnection.Execute strExc(0), intI
                        
                        strExc(0) = "delete from caseprogress where cp09='" & strCP901 & "' and cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' "
                        'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                        'Pub_SeekTbLog strExc(0)
                        Pub_SeekTbLog strExc(0), , , , Me.Caption & "(" & Me.Name & ")"
                        cnnConnection.Execute strExc(0), intI
                  End If
            End If
       End If
       'end 2018/04/19
       
        'Added by Lydia 2023/05/10 因FCP新案急件重新認領，修改進度檔若有提申後告代、主動修正再發一次mail通知舊和新承辦人之事。
                                                 'Memo by Lydia 2023/05/10 如果主旨或內文有變，請一併查看frm060105_2的email是否要一致
        If strMailOld <> "" Then
           strMailOld = Mid(strMailOld, 2)
           strMailSub = Mid(strMailSub, 2)
           strMailSub = "【1.請分案 2.進行" & strMailSub & "】Our Ref: " & strCase(1) & "-" & strCase(2) & " [INCOM." & IIf(InStr(strMailSub, "告代") > 0, "901", "203") & "]"
           strMailSub = PUB_GetSetMailSubF2(m_PA75) & strMailSub
           strMailCont = "本案已提申且重新命名完畢 , 承辦工程師有修改, 請新承辦工程師處理後續" & vbCrLf & _
                               "新工程師：" & lblData(8).Caption & "　　　　　(原工程師：" & PUB_ReadUserData(strMailOld) & ")" & vbCrLf & vbCrLf & _
                               "1.主管請分案" & vbCrLf & _
                               "2.工程師請進行以下事項:" & vbCrLf & strMailCont
           strExc(1) = PUB_GetFCPEngSup(m_TCT10) & ";" & strMailOld & ";backup"  '新工程師之主管;舊工程師;backup
           strExc(0) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                " values( '" & strUserNum & "','" & m_TCT10 & "',to_char(sysdate,'yyyymmdd')" & _
                ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strMailSub) & "','" & ChgSQL(strMailCont) & "','" & strExc(1) & "')"
           cnnConnection.Execute strExc(0)
        End If
        'end 2023/05/10
        
        '提申後從命名系統修改專利名稱則有欄位註記
        'Modifeid by Lydia 2019/09/18 + bUpdPA : FCP-61274提申日6/10，之後於7/3修改名稱，未能加註
        'If Val(strCase(9)) > 0 And (m_PA05 <> txtData(3) Or m_PA06 <> txtData(4)) Then
        '     strExc(0) = "update transcasetitle set tct15='Y' where tct01='" & m_TCT01 & "' and tct15 is null "
        If Val(strCase(9)) > 0 And (m_PA05 <> txtData(3) Or m_PA06 <> txtData(4) Or bUpdPA = True) Then
             strExc(0) = "update transcasetitle set tct15='Y' where tct01='" & m_TCT01 & "' "
        'end 2019/09/18
             cnnConnection.Execute strExc(0), intI
        End If
       'end 2018/04/19
        'Added by Lydia 2018/10/18 黑白圖提申記錄在新案進度的備註
        If bUpdCP64 = True Then
             strExc(0) = "Update caseprogress set cp64=" & CNULL(ChangeTStringToTDateString(strSrvDate(1)) & " 命名-" & IIf(Chk2(49).Value = vbChecked, "彩圖", "黑白圖") & "提申(" & strUserNum & ");") & "||cp64  where cp09=" & CNULL(m_TCT01)
             cnnConnection.Execute strExc(0), intI
        End If
        'end 2018/10/18
   
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
        'Added by Lydia 2018/09/27 若輸入相似度後,程序未執行新案建檔(ex.FCP-59234) ; 107/9/13以前在輸入核稿人會判斷tf07有傳票號碼會自動刪除記錄
        ElseIf m_TF01t = "" And m_TF01 <> "" And InStr("201,209,235", m_TF01pty) > 0 And (txtData(6) <> "" Or txtData(7) <> "") Then
             'Modified by Lydia 2023/05/30 直接存入txtData(6)
             strExc(0) = "insert into transfee (TF01,TF20,TF19) values (" & CNULL(m_TF01) & ", " & CNULL(txtData(6)) & ", " & CNULL(txtData(7), True) & ") "
             cnnConnection.Execute strExc(0), intI
        'end 2018/09/27
        End If
        
        'Added by Lydia 2018/09/21 記錄欲翻譯人員(未發文) :在主管確認階段或卷宗區(工程師主管可修改)
        'Modified by Lydia 2018/11/26 主任直接在前一畫面的員工編號選擇工程師主管身份進入(ex.FCP-59961)
        'If Val(m_TF01cp27) = 0 And ((bDifference = True And iType = 3 And (InStr(strExSql, "TCT27") > 0 Or InStr(strExSql, "TCT28") > 0)) Or (iType = 1 And (strUserNum = m_TCT04 Or bolUpdMan = True))) Then
        'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
        'If Val(m_TF01cp27) = 0 And ((bDifference = True And iType = 3 And (InStr(strExSql, "TCT27") > 0 Or InStr(strExSql, "TCT28") > 0)) Or (iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True))) Then
        If Val(m_TF01cp27) = 0 And ((bDifference = True And iType = 3 And (InStr(strExSql, "TCT27") > 0 Or InStr(strExSql, "TCT28") > 0)) _
              Or (iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True))) Then
                  '-----------先刪除
                  'Mark by Lydia 2025/10/14 調整刪除條件;有新增記錄才要刪除，避免刪到從認翻譯作業產生的記錄
                  'strExc(0) = "delete from transfeeassign where tfa01='" & m_TF01 & "' "
                  'cnnConnection.Execute strExc(0), intI
                  'end 2025/10/14
                  If Trim(m_TCTList(11).fiNewData) <> "" Or Trim(m_TCTList(12).fiNewData) <> "" Then
                      strExc(1) = "": strExc(2) = ""
                      'Modified by Lydia 2025/03/13 新增國外翻譯社
                      'If Val(m_TCTList(11).fiNewData) > 0 And m_TCTList(11).fiNewData <> "4" Then '國外翻譯社
                      If Val(m_TCTList(11).fiNewData) > 0 Then
                           strExc(1) = Pub_GetTct27ID(m_TCT10, m_TCTList(11).fiNewData, m_TCTList(12).fiNewData)
                      'Modified by Lydia 2025/03/13 新增國外翻譯社
                      'ElseIf m_TCTList(11).fiNewData = "4" Then    '其他
                      ElseIf m_TCTList(11).fiNewData = "Z" Then
                           strExc(1) = UCase(m_TCTList(12).fiNewData)
                           If InStr("A,B", Right(strExc(1), 1)) > 0 Then
                                'Modified by Lydia 2019/05/31 含外翻人員
                                'strExc(1) = GetPrjSalesNM_2(Mid(m_TCTList(12).fiNewData, 1, Len(m_TCTList(12).fiNewData) - 1))
                                strExc(1) = GetPrjSalesNM_2(Mid(m_TCTList(12).fiNewData, 1, Len(m_TCTList(12).fiNewData) - 1), , , , True)
                                strExc(2) = Right(UCase(m_TCTList(12).fiNewData), 1)
                           End If
                      Else
                          strExc(1) = m_TCT10
                          strExc(2) = m_TCTList(11).fiNewData
                      End If
                      'Added by Lydia 2025/10/14 調整刪除條件;10/7發生FCP-074509的94012的認領記錄不存在，依順序是命名作業主任確認後直接到認翻譯作業認領，之後才有主管確認；
                      strExc(0) = "delete from transfeeassign where tfa01='" & m_TF01 & "' and tfa04='" & strExc(1) & "' "
                      cnnConnection.Execute strExc(0), intI
                      'end 2025/10/14
                      'Memo by Lydia 2025/10/14 主管確認產生的認翻譯人員記錄，包含已確認-認翻譯人員TFA07,TFA08
                      strExc(0) = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05,tfa06,tfa07,tfa08) select " & _
                                        "'" & m_TF01 & "', to_char(sysdate, 'YYYYMMDD') , to_char(sysdate, 'HH24MISS'), '" & strExc(1) & "','" & strExc(2) & "' " & _
                                        ",'" & strUserNum & "', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24MISS') from dual "
                      cnnConnection.Execute strExc(0), intI
                      'Added by Lydia 2025/10/15 選擇:命名作業指定翻譯人員，刪除認領翻譯人員，同時發送email通知認領人員
                      If m_strTFAcon <> "" Then
                         strExc(0) = "select st01,st02,decode(tfa05,'A','下班','上班') as tfa05 from transfeeassign,staff where tfa01='" & m_TF01 & "' and instr('" & m_strTFAcon & "',tfa04) > 0 and tfa06 is null and tfa04=st01(+)"
                         intI = 1
                         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                         If intI = 1 Then
                            strExc(4) = GetTCT27val
                            Do While Not RsTemp.EOF
                               strExc(0) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & "欲認領新案翻譯人員(" & RsTemp.Fields("st02") & "-" & RsTemp.Fields("tfa05")
                               strExc(3) = "主管命名作業已確認，翻譯人員為命名作業指定翻譯人員：" & strExc(4) & "，故已刪除本案認領翻譯。"
                               strSql = " insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                       " values( '" & strUserNum & "','" & m_strTFAcon & "',to_char(sysdate,'yyyymmdd')" & _
                                       ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(0)) & "','" & ChgSQL(strExc(3)) & "',null)"
                               cnnConnection.Execute strSql
                               strExc(0) = "delete from transfeeassign where tfa01='" & m_TF01 & "' and tfa04='" & RsTemp.Fields("st01") & "' "
                               cnnConnection.Execute strExc(0), intI
                               Sleep 100
                               RsTemp.MoveNext
                            Loop
                         End If
                      End If
                      'end 2025/10/15
                  End If
        End If
        'end 2018/09/21
        
        'Added by Lydia 2018/09/20 若已分案又改翻譯人員通知Sharon
        If bDifference = True And (InStr(strExSql, "TCT27") > 0 Or InStr(strExSql, "TCT28") > 0) And (m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) = 0) Then
             strExc(0) = Pub_GetSpecMan("M")
             strExc(1) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & " 修改認領翻譯人員"
             strExc(2) = Pub_GetTct27ID(m_TCT10, m_TCTList(11).fiOldData, m_TCTList(12).fiOldData, , strExc(3))
             strExc(2) = Pub_GetTct27ID(m_TCT10, m_TCTList(11).fiNewData, m_TCTList(12).fiNewData, , strExc(4))
             strExc(2) = "原認領人員：" & strExc(3) & vbCrLf & _
                              "新認領人員：" & strExc(4)
             If strExc(0) <> "" Then
                 strExc(3) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                                  " values( '" & strUserNum & "' , '" & strExc(0) & "' , '" & strSrvDate(1) & "' , '" & Format(ServerTime, "000000") & "' , '" & strExc(1) & "' , '" & strExc(2) & "')"
                 cnnConnection.Execute strExc(3)
             End If
        End If
        'end 2018/09/20
       
       'Move by Lydia 2018/07/13 從上面移到下面
       'Modified by Lydia 2018/07/18 需分別判斷主管(代主管)確認 ; ex.FCP-59439在主任確認(8/16)就產生告代，工程師送程序發文在主管確認(同在8/17)之前，當主管確認時檢查無未發文之告代則又產生B類收文告代。
       'If bDifference = True And (InStr(strExSql, "TCT118") > 0 Or InStr(strExSql, "TCT21") > 0) Then
       'Modified by Lydia 2018/11/26 主任直接在前一畫面的員工編號選擇工程師主管身份進入
       'If (iType = 1 And (strUserNum = m_TCT04 Or bolUpdMan = True)) Or (iType = 3 And bDifference = True And (InStr(strExSql, "TCT118") > 0 Or InStr(strExSql, "TCT21") > 0)) Then
       'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
       'If (iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) Or (iType = 3 And bDifference = True And (InStr(strExSql, "TCT118") > 0 Or InStr(strExSql, "TCT21") > 0)) Then
       If (iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True)) _
             Or (iType = 3 And bDifference = True And (InStr(strExSql, "TCT118") > 0 Or InStr(strExSql, "TCT21") > 0)) Then
            'Added by Lydia 2018/04/20 一案兩請和彩圖提申通知
            strTo = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4)) '程序
            strExc(1) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "")
            strExc(9) = Format(ServerTime, "000000")
            '一案兩請：發mail通知程序和承辦
            'Modified by Lydia 2018/07/13
            'If Chk2(0).Value = vbChecked Then
            'Modified by Lydia 2018/07/18 需分別判斷主管(代主管)確認
            'If Chk2(0).Value = vbChecked And InStr(strExSql, "TCT21") > 0 Then
            'Modified by Lydia 2019/12/02 重跑命名流程: 要記錄一案兩請和彩圖提申,因為承辦會去詢問代理人
            'If Chk2(0).Value = vbChecked And (iType = 1 Or (iType = 3 And bDifference = True And InStr(strExSql, "TCT21") > 0)) Then
            '第一次主管確認(代主管)直接通知, 之後重跑命名or主管自行修改要分析是否有變更
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If Chk2(0).Value = vbChecked And ( _
                (m_TCT14 = "" And iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) _
                Or (m_TCT14 <> "" And iType = 1 And (InStr(m_TCT14, "一案兩請") > 0 Or bUpdTCT21 = True)) _
                Or (iType = 3 And m_UserNo = m_TCT04 And bDifference = True And InStr(strExSql, "TCT21") > 0)) Then
            'end 2019/12/02
            If Chk2(0).Value = vbChecked And ( _
                (m_TCT14 = "" And iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True)) _
                Or (m_TCT14 <> "" And iType = 1 And (InStr(m_TCT14, "一案兩請") > 0 Or bUpdTCT21 = True)) _
                Or (iType = 3 And InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 And bDifference = True And InStr(strExSql, "TCT21") > 0)) Then
            'end 2022/10/12
                strExc(2) = PUB_GetFCPSalesNo(strCase(1), strCase(2), strCase(3), strCase(4)) '承辦
                strExc(3) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                                 " values( '" & strUserNum & "' , '" & strTo & ";" & strExc(2) & "' , '" & strSrvDate(1) & "' , '" & strExc(9) & "' , '" & strExc(1) & " 可一案兩請，請承辦人員與客戶聯絡！" & "' , '同主旨')"
                cnnConnection.Execute strExc(3)
            End If
            '彩圖提申：發mail通知程序和cc:主管
            'Modified by Lydia 2018/07/13
            'If Chk2(49).Value = vbChecked Then
            'Modified by Lydia 2018/07/18 需分別判斷主管(代主管)確認
            'If Chk2(49).Value = vbChecked And InStr(strExSql, "TCT118") > 0 Then
            'Modified by Lydia 2019/12/02 重跑命名流程: 要記錄一案兩請和彩圖提申,因為承辦會去詢問代理人
            'If Chk2(49).Value = vbChecked And (iType = 1 Or (iType = 3 And bDifference = True And InStr(strExSql, "TCT118") > 0)) Then
            '第一次主管確認(代主管)直接通知, 之後重跑命名or主管自行修改要分析是否有變更
            'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
            'If Chk2(49).Value = vbChecked And ( _
                (m_TCT14 = "" And iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) _
                Or (m_TCT14 <> "" And iType = 1 And (InStr(m_TCT14, "彩圖提申") > 0 Or bUpdTCT118 = True)) _
                Or (iType = 3 And m_UserNo = m_TCT04 And bDifference = True And InStr(strExSql, "TCT118") > 0)) Then
            'end 2019/12/02
            If Chk2(49).Value = vbChecked And ( _
                (m_TCT14 = "" And iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True)) _
                Or (m_TCT14 <> "" And iType = 1 And (InStr(m_TCT14, "彩圖提申") > 0 Or bUpdTCT118 = True)) _
                Or (iType = 3 And InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 And bDifference = True And InStr(strExSql, "TCT118") > 0)) Then
            'end 2022/10/12
                'Modified by Lydia 2024/10/07 debug: 工程師組別strCase(5)>>strCase(7)
                If (strSrvDate(1) < "20230501" Or (strSrvDate(1) >= "20230501" And strCase(7) = "3")) Then 'Added by Lydia 2023/04/26 (電子電機、化學、機械 三組)直接以彩圖製作成ori版本提申(自112年5月1日起實施); 日文組仍維持原程式=>需判斷是否以彩圖提申
                    strExc(9) = Format(Val(strExc(9)) + 1, "000000")
                    strExc(2) = ""
                    strExc(0) = "select nvl(st52,'N') from staff where st01='" & strTo & "' "
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                        If "" & RsTemp.Fields(0) <> "N" Then strExc(2) = "" & RsTemp.Fields(0)
                    End If
                    'Modified by Lydia 2018/07/13 改主旨提醒命名人員,程序改為cc
                    'strExc(3) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                     " values( '" & strUserNum & "' , '" & strTo & "' , '" & strSrvDate(1) & "' , '" & strExc(9) & "' , '" & strExc(1) & " 須彩圖提申，請製作彩圖提申本！" & "' , '同主旨' , '" & strExc(2) & "')"
                    'Modified by Lydia 2020/02/15 主旨+且檔名請命名為FCP0XXXXXX.FIX.ORI.pdf
                    'Modified by Lydia 2021/08/03 debug 命名為FCP0=>命名為FCP
                    strExc(3) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                     " values( '" & strUserNum & "' , '" & m_TCT10 & "' , '" & strSrvDate(1) & "' , '" & strExc(9) & "' , '" & strExc(1) & " 須彩圖提申，請確認是否已上傳彩圖提申本！且檔名請命名為FCP" & strCase(2) & ".FIX.ORI.pdf" & "' , '同主旨' , '" & strTo & IIf(strExc(2) <> "", ";" & strExc(2), "") & "')"
                    cnnConnection.Execute strExc(3)
                End If 'Added by Lydia 2023/04/26
            End If  'end 2018/04/20
        End If 'end 2018/07/13
       
       'Added by Lydia 2021/04/16 外專日文組新案無卷命名email設定：命名完成通知email(工程師主管確認)
       'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
       'If (iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) Or (iType = 3 And bDifference = True) And Val(n_CP27) = 0 Then
       'Modified by Lydia 2023/02/19 因為之前「通知命名修改的內容Email(FCP-xxxxx命名區資料異動，請主管至未命名區\待確認-重新確認)，請cc程序人員」，工程師最後主管確認之【命名完成】維持要通知程序人員
       'If (iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True)) Or (iType = 3 And bDifference = True) And Val(n_CP27) = 0 Then
       If (iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True)) Then
          'If strCase(7) = "3" Then 'Remove by Lydia 2021/05/13 配合防疫措施，英文組使用新案無卷命名email設定
               If PUB_GetTCTmail(True, 4, strCase(1), strCase(2), strCase(3), strCase(4), m_TCT01) Then
               End If
          'End If  'Remove by Lydia 2021/05/13 配合防疫措施，英文組使用新案無卷命名email設定
       End If
       'end 2021/04/16
       
       'Added by Lydia 2021/10/19 FMP案按確定時，有收文901,902,924,942,209,203,228收文性質，承辦人直接更新為命名人員。
       'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
       'If strCase(1) = "P" And (iType = 1 And (m_UserNo = m_TCT04 Or bolUpdMan = True)) And Val(n_CP27) = 0 Then
       If strCase(1) = "P" And (iType = 1 And (InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 Or bolUpdMan = True)) And Val(n_CP27) = 0 Then
           strExc(0) = "select cp09 from caseprogress where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' " & _
                             "and cp10 in ('901','902','924','942','209','203','228') and cp158=0 and cp159=0 and cp14<>'" & m_TCT10 & "' order by cp09 "
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  strTmp = "update caseprogress set cp14='" & m_TCT10 & "'  where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' and cp09='" & RsTemp.Fields("cp09") & "' "
                  Pub_SeekTbLog strTmp
                  cnnConnection.Execute strTmp
                  RsTemp.MoveNext
               Loop
           End If
       End If
       'end 2021/10/19
       
       'Added by Lydia 2021/04/22 工程師完成命名時能自動發email通知上級主管
       If iType = 1 Then
           strExc(0) = "": strExc(1) = "": strExc(2) = ""
           'Modified by Lydia 2022/10/12 系統特殊設定之工程師主管(配合特殊情況之指定職代，增加判斷)
           'If m_UserNo = m_TCT04 And Val(m_TCT05) = 0 Then '各組主管
           If InStr(m_TCT04 & "," & s_TCT04m, m_UserNo) > 0 And Val(m_TCT05) = 0 Then   '各組主管
                '代主任確認
                If bolUpdMan2 = True Then
                    strExc(0) = m_TCT07
                    strExc(1) = m_UserNo
                End If
                strExc(2) = "命名完成，可以進卷宗區查看命名記錄。"
           Else
                strExc(0) = m_TCT04
                '代主管確認
                If bolUpdMan = True Then
                     strExc(1) = m_UserNo
                     strExc(2) = "命名完成，可以進卷宗區查看命名記錄。"
                Else
                     strExc(2) = "命名人員已完成命名，請主管進行確認。"
                End If
           End If
           'Added by Lydia 2024/03/18 外專機械設計組人員異動調整程式：針對內專主管直接進行代主管確認
           If m_TCT07 = m_UserNo And Mid(m_TCT07, 4, 1) = "9" And m_UserNo = strUserNum Then
               strExc(0) = ""
           End If
           'end 2024/03/18
           If strExc(0) <> "" Then
                strExc(2) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & strExc(2)
                strExc(3) = Mid(strExc(2), 1, Len(strExc(2)) - 1) & vbCrLf
                strExc(3) = strExc(3) & "命名人員：" & m_TCT10 & " " & lblData(8).Caption
                If bolUpdMan = True Then
                     strExc(2) = strExc(2) & "【代主管確認】"
                End If
                If bolUpdMan2 = True Then
                     strExc(2) = strExc(2) & "【代確認】"
                End If
                strExc(0) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                    " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd')" & _
                    ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "', '" & strExc(1) & "'  )"
                cnnConnection.Execute strExc(0)
           End If
       End If
       'end 2021/04/22
       
    cnnConnection.CommitTrans
   
   OnSaveData = True
Err03:
   If Err.Number <> 0 Then
      MsgBox Err.Description
      cnnConnection.RollbackTrans
   End If
End Function

Private Sub Opt1_Click(Index As Integer)
    If opt1(Index).Value = True Then
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
                  cmdFile.Visible = True
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
                Cancel = True
            'end 2018/06/01
            Else
                'Added by Lydia 2018/09/21 判斷在職員工名稱
                If Trim(txtData(47).Text) <> Trim(m_TCTList(12).fiOldData) Then
                    'Modified by Lydia 2019/05/08 + 含外翻人員
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
   If Index <> 2 Then '因為主管不改翻譯人員，所以不鎖定
      txtData(Index).SetFocus
      Txtdata_GotFocus Index
   End If
   Cancel = True
End Sub

Private Function TxtValidate1() As Boolean
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
   
   If Frame1.Visible = True Then
      inB = 0
      For Each oOpt In opt1
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
   If strSrvDate(1) >= "20180813" And (m_iStiu = "M" Or cmdOK(3).Visible = True) And Chk2(1).Value = vbChecked And m_TF01 <> "" And m_TF01pty = "201" And Val(m_TF01cp27) = 0 _
             And (m_TF29 = "Y" Or m_TF01cp14 = "" Or (m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) > 0)) Then
       'Modified by Ldia 2019/04/15 案號可以6碼或5碼
       'strExc(1) = Dir(strResPath & "\" & strCase(1) & strCase(2) & "*.res.doc*")
       ''Added by Lydia 2018/09/19 開放PDF (ex.FCP-59599)
       'If strExc(1) = "" Then
       '     strExc(1) = Dir(strResPath & "\" & strCase(1) & strCase(2) & "*.res.pdf")
       'End If
       ''end 2018/09/19
       strExc(1) = Dir(strResPath & "\" & strCase(1) & "*" & Val(strCase(2)) & ".res.doc*")
       If strExc(1) = "" Then strExc(1) = Dir(strResPath & "\" & strCase(1) & "*" & Val(strCase(2)) & ".res.pdf")
       'ennd 2019/04/15
       
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
                   cmdFile.Visible = True
                   cmdFile.SetFocus
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
                            'Modified by Lydia 2018/09/19 開放PDF
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
                   MsgBox "下列檔案不符合命名規則(ex.FCP012345.RES.Doc 或 FCP012345.RES.PDF)：" & strExc(2), vbCritical
                   Exit Function
               End If
           End If
       End If
   End If
   'end 2018/07/12
   
   'Added by Lydia 2023/01/18 命名作業不可新增告代和主動修正
   If Frame5(8).Enabled = True And strNotBList <> "" And (Chk2(45).Value = vbChecked Or Chk2(47).Value = vbChecked) Then
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
      If PUB_GetITStoList(Me.Name, strExc(10), strExc(9), False, False, , , "Y01") = True Then 'Memo by Lydia 2025/10/02 考慮主管可能沒有修改勾選，預設都要產生Word提示
      End If
      If strExc(2) = "" Then 'Added by Lydia 2025/10/02
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
             'Remove by Lydia 2018/09/20 不限制,修改時發email通知
             'ElseIf m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) = 0 Then
             '    MsgBox "新案翻譯已分案，不可變更欲翻譯人員！", vbCritical
             '    Exit Function
             'end 2018/09/20
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
             End If
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
   'Modified by Lydia 2023/04/26 (電子電機、化學、機械 三組)直接以彩圖製作成ori版本提申(自112年5月1日起實施); 日文組仍維持原程式=>需判斷是否以彩圖提申
   'If m_PA63 = "Y" Then
   'Modified by Lydia 2024/10/07 debug: 工程師組別strCase(5)>>strCase(7)
   If m_PA63 = "Y" And (strSrvDate(1) < "20230501" Or (strSrvDate(1) >= "20230501" And strCase(7) = "3")) Then
       If Chk2(49).Value = vbChecked And m_TCT01cp64 <> "" And InStr(m_TCT01cp64, "命名-黑白圖提申") > 0 Then
            inB = MsgBox("本案客戶有提供彩圖，工程師在命名階段選擇黑白圖提申，請判斷是否以彩圖提申" & vbCrLf & "是：彩圖提申" & vbCrLf & "否：黑白圖提申" & vbCrLf & "取消：回到命名作業再重新判斷", vbInformation + vbYesNoCancel + vbDefaultButton3)
            If inB = 2 Then 'Cancel
                Exit Function
            ElseIf inB = 7 Then 'No
                 Chk2(49).Value = vbUnchecked
            Else
                 bUpdCP64 = True
            End If
       End If
       If Chk2(49).Value = vbUnchecked And (m_TCT01cp64 = "" Or (m_TCT01cp64 <> "" And InStr(m_TCT01cp64, "命名-黑白圖提申") = 0)) Then
            inB = MsgBox("本案客戶有提供彩圖，請判斷是否以彩圖提申" & vbCrLf & "是：彩圖提申" & vbCrLf & "否：黑白圖提申" & vbCrLf & "取消：回到命名作業再重新判斷", vbInformation + vbYesNoCancel + vbDefaultButton3)
            If inB = 2 Then 'Cancel
                Exit Function
            ElseIf inB = 7 Then 'No
                 bUpdCP64 = True
            Else
                 Chk2(49).Value = vbChecked
                 bUpdCP64 = True
            End If
       End If
   End If
   'end 2018/10/18
   
   'Added by Lydia 2020/02/17 檢查「名稱有特殊字」
   If (strCase(1) = "P" Or strCase(1) = "FCP") And (m_iStiu = "M" Or cmdOK(3).Visible = True) Then
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
   'end 2020/02/17
   
   'Added by Lydia 2023/03/10 FMP大陸新藥發明專利權期限補償控管
   If Chk2(51).Visible = True And ((m_iStiu = "M" And cmdOK(0).Enabled = True And cmdOK(1).Enabled = True) Or cmdOK(3).Visible = True) Then
      If Chk2(51).Value = 1 And Opt5(0).Value = False Then
          MsgBox "專利權期間延長相關的案件類別請選擇" & Opt5(0).Caption, vbExclamation, "大陸新藥發明專利權期限補償控管"
          SSTab1.Tab = 0
          Exit Function
      End If
   End If
   'end 2023/03/10
   
   'Added by Lydia 2024/10/17 英文組的彩圖依照PA63設定--from Phoebe ; ex.FCP-72571 'Added by Lydia 2023/04/26 (電子電機、化學、機械 三組)直接以彩圖製作成ori版本提申(自112年5月1日起實施); 日文組仍維持原程式=>需判斷是否以彩圖提申
   If strSrvDate(1) >= "20230501" And m_iStiu = "M" And strCase(7) <> "3" And m_PA63 = "Y" And Chk2(49).Value = vbUnchecked Then
       MsgBox "承辦人員於接洽單註記有彩圖提申！", vbCritical + vbOKOnly
       SSTab1.Tab = 0
       Exit Function
   End If
   'end 2023/04/26
    
   'Added by Lydia 2021/09/27 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/09/27
   
   '清除畫面所有欄位的跳行符號
   PUB_FilterFormText Me
   
   TxtValidate1 = True
End Function

'Modified by Lydia 2019/09/02 +輸出PDF檔檔名(pFileName)
Public Sub PUB_PrintTCTcon(ByVal iTCT01 As String, ByVal iPrtName As String, Optional ByVal strOldPrinter As String = "", Optional ByVal pFileName As String)
Dim rsPD As New ADODB.Recordset
Dim rsPrt As New ADODB.Recordset
Dim inA As Integer
Dim oOrt As Integer
Dim strTitle(0 To 15) As String '列印抬頭
Dim tmpArr As Variant
Dim m_Start01 As Integer
Dim strTempB As String '圖式-其他問題
Dim m_PA08 As String '專利種類

    If iTCT01 = "" Or iPrtName = "" Then Exit Sub
    
    strSql = "select a.*,b.cp01,b.cp02,b.cp03,b.cp04,sqldatet(b.cp05) cp05,sqldatet(b.cp06) cp06,sqldatet(b.cp07) cp07,b.cp10," & _
             "s1.st02 tct04n,s2.st02 tct07n,s3.st02 tct10n " & _
             "from TransCaseTitle a ,caseprogress b, staff s1, staff s2, staff s3 " & _
             "where tct01='" & iTCT01 & "' and tct01=cp09(+) and cp10 in (" & NewCasePtyList & ") " & _
             "and tct04=s1.st01(+) and tct07=s2.st01(+) and tct10=s3.st01(+) "
    inA = 1
    Set rsPrt = ClsLawReadRstMsg(inA, strSql)
    If inA = 0 Then
        Exit Sub
    Else
        '法限
        strTitle(0) = "法定期限：" & rsPrt.Fields("cp07")
        '收文日期
        strTitle(10) = "收文日期：" & rsPrt.Fields("cp05")
        '所限
        strTitle(3) = "本所期限：" & rsPrt.Fields("cp06")
        strTitle(1) = "譯畢期限：" '有期限才顯示,併成一行
        If "" & rsPrt.Fields("tct02") <> "" Then
            strTitle(1) = strTitle(1) & "急件，請於" & ChangeTStringToTDateString(TransDate(rsPrt.Fields("TCT02"), 1)) & " " & Format("" & rsPrt.Fields("TCT03"), "00:00") & "前譯畢名稱"
        End If
        strTitle(2) = ""
        
        'Modified by Lydia 2018/03/06 改成命名人員
        'strTitle(5) = "" '"列印日期" & ChangeTStringToTDateString(strSrvDate(2)) =>空下來
        strTitle(5) = "命名人員：" & "" & rsPrt.Fields("tct10n")
        
        '本所案號
        strTitle(13) = "本所案號：" & rsPrt.Fields("cp01") & "-" & rsPrt.Fields("cp02") & "-" & rsPrt.Fields("cp03") & "-" & rsPrt.Fields("cp04")
        
        '抓基本資料
        strSql = "select pa150,DECODE(pa150,'1','" & PUB_GetFCPGrpName("1") & "','2','" & PUB_GetFCPGrpName("2") & "','3','" & PUB_GetFCPGrpName("3") & "','4','" & PUB_GetFCPGrpName("4") & "',pa150) grpname, " & _
                     "pa08,pa09,pa26,pa75,cu10,n1.na03 cna03,fa10,n2.na03 fna03 " & _
                    ", pa05,pa06,pa07,pa158,nvl(fa104,'N') fa104,nvl(cu174,'N') cu174 " & _
                    "from patent,fagent,customer,nation n1,nation n2 " & _
                    "where pa01='" & rsPrt.Fields("cp01") & "' and pa02='" & rsPrt.Fields("cp02") & "'  and pa03='" & rsPrt.Fields("cp03") & "' and pa04='" & rsPrt.Fields("cp04") & "' " & _
                    "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and cu10=n1.na01(+) " & _
                    "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=n2.na01(+) "
        inA = 1
        Set rsPD = ClsLawReadRstMsg(inA, strSql)
        If inA = 1 Then
           'Mark by Lydia 2018/04/18
           'strTitle(7) = "總收文號："
           strTitle(12) = "2.中說類型："
           strTitle(15) = "" & rsPD.Fields("pa07") '案件名稱(日)
           '專利種類
           strTitle(9) = "1.專利種類：" & PUB_GetPatentKindName("" & rsPD.Fields("pa08"), "" & rsPD.Fields("pa09"))
           m_PA08 = "" & rsPD.Fields("pa08")
           strTitle(14) = "分案組別：" & rsPD.Fields("grpname")
           '國籍
           If Trim("" & rsPD.Fields("fna03") <> "") Then
               strTitle(11) = "代理人國籍：" & rsPD.Fields("fna03")
           Else
               strTitle(11) = "申請人國籍：" & rsPD.Fields("cna03")
           End If
           
           'FCP是否電子送件,外專發文時做檢查用
           strExc(1) = ""
           If "" & rsPD.Fields("fa104") = "Y" Then
              strExc(1) = "代理人"
           End If
           If "" & rsPD.Fields("cu174") = "Y" Then
              strExc(1) = strExc(1) & IIf(strExc(1) <> "", "、", "") & "客戶"
           End If
           '電子送件
           'strTitle(2) = IIf(strExc(1) <> "", "電子送件", "") '換位置
           strTitle(4) = IIf(strExc(1) <> "", "電子送件", "")

           'Modified by Lydia 2018/04/18 +實審416,回代902,主動修正203, 限制A類收文
           'Modified by Lydia 2018/05/10 +FMP案414,938,939,106,228
           strSql = "select 2 as ord1 ,sqldatet(cp05) cp05,sqldatet(cp06) cp06,sqldatet(cp07)cp07,cp09,cp10,cpm03,cp13,st02 " & _
                       "from caseprogress,staff,casepropertymap " & _
                       "where cp01='" & rsPrt.Fields("cp01") & "' and cp02='" & rsPrt.Fields("cp02") & "'  and cp03='" & rsPrt.Fields("cp03") & "' and cp04='" & rsPrt.Fields("cp04") & "' " & _
                       "and cp10 in (" & GetAddStr(FcpTctPtys & ",416,924,902,203,414,938,939,106,228") & ") and substr(cp09,1,1)='A' and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) "
           strSql = strSql & "union all select 1 as ord1 ,sqldatet(cp05) cp05,sqldatet(cp06) cp06,sqldatet(cp07)cp07,cp09,cp10,cpm03,cp13,st02 " & _
                       "from caseprogress,staff,casepropertymap " & _
                       "where cp01='" & rsPrt.Fields("cp01") & "' and cp02='" & rsPrt.Fields("cp02") & "'  and cp03='" & rsPrt.Fields("cp03") & "' and cp04='" & rsPrt.Fields("cp04") & "' " & _
                       "and cp10 in (" & GetAddStr(NewCasePtyList) & ") and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) "
           strSql = strSql & "order by ord1,cp09 "
           inA = 1
           Set RsTemp = ClsLawReadRstMsg(inA, strSql)
           If inA = 1 Then
              With RsTemp
                  .MoveFirst
                  Do While Not .EOF
                     If Val("" & RsTemp.Fields("ord1")) = 2 Then '中說進度
                        '總收文號
                        'Modified by Lydia 2018/04/18 第一道收文號顯示9碼,後面顯示3碼
                        'strTitle(7) = strTitle(7) & "," & Right("" & RsTemp.Fields("cp09"), 6)
                        strTitle(7) = strTitle(7) & "," & Right("" & RsTemp.Fields("cp09"), 3)
                        If InStr(FcpTctPtys, "" & RsTemp.Fields("cp10")) > 0 Then
                            '中說類型
                            If "" & RsTemp.Fields("cp10") = "242" Then
                               '製作中說210＆外文提申本242 是一起產生
                               strTitle(12) = strTitle(12) & "＆外文提申本"
                            Else
                               strTitle(12) = strTitle(12) & RsTemp.Fields("cpm03")
                            End If
                            '智權人員
                            strTitle(8) = "智權人員：" & RsTemp.Fields("st02")
                        End If
                     Else
                        '總收文號
                        'Modified by Lydia 2018/04/18 第一道收文號顯示9碼
                        'strTitle(7) = strTitle(7) & Right("" & RsTemp.Fields("cp09"), 6)
                        strTitle(7) = strTitle(7) & "" & RsTemp.Fields("cp09")
                        '新申請案-案件性質
                        strTitle(6) = "案件性質：" & RsTemp.Fields("cpm03")
                     End If
                     .MoveNext
                  Loop
              End With
           End If
        End If
        'Added by Lydia 2018/04/18 因為收文號太長,只顯示第1個和最後一個收文號
        strTitle(7) = "總收文號：" & Mid(strTitle(7), 1, 9) & "~" & Right(strTitle(7), 3)
        
RePrint:
        '設定印表機
        If iPrtName <> "" Then
           PUB_RestorePrinter iPrtName
        End If
        oOrt = Printer.Orientation
        
        'Added by Lydia 2019/09/02 輸出PDF檔
        If pFileName <> "" Then
             frmPDF.Show
             frmPDF.StartProcess App.path, pFileName & ".PDF"
        End If
        
        '開始列印
        iPage = 1
        Printer.PaperSize = 9 '設定紙張 A4
        Printer.Orientation = 1 '直印
        lngLineHeight = 300
        lngPageHeight = Printer.ScaleHeight
        lngPageWidth = Printer.ScaleWidth
        m_TBWidth = Printer.ScaleWidth - 700
        
        '列印抬頭
        Call PrintStaticData(strTitle)
        
        iPrint = iPrint - lngLineHeight / 2 '直接跳一行有點多
        'Memo by Lydia 2019/09/17 修改時,請一併修改ProcDataWord
        For inA = m_FS To TF_TCT
            'Modified by Lydia 2018/04/18 +117,118
            'Modified by Lydia 2021/04/09 +119 有序列表
            'Modified by Lydia 2023/03/10 +120 專利權期間延長相關
            If InStr(TF_TCTnotFS & ",116,117,118,119,120", Format(inA, "000")) = 0 Then 'Added by Lydia 2018/03/01 加判斷
                strExc(1) = "" & rsPrt.Fields(inA - 1)
                If inA = 27 And (strExc(1) = "A" Or strExc(1) = "B") Then    '欲翻譯此案件者
                   strExc(1) = "" & rsPrt.Fields("TCT10") & "-" & strExc(1)
                End If
                If GetDataDef("W", inA, strExc(1)) = True Then
                End If
                'Added by Lydia 2018/03/01 主動修正+是否不請款
                If inA = 19 Then
                     'Added by Lydia 2018/04/18 +提申後/提申前
                     If GetDataDef("W", 117, "" & rsPrt.Fields(117 - 1)) = True Then
                     End If
                     '不請款
                     If GetDataDef("W", 116, "" & rsPrt.Fields(116 - 1)) = True Then
                     End If
                'Added by Lydia 2018/04/18 彩圖提申(118)和一案兩請同一行
                ElseIf inA = 21 Then
                     If GetDataDef("W", 118, "" & rsPrt.Fields(118 - 1)) = True Then
                     End If
                     'Added by Lydia 2021/04/09 有序列表
                     If GetDataDef("W", 119, "" & rsPrt.Fields(119 - 1)) = True Then
                     End If
                     'end 2021/04/09
                     'Added by Lydia 2023/03/10 專利權期間延長相關
                     If "" & rsPrt.Fields("TCT120") <> "" Then
                        If GetDataDef("W", 120, "" & rsPrt.Fields(120 - 1)) = True Then
                        End If
                     End If
                     'end 2023/03/10
                'end 2018/04/18
                End If
                'end 2018/03/01
                
                If bolChgText = True Then
                    '案件名稱
                    If inA = 16 Or inA = 17 Then
                       PrintDetail strTempA, True
                       If inA = 17 Then  '案件名稱(日)
                          PrintDetail String(4, "　") & "(日)：" & strTitle(15), True
                       End If
                    '主動修正和告代：先合併為一個字串，後逐行列印
                    ElseIf inA = 20 Then
                       tmpArr = Empty
                       tmpArr = Split(strTempA, "|")
                       iPrint = m_Start01 '回到記錄Y座標
                       For intI = 0 To UBound(tmpArr)
                           If Trim(tmpArr(intI)) <> "" Then
                              PrintDetail tmpArr(intI), False, True
                           End If
                       Next intI
                    Else
                       If inA = 21 Then
                          iPrint = m_Start01 '回到記錄Y座標
                          '有設計案屬性，要跳行
                          If "" & rsPrt.Fields("TCT18") <> "" Then
                             PrintNewLine
                          End If
                       End If
                       
                       '----------不顯示
                       '標的不一致：若未勾選，則不顯示資料
                       If (inA = 84 Or inA = 86) And "" & rsPrt.Fields("TCT82") = "" Then
                          strTempA = ""
                       End If
                       '非設計案-不顯示設計案屬性
                       If inA = 18 And m_PA08 <> "3" Then
                          strTempA = ""
                       End If
                       '----------不顯示
                       
                       '圖式-其他問題放最後面列印
                       If inA = 105 Then
                          strTempB = strTempA
                          strTempA = ""
                       End If
                       PrintDetail strTempA
                    End If
                End If
            End If
            '畫分隔線
            If Format(inA, "000") = "017" Then
               PrintAreaLine True, True
               m_Start01 = iPrint '記錄Y座標
            ElseIf (Format(inA, "000") = "028" And "" & rsPrt.Fields("TCT29") <> "") Or _
                   (Format(inA, "000") = "067" And "" & rsPrt.Fields("TCT68") <> "") Or _
                   (Format(inA, "000") = "094" And "" & rsPrt.Fields("TCT95") <> "") Then
               PrintAreaLine
            End If
        Next inA
        
        '圖式-其他問題放最後面列印
        If strTempB <> "" Then
           PrintDetail strTempB
        End If
        Printer.EndDoc
        
        'Added by Lydia 2019/09/02 輸出PDF檔
        If pFileName <> "" Then
            frmPDF.EndtProcess
            Unload frmPDF
        End If
    End If
    
Set rsPD = Nothing
Set rsPrt = Nothing
ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub PrintDetail(ByVal nText As String, Optional bAll As Boolean = False, Optional ByVal bRight As Boolean = False)
Dim inP As Integer
    
    If Trim(nText) = "" Then Exit Sub  '無資料，不列印
    PrintNewLine
    If bRight = False Then
       SmartPrint nText, ciPrtX, iPrint, IIf(bAll = False, 146, 175), lngLineHeight
    Else
       SmartPrint nText, m_MidCol + 150, iPrint, 40, lngLineHeight, False
    End If

End Sub

'超過長度，跳行列印
'p_dblWidth = 列印寬(mm)
'p_lngLineH = 跳行高(twips)
Private Sub SmartPrint(ByVal p_Data As String, ByRef p_lngX As Integer, ByRef p_lngY As Integer, Optional p_lngWidth As Long = 180, Optional p_lngLineH As Long = 300, Optional ByVal bolJumpLine As Boolean = True)
   Dim strData As String, strCache As String, i As Integer
      
   strData = p_Data
   For i = 1 To Len(p_Data)
      If Printer.TextWidth(strCache & Mid(strData, i, 1)) > p_lngWidth * 56.7 Then
         Printer.CurrentX = p_lngX
         Printer.CurrentY = p_lngY
         Printer.Print strCache
         strCache = Mid(strData, i, 1)
         If bolJumpLine = True Then PrintNewLine
      Else
         strCache = strCache & Mid(strData, i, 1)
      End If
   Next
   If strCache <> "" Then
      Printer.CurrentX = p_lngX
      Printer.CurrentY = p_lngY
      Printer.Print strCache
   End If
End Sub

Private Sub PrintStaticData(ByRef PTitle() As String)
Dim fLeft(0 To 2) As Integer
Dim inP As Integer
Dim inP2 As Integer

    fLeft(0) = ciStartX
    fLeft(1) = fLeft(0) + 150 + m_TBWidth / 3
    fLeft(2) = fLeft(1) + 250 + m_TBWidth / 3
    
    iPrint = 300
    '版本
    'Modified by Lydia 2018/04/18
    'strExc(0) = "2017-10"
    strExc(0) = "2018-04"
    Printer.Font.Name = "細明體"
    Printer.Font.Size = ciFontSize - 3
    Printer.CurrentX = 900
    Printer.CurrentY = iPrint + 200
    Printer.Print strExc(0)
    'Modified by Lydia 2020/03/31 事務所名稱在事務所合併日前以1抓公司檔,合併後以2抓
    'strExc(0) = "台一國際專利商標事務所外專案件命名記錄"
    If strSrvDate(1) >= 事務所合併日 Then
        strExc(0) = CompNameQuery("2")
    Else
        strExc(0) = CompNameQuery("1")
    End If
    strExc(0) = strExc(0) & "外專案件命名記錄"
    'end 2020/03/31
    
    Printer.Font.Name = "標楷體"
    Printer.Font.Size = ciTitleFontSize
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(0)) / 2)
    Printer.CurrentY = iPrint
    Printer.Print strExc(0)
    
    PrintNewLine
    Printer.Font.Name = "細明體"
    Printer.Font.Size = ciFontSize - 2
    For inP = 0 To 14
       inP2 = inP Mod 3
       If inP2 = 0 Then
          PrintNewLine
       End If
       '電子送件,加粗體
       If inP = 4 Then
           Printer.Font.Name = "標楷體"
           Printer.Font.Size = ciFontSize + 2
           Printer.Font.Bold = True
           Printer.CurrentX = fLeft(0 + inP2)
           Printer.CurrentY = iPrint
           Printer.Print PTitle(inP)
           Printer.Font.Name = "細明體"
           Printer.Font.Size = ciFontSize - 2
           Printer.Font.Bold = False
       Else
           Printer.CurrentX = fLeft(0 + inP2)
           Printer.CurrentY = iPrint
           Printer.Print PTitle(inP)
       End If
    Next
    iPrint = iPrint + 150
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(10800 / Printer.TextWidth("-"), "-")
    PrintNewLine
    
    Printer.Font.Size = ciFontSize
    Printer.Font.Bold = False
    PrintTableLine
End Sub

'換行判斷
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 4)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintTableLine '畫表格
   End If
End Sub

Private Sub PrintTableLine()
    Select Case iPage
        Case 1  '第一頁
            'Table-上、下邊界
            Printer.Line (ciStartX, iPrint)-(m_TBWidth, iPrint)
            Printer.Line (ciStartX, m_Tbottom)-(m_TBWidth, m_Tbottom)
            'Table-左、右邊界
            Printer.Line (ciStartX, iPrint)-(ciStartX, m_Tbottom)
            Printer.Line (m_TBWidth, iPrint)-(m_TBWidth, m_Tbottom)
        Case Else
            iPrint = ciPrtY
            'Table-上、下邊界
            Printer.Line (ciStartX, m_TTop)-(m_TBWidth, m_TTop)
            Printer.Line (ciStartX, m_Tbottom)-(m_TBWidth, m_Tbottom)
            'Table-左、右邊界
            Printer.Line (ciStartX, m_TTop)-(ciStartX, m_Tbottom)
            Printer.Line (m_TBWidth, m_TTop)-(m_TBWidth, m_Tbottom)
            '收文主動修正和告代欄位=>備註
            Printer.Line (m_MidCol, m_TTop)-(m_MidCol, m_Tbottom)
    End Select
End Sub

'畫分隔線
Private Sub PrintAreaLine(Optional ByVal bolAll As Boolean = False, Optional ByVal bolRight As Boolean = False)

    PrintNewLine
    If iPrint <> ciPrtY Then
       If bolAll = True Then
           Printer.Line (ciStartX, iPrint)-(m_TBWidth, iPrint)
       Else
           Printer.Line (ciStartX, iPrint)-(m_MidCol, iPrint)
       End If
       If bolRight = True Then
          '畫主動修正和告代欄位
          Printer.Line (m_MidCol, iPrint)-(m_MidCol, m_Tbottom)
       Else
          iPrint = iPrint - lngLineHeight / 2 '直接跳一行有點多
       End If
       
    End If
End Sub

'Added by Lydai 2017/12/27 外文本
Private Sub cmdOpen_Click()
Dim hLocalFile As Long 'Added by Lydia 2018/06/21

On Error GoTo ErrHand01 'Added by Lydia 2018/03/23 無權限的錯誤要改訊息

    'Added by Lydia 2020/01/20 開啟[原始檔區]
    If InStr(CmdOpen.Caption, "原始檔") > 0 Then
        If PUB_CheckFormExist("frm100101_M") Then
            MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
        Else
            If CmdOpen.Tag = "" Then
                MsgBox strCase(1) & "-" & strCase(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
            Else
                strExc(1) = ""
                frm100101_M.m_strKey = CmdOpen.Tag '多筆總收文號
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
'        strExc(1) = Pub_GetFCPcaseFilePath(strCase(2), , strCase(1))
'        If Dir(strExc(1) & "\*.*") <> "" Then
'             'Modified by Lydia 2018/06/21 用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
'             'SHELL "Explorer.exe " & strExc(1), vbNormalFocus  '開啟案件資料夾
'             ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
'        Else
'             MsgBox lblData(6).Caption & "在" & strExc(1) & "的資料夾不存在或無檔案!", vbInformation
'        End If
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

'Added by Lydia 2018/04/18 取得費用,規費和點數
Private Sub OnUpdateFee(ByVal strCP10 As String, ByVal strCP20 As String, ByRef nCP16 As String, ByRef nCP17 As String, ByRef nCP18 As String)
   Dim lngFee As Long
   
   nCP16 = ""
   nCP17 = ""
   nCP18 = ""
   '是否向客戶收款預設為N者,不預設費用
   If strCP20 <> "N" Then
      nCP17 = GetPatentOfficialFee(strCase(1), strCP10, "", strCase(5), strCase(5), strCase(11), strCase(10), strCase(2), strCase(3), strCase(4), n_CP118)
      lngFee = Val(GetFCPFee(strCase(1), strCP10)) + Val(nCP17)
      ' 費用
      If lngFee > 0 Then
         nCP16 = Format(lngFee)
         '點數
         nCP18 = Format((Val(lngFee) - Val(nCP17)) / 1000, "0.0")
      End If
   End If
End Sub

'Added by Lydia 2018/04/18 計算承辦期限
'Modified by Lydia 2021/08/23 +本所期限
'Mark by Lydia 2022/08/04 改成共用模組PUB_GetTCTbCP48
'Private Function GetCP48(ByVal iKind As String, ByVal iCP10 As String, ByRef iCP06 As String) As String
''Memo by Lydia 2022/04/28 工程師命名作業收文FMP主修和告代的承辦期限規則：提申前的承辦期限規則與FCP案一致；
''Memo by Lydia 2021/09/08 提申前的告代的承辦期限以系統日+(依案件國家收費表天數計算,ex.告代的FCP案5天,FG案6天,P案7天)個工作天 ;
'                                        '所限和承辦期限不可大於新案的所限,若期限超過系統日改用系統日; ex.FCP65474, FCP65176是8月收文等到說明書才進行命名作業
'                                        '提申前的主動修正的本所期限=提申的本所期限(和分案作業相同);承辦期限=本所期限往前-5個工作天
'
''end 2021/09/08
''Memo by Lydia 2021/08/23 規則
''主動修正:
''提申前:  本所期限 = 新案收文日起算15個工作天(本所期限若超過提申所限, 則以提申所限當本所期限)，承辦期限=本所期限往前-5個工作天
''提申後: 1. 新案翻譯已發文並且與申請日同一天，承辦期限 = 新案發文日起算15個工作天，本所期限 = 承辦期限 + 再加5個工作天
''              2. 新案翻譯未發文，本所期限 = 新案翻譯的本所期限，承辦期限=本所期限往前-5個工作天
''告代:
''提申前: 承辦期限 = 新案收文日起算(依案件國家收費表天數計算,FCP案5天,FG案6天,P案7天)個工作天，本所期限 = 承辦期限 + 再加5個工作天(本所期限若超過提申所限, 則以提申所限當本所期限)
''提申後(等到新案發文一併更新): 承辦期限 = 新案發文日起算6個工作天，本所期限 = 承辦期限 + 再加5個工作天
''增加判斷新案已提申，從命名收文【提申後告代】承辦期限 = 收文日起算6個工作天，本所期限 = 承辦期限 + 再加5個工作天(本所期限若超過新案翻譯的本所期限, 則以新案翻譯的所限當本所期限)
''end 2021/08/23
''Memo by Lydia 2018/04/18 規則
''告代：(提申前)收文日/(提申後)新案發文日起算6個工作天
''主動修正：(提申前)收文日/(提申後+新案翻譯已發文並且與申請日同一天)新案發文日起算15個工作天/
''                        (提申後+新案翻譯未發文)=新案翻譯的本所期限
''提申前的承辦期限若超過提申所限，則以提申所限當承辦期限
''end 2018/04/18
'Dim oStdDate As String '起始日
'Dim strDate1 As String
'Dim strTmp1 As String, strTmp2 As String
'  iCP06 = "" 'Added by Lydia 2021/08/23
'  If iKind = "1" Then    '提申後
'    'Added by Lydia 2022/04/28 工程師命名作業收文FMP主修和告代的承辦期限規則：提申前的承辦期限規則與FCP案一致；
'    If strCase(1) = "P" Then
'        strDate1 = Pub_GetFMPbCP48("1", strCase, iCP10, iCP06)
'    Else
'    'end 2022/04/28
'       If Val(strCase(9)) > 0 Then '有申請日才算期限
'            '告代
'            If iCP10 = "901" Then
'                 'Modified by Lydia 2021/08/23
'                 'strDate1 = Pub_GetHandleDay(strCase(1), strCase(6), iCP10, n_CP27) '承辦期限 = 新案發文日起算6個工作天
'                 '承辦期限 = 收文日起算6個工作天
'                 strDate1 = Pub_GetHandleDay(strCase(1), strCase(6), iCP10, strSrvDate(1))
'                 '本所期限 = 承辦期限 + 再加5個工作天
'                 iCP06 = PUB_GetFCPOurDeadline(strDate1, , , , "N")
'                 '本所期限若超過新案翻譯的本所期限, 則以新案翻譯的所限當本所期限
'                 If iCP06 > tCP06 Then iCP06 = tCP06
'            '主動修正
'            ElseIf iCP10 = "203" Then
'                 '(提申後+新案翻譯已發文=新案發文日)新案發文日起算15個工作天
'                 If Val(tCP27) = Val(strCase(9)) Then
'                      strDate1 = Pub_GetHandleDay(strCase(1), strCase(6), iCP10, tCP27)
'                      iCP06 = PUB_GetFCPOurDeadline(strDate1, , , , "N")  'Added by Lydia 2021/08/23 本所期限 = 承辦期限 + 再加5個工作天
'                 Else
'                 '(提申後+新案翻譯未發文)=新案翻譯的本所期限
'                      'Modified by Lydia 2021/08/23 新案翻譯未發文，本所期限 = 新案翻譯的本所期限，承辦期限=本所期限往前-5個工作天
'                      'strDate1 = tCP06
'                      iCP06 = tCP06
'                      strDate1 = CompWorkDay(6, iCP06, 1)  '承辦期限=本所期限往前-5個工作天
'                      'end 2021/08/23
'                 End If
'            End If
'       Else
'            strDate = ""
'       End If
'    End If 'Added by Lydia 2022/04/28
'  ElseIf iKind = "2" Then '提申前
'       'Added by Lydia 2021/08/23
'       If iCP10 = "203" Then   '主動修正
'            'Modified by Lydia 2021/09/08 以系統日+(依案件國家收費表天數計算)個工作天
'            'Mark by Lydia 2021/09/08 提申前的主動修正的本所期限=提申的本所期限(和分案作業相同);承辦期限=本所期限往前-5個工作天
'            'iCP06 = Pub_GetHandleDay(strCase(1), strCase(6), iCP10, TransDate(Replace(lblData(5).Caption, "/", ""), 2), TransDate(Replace(lblData(1).Caption, "/", ""), 2))
'            'If iCP06 > TransDate(Replace(lblData(1).Caption, "/", ""), 2) Then
'                 iCP06 = TransDate(Replace(lblData(1).Caption, "/", ""), 2)  '提申所限當本所期限
'            'End If 'Mark by Lydia 2021/09/08
'            If iCP06 < strSrvDate(1) Then iCP06 = strSrvDate(1) 'Added by Lydia 2021/09/08 若期限超過系統日改用系統日
'
'            '承辦期限=本所期限往前-5個工作天
'            strDate1 = CompWorkDay(6, iCP06, 1)
'            If strDate1 < strSrvDate(1) Then strDate1 = strSrvDate(1) '若期限超過系統日改用系統日
'
'       Else
'       'end 2021/08/23
'            'Memo by Lydia 2021/08/23 承辦期限 = 新案收文日起算6個工作天，本所期限 = 承辦期限 + 再加5個工作天(本所期限若超過提申所限, 則以提申所限當本所期限)
'            'Modified by Lydia 2021/09/08 以系統日+(依案件國家收費表天數計算)個工作天
'            'strDate1 = Pub_GetHandleDay(strCase(1), strCase(6), iCP10, TransDate(Replace(lblData(5).Caption, "/", ""), 2), TransDate(Replace(lblData(1).Caption, "/", ""), 2))
'            'If strDate1 > TransDate(Replace(lblData(1).Caption, "/", ""), 2) Then  '所限和承辦期限不可大於新案的所限
'            strDate1 = Pub_GetHandleDay(strCase(1), strCase(6), iCP10, strSrvDate(1), TransDate(Replace(lblData(1).Caption, "/", ""), 2))
'            If strDate1 > TransDate(Replace(lblData(1).Caption, "/", ""), 2) And TransDate(Replace(lblData(1).Caption, "/", ""), 2) <> "" Then  '所限和承辦期限不可大於新案的所限
'                strDate1 = TransDate(Replace(lblData(1).Caption, "/", ""), 2)
'            End If
'            If strDate1 < strSrvDate(1) Then strDate1 = strSrvDate(1) 'Added by Lydia 2021/09/08 若期限超過系統日改用系統日
'            'end 2021/09/08
'            'Added by Lydia 2021/08/23 本所期限 = 承辦期限 + 再加5個工作天
'            iCP06 = CompWorkDay(6, strDate1)  '本所期限 = 承辦期限 + 再加5個工作天
'            'Modified by Lydia 2021/09/08
'            'If iCP06 > TransDate(Replace(lblData(1).Caption, "/", ""), 2) Then
'            If iCP06 > TransDate(Replace(lblData(1).Caption, "/", ""), 2) And TransDate(Replace(lblData(1).Caption, "/", ""), 2) <> "" Then
'                 iCP06 = TransDate(Replace(lblData(1).Caption, "/", ""), 2)  '提申所限當本所期限
'            End If
'            If iCP06 < strSrvDate(1) Then iCP06 = strSrvDate(1) 'Added by Lydia 2021/09/08 若期限超過系統日改用系統日
'            'end 2021/08/23
'       End If 'Added by Lydia 2021/08/23
'  ElseIf iKind = "3" Then '當日告代
'       strDate1 = strSrvDate(1)
'  End If
'
'  GetCP48 = strDate1
'End Function
'End Mark by Lydia 2022/08/04

'Added by Lydia 2018/07/12 上傳相似比對結果檔案
Private Sub cmdFile_Click()
   
   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.lblCaseNo = strCase(1) & "-" & strCase(2) & "-" & strCase(3) & "-" & strCase(4)
   frm090801_8.Label4.Visible = False
   frm090801_8.bolNotPDF = True
   frm090801_8.Show vbModal
End Sub

'Added by Lydia 2019/09/17 讀取修改前的資料
Private Function ProcDataWord1() As String
Dim tmpArr1 As Variant
Dim intJ As Integer, intP As Integer
Dim nIndex As Integer

    For nIndex = 0 To m_TCTCount - 1
        intJ = Val(Replace(m_TCTList(nIndex).fiName, "TCT", ""))
        'Modified by Lydia 2021/04/09 +119 有序列表
        If intJ >= m_FS And intJ <= TF_TCT And InStr(TF_TCTnotFS & ",116,117,118,027,028,119", Format(intJ, "000")) = 0 Then '跳過認翻譯人員TCT27,TCT28
            strExc(1) = m_TCTList(nIndex).fiOldData
            'If intJ = 27 And (strExc(1) = "A" Or strExc(1) = "B") Then    '欲翻譯此案件者
            '   strExc(1) = m_TCT10 & "-" & strExc(1)
            'End If
            If GetDataDef("W", intJ, strExc(1)) = True Then
            End If
            '主動修正+是否不請款
            If intJ = 19 Then
                 '提申後/提申前
                 If GetDataDef("W", 117, m_TCTList(117 - m_FS - 4).fiOldData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
                 '不請款
                 If GetDataDef("W", 116, m_TCTList(116 - m_FS - 4).fiOldData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
            '彩圖提申(118)和一案兩請同一行
            ElseIf intJ = 21 Then
                 If GetDataDef("W", 118, m_TCTList(118 - m_FS - 4).fiOldData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
                 'Added by Lydia 2021/04/09 有序列表
                 If GetDataDef("W", 119, m_TCTList(119 - m_FS - 4).fiOldData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
                 'end 2021/04/09
'            '未異動：不顯示
'            ElseIf intJ = 16 Or intJ = 17 Or intJ = 18 Or intJ = 25 Then
'                 '案件名稱、設計案屬性、案件類別
'                 If m_TCTList(nIndex).fiOldData = m_TCTList(nIndex).fiNewData Then
'                     bolChgText = True
'                     GoTo JumpToNext
'                 End If
            End If
            
            If bolChgText = True Then
                '案件名稱(英)
                If intJ = 17 Then
                   strTempA = "案件名稱" & Trim(strTempA)
                Else
                   '----------不顯示
                   '標的不一致：若未勾選，則不顯示資料
                   If (intJ = 84 Or intJ = 86) And m_TCTList(82 - m_FS).fiOldData = "" Then
                      strTempA = ""
                   End If
                   '非設計案-不顯示設計案屬性
                   If intJ = 18 And strCase(5) <> "3" Then
                      strTempA = ""
                   End If
                   '----------不顯示
                End If
                ProcDataWord1 = ProcDataWord1 & IIf(Trim(strTempA) <> "", vbCrLf, "") & strTempA
            End If
JumpToNext:
        End If
    Next nIndex
End Function

'Added by Lydia 2019/09/17 讀取修改後的資料
Private Function ProcDataWord2() As String
Dim tmpArr1 As Variant
Dim intJ As Integer, intP As Integer
Dim nIndex As Integer

    For nIndex = 0 To m_TCTCount - 1
        intJ = Val(Replace(m_TCTList(nIndex).fiName, "TCT", ""))
        'Modified by Lydia 2021/04/09 +119 有序列表
        If intJ >= m_FS And intJ <= TF_TCT And InStr(TF_TCTnotFS & ",116,117,118,027,028,119", Format(intJ, "000")) = 0 Then  '跳過認翻譯人員TCT27,TCT28
            strExc(1) = m_TCTList(nIndex).fiNewData
            'If intJ = 27 And (strExc(1) = "A" Or strExc(1) = "B") Then    '欲翻譯此案件者
            '   strExc(1) = m_TCT10 & "-" & strExc(1)
            'End If
            If GetDataDef("W", intJ, strExc(1)) = True Then
            End If
            '主動修正+是否不請款
            If intJ = 19 Then
                 '提申後/提申前
                 If GetDataDef("W", 117, m_TCTList(117 - m_FS - 4).fiNewData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
                 '不請款
                 If GetDataDef("W", 116, m_TCTList(116 - m_FS - 4).fiNewData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
            '彩圖提申(118)和一案兩請同一行
            ElseIf intJ = 21 Then
                 If GetDataDef("W", 118, m_TCTList(118 - m_FS - 4).fiNewData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
                 'Added by Lydia 2021/04/09 有序列表
                 If GetDataDef("W", 119, m_TCTList(119 - m_FS - 4).fiNewData) = True Then  '排除不修改的欄位(tf_tctnotfs)
                 End If
                 'end 2021/04/09
            '未異動：不顯示
'            ElseIf intJ = 16 Or intJ = 17 Or intJ = 18 Or intJ = 25 Then
'                 '案件名稱、設計案屬性、案件類別
'                 If m_TCTList(nIndex).fiOldData = m_TCTList(nIndex).fiNewData Then
'                     bolChgText = True
'                     GoTo JumpToNext
'                 End If
            End If
            
            If bolChgText = True Then
                '案件名稱
                If intJ = 17 Then
                   strTempA = "案件名稱" & Trim(strTempA)
                Else
                   '----------不顯示
                   '標的不一致：若未勾選，則不顯示資料
                   If (intJ = 84 Or intJ = 86) And m_TCTList(82 - m_FS).fiNewData = "" Then
                      strTempA = ""
                   End If
                   '非設計案-不顯示設計案屬性
                   If intJ = 18 And strCase(5) <> "3" Then
                      strTempA = ""
                   End If
                   '----------不顯示
                End If
                ProcDataWord2 = ProcDataWord2 & IIf(Trim(strTempA) <> "", vbCrLf, "") & strTempA
            End If
        End If
JumpToNext:
    Next nIndex
End Function

'Added by Lydia 2020/02/17
Private Sub CmdPA174_Click()
    Call ProcPA174toFile("N")
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟/維護FCP0xxxxx.新案性質.案件名稱.doc
Private Sub ProcPA174toFile(ByVal pKind As String)
Dim strKind As String

    If ChkPA174.Value = vbUnchecked Then
        MsgBox "請先勾選「有特殊字」!", vbInformation + vbOKOnly, Me.Caption
    Else
        If Frame6.Enabled = False Then '提申後，除案件名稱外，其餘皆可修改。「名稱有特殊字」比照辦理
            strKind = "0"
        ElseIf pKind = "Y" Then 'bolAskPA174
            strKind = "3"
        Else
            strKind = IIf(m_iStiu = "M" Or cmdOK(3).Visible = True, "1", "0")
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

'Added by Lydia 2025/10/15 取得指定翻譯人員的名稱/ID
Private Function GetTCT27val(Optional ByRef pConID As String) As String
Dim strB1 As String
   
   pConID = ""
   GetTCT27val = ""
   If txtData(2) = "A" Or txtData(2) = "B" Then
      pConID = m_TCT10
      GetTCT27val = lblData(8).Caption & "-" & IIf(txtData(2) = "A", "下班", "上班")
   ElseIf txtData(47) <> "" Then
      pConID = GetPrjSalesNM_2(Mid(txtData(47).Text, 1, Len(txtData(47).Text) - 1), , , , True)
      GetTCT27val = Mid(txtData(47).Text, 1, Len(txtData(47).Text) - 1) & "-" & IIf(Right(UCase(txtData(47)), 1) = "A", "下班", "上班")
   Else
      For Each oChk In Chk27
         If oChk.Value = vbChecked Then
            pConID = Pub_SetF51Order("", oChk.Index)
            GetTCT27val = oChk.Caption
         End If
      Next
   End If

End Function
