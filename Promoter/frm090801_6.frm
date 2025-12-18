VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm090801_6 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "減免資格設定"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4125
      TabIndex        =   1
      Top             =   5175
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3135
      TabIndex        =   0
      Top             =   5175
      Width           =   930
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5025
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   8864
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "中小企業"
      TabPicture(0)   =   "frm090801_6.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkAD15JP1(14)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkAD15JP1(13)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkAD15JP1(12)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkAD15JP1(11)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkAD15JP1(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkAD15JP1(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkAD15JP1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkAD15JP1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkAD15JP1(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkAD15JP1(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkAD15JP1(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkAD15JP1(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkAD15JP1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkAD15JP1(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblNotice(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "獨資企業"
      TabPicture(1)   =   "frm090801_6.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkAD15JP2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkAD15JP2(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkAD15JP2(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkAD15JP2(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkAD15JP2(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkAD15JP2(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkAD15JP2(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblNotice(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblNotice(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "小企業"
      TabPicture(2)   =   "frm090801_6.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkAD15JP3(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkAD15JP3(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblNotice(3)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblNotice(4)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "新興企業"
      TabPicture(3)   =   "frm090801_6.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkAD15JP4(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "chkAD15JP4(2)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblNotice(5)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblNotice(6)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "大學"
      TabPicture(4)   =   "frm090801_6.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkAD15JP5(1)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblNotice(7)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "個人"
      TabPicture(5)   =   "frm090801_6.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "chkAD15JP6(1)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "chkAD15JP6(2)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "chkAD15JP6(3)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "lblNotice(8)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "lblNotice(9)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "台灣"
      TabPicture(6)   =   "frm090801_6.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Frame1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame2"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "台灣專利中小企業符合減免之資格"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   270
         TabIndex        =   51
         Top             =   750
         Width           =   3885
         Begin VB.TextBox txtAD16 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   270
            Index           =   5
            Left            =   2280
            MaxLength       =   9
            TabIndex        =   53
            Top             =   450
            Width           =   1005
         End
         Begin VB.TextBox txtAD16 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   270
            Index           =   6
            Left            =   810
            MaxLength       =   3
            TabIndex        =   52
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox chkAD15 
            Caption         =   "依法辦理公司登記或商業登記，實收資本額在新臺幣1億元以下：                        元"
            Height          =   555
            Index           =   5
            Left            =   120
            TabIndex        =   54
            Top             =   180
            Width           =   3540
         End
         Begin VB.CheckBox chkAD15 
            Caption         =   "經常僱用員工數未滿200人之事業：員工數          人"
            Height          =   555
            Index           =   6
            Left            =   120
            TabIndex        =   55
            Top             =   690
            Width           =   3390
         End
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "旅館業（總資本額5,000萬日圓以下）"
         Height          =   200
         Index           =   14
         Left            =   -74760
         TabIndex        =   40
         Top             =   3780
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "旅館業（員工200人以下）"
         Height          =   200
         Index           =   13
         Left            =   -74760
         TabIndex        =   39
         Top             =   3570
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "軟體或資料處理業（總資本額3億日圓以下）"
         Height          =   200
         Index           =   12
         Left            =   -74760
         TabIndex        =   38
         Top             =   3360
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "軟體或資料處理業（員工300人以下）"
         Height          =   200
         Index           =   11
         Left            =   -74760
         TabIndex        =   37
         Top             =   3150
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "橡膠製造業（汽車和飛機輪胎、內胎和工業用皮帶的製造業除外）（總資本額3億日圓以下）"
         Height          =   375
         Index           =   10
         Left            =   -74760
         TabIndex        =   36
         Top             =   2760
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "橡膠製造業（汽車和飛機輪胎、內胎和工業用皮帶的製造業除外）（員工900人以下）"
         Height          =   375
         Index           =   9
         Left            =   -74760
         TabIndex        =   35
         Top             =   2370
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "零售業（總資本額5,000萬日圓以下）"
         Height          =   200
         Index           =   8
         Left            =   -74760
         TabIndex        =   34
         Top             =   2160
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "零售業（員工50人以下）"
         Height          =   200
         Index           =   7
         Left            =   -74760
         TabIndex        =   33
         Top             =   1950
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "服務業（總資本額5,000萬日圓以下）"
         Height          =   200
         Index           =   6
         Left            =   -74760
         TabIndex        =   32
         Top             =   1740
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "服務業（員工100人以下）"
         Height          =   200
         Index           =   5
         Left            =   -74760
         TabIndex        =   31
         Top             =   1530
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "批發業（總資本額1億日圓以下）"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   30
         Top             =   1320
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "批發業（員工100人以下）"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   29
         Top             =   1110
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "製造業、建築業、運輸業（總資本額3億日圓以下）"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   28
         Top             =   900
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP1 
         Caption         =   "製造業、建築業、運輸業（員工300人以下）"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   27
         Top             =   690
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP2 
         Caption         =   "製造業、建築業、運輸業（員工300人以下）"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   26
         Top             =   1290
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP2 
         Caption         =   "批發業（員工100人以下）"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   25
         Top             =   1500
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP2 
         Caption         =   "服務業（員工100人以下）"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   24
         Top             =   1710
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP2 
         Caption         =   "零售業（員工50人以下）"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   23
         Top             =   1920
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP2 
         Caption         =   "橡膠製造業（汽車和飛機輪胎、內胎和工業用皮帶的製造業除外）（員工900人以下）"
         Height          =   375
         Index           =   5
         Left            =   -74760
         TabIndex        =   22
         Top             =   2130
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP2 
         Caption         =   "旅館業（員工200人以下）"
         Height          =   200
         Index           =   7
         Left            =   -74760
         TabIndex        =   21
         Top             =   2730
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP2 
         Caption         =   "軟體或資料處理業（員工300人以下）"
         Height          =   200
         Index           =   6
         Left            =   -74760
         TabIndex        =   20
         Top             =   2520
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP3 
         Caption         =   "一般企業：員工20人以下（貿易或服務業公司員工5人以下）"
         Height          =   405
         Index           =   1
         Left            =   -74760
         TabIndex        =   19
         Top             =   870
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP3 
         Caption         =   $"frm090801_6.frx":00C4
         Height          =   405
         Index           =   2
         Left            =   -74760
         TabIndex        =   18
         Top             =   1950
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP4 
         Caption         =   "中小型企業：公司成立未滿10年且總資本額3億日圓以下"
         Height          =   405
         Index           =   1
         Left            =   -74760
         TabIndex        =   17
         Top             =   870
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP4 
         Caption         =   "獨資企業：公司成立未滿10年"
         Height          =   240
         Index           =   2
         Left            =   -74760
         TabIndex        =   16
         Top             =   2250
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP5 
         Caption         =   "大學"
         Height          =   240
         Index           =   1
         Left            =   -74760
         TabIndex        =   15
         Top             =   870
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP6 
         Caption         =   "年所得合計未滿日幣150萬"
         Height          =   240
         Index           =   1
         Left            =   -74760
         TabIndex        =   14
         Top             =   870
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP6 
         Caption         =   "年所得合計未滿日幣250萬"
         Height          =   240
         Index           =   2
         Left            =   -74760
         TabIndex        =   13
         Top             =   1920
         Width           =   4500
      End
      Begin VB.CheckBox chkAD15JP6 
         Caption         =   "獨資企業房屋土地交易所得額及營利事業所得額合計未滿日幣290萬"
         Height          =   360
         Index           =   3
         Left            =   -74760
         TabIndex        =   12
         Top             =   2190
         Width           =   4500
      End
      Begin VB.Frame Frame1 
         Caption         =   "台灣專利中小企業符合減免之資格(舊)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   300
         TabIndex        =   3
         Top             =   2220
         Visible         =   0   'False
         Width           =   3525
         Begin VB.TextBox txtAD16 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   270
            Index           =   4
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   8
            Top             =   2250
            Width           =   375
         End
         Begin VB.TextBox txtAD16 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   270
            Index           =   3
            Left            =   2835
            MaxLength       =   3
            TabIndex        =   7
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtAD16 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   270
            Index           =   2
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   6
            Top             =   900
            Width           =   1005
         End
         Begin VB.TextBox txtAD16 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   4
            Top             =   420
            Width           =   1005
         End
         Begin VB.CheckBox chkAD15 
            Caption         =   "製造業、營造業、礦業及土石採取業實收資本額八千萬以下：                        元"
            Height          =   375
            Index           =   1
            Left            =   90
            TabIndex        =   5
            Top             =   240
            Width           =   3390
         End
         Begin VB.CheckBox chkAD15 
            Caption         =   "前項除外之其他行業前一年營業額一億元以下：                        元"
            Height          =   375
            Index           =   2
            Left            =   90
            TabIndex        =   9
            Top             =   720
            Width           =   3390
         End
         Begin VB.CheckBox chkAD15 
            Caption         =   "我國製造業、營造業、礦業及土石採取業實收資本額新台幣八千萬以上但經常僱用員工數未滿200人：員工數          人"
            Height          =   765
            Index           =   3
            Left            =   90
            TabIndex        =   11
            Top             =   1110
            Width           =   3390
         End
         Begin VB.CheckBox chkAD15 
            Caption         =   "我國前項除外之其他行業前一年營業額一億元以上者但經常僱用員工數未滿100人：員工數          人"
            Height          =   555
            Index           =   4
            Left            =   90
            TabIndex        =   10
            Top             =   1890
            Width           =   3390
         End
      End
      Begin VB.Label lblNotice 
         Caption         =   $"frm090801_6.frx":00FF
         Height          =   945
         Index           =   1
         Left            =   -74730
         TabIndex        =   50
         Top             =   4020
         Width           =   4545
      End
      Begin VB.Label lblNotice 
         Caption         =   "減免額度：實體審查規費減免50%，年費（第1-10年）規費減免50%。"
         Height          =   465
         Index           =   0
         Left            =   -74760
         TabIndex        =   49
         Top             =   3060
         Width           =   4485
      End
      Begin VB.Label lblNotice 
         Caption         =   "減免資格的行業別中小企業，惟僅限制各行業別的員工數，不審核資本額"
         Height          =   465
         Index           =   2
         Left            =   -74760
         TabIndex        =   48
         Top             =   780
         Width           =   4485
      End
      Begin VB.Label lblNotice 
         Caption         =   $"frm090801_6.frx":01D3
         Height          =   735
         Index           =   3
         Left            =   -74760
         TabIndex        =   47
         Top             =   1290
         Width           =   4605
      End
      Begin VB.Label lblNotice 
         Caption         =   "減免額度：實體審查規費減免66%，年費（第1-10年）規費減免66%。"
         Height          =   405
         Index           =   4
         Left            =   -74760
         TabIndex        =   46
         Top             =   2430
         Width           =   4605
      End
      Begin VB.Label lblNotice 
         Caption         =   $"frm090801_6.frx":0269
         Height          =   735
         Index           =   5
         Left            =   -74760
         TabIndex        =   45
         Top             =   1320
         Width           =   4605
      End
      Begin VB.Label lblNotice 
         Caption         =   "減免額度：實體審查規費減免66%，年費（第1-10年）規費減免66%。"
         Height          =   405
         Index           =   6
         Left            =   -74760
         TabIndex        =   44
         Top             =   2550
         Width           =   4605
      End
      Begin VB.Label lblNotice 
         Caption         =   $"frm090801_6.frx":02FF
         Height          =   675
         Index           =   7
         Left            =   -74760
         TabIndex        =   43
         Top             =   1170
         Width           =   4605
      End
      Begin VB.Label lblNotice 
         Caption         =   "減免額度：免繳實體審查規費及第1-3年年費，第4-10年年費減免50%。"
         Height          =   435
         Index           =   8
         Left            =   -74760
         TabIndex        =   42
         Top             =   1170
         Width           =   4605
      End
      Begin VB.Label lblNotice 
         Caption         =   "減免額度：實體審查規費減免50%，年費（第1-10年）減免50%。"
         Height          =   435
         Index           =   9
         Left            =   -74760
         TabIndex        =   41
         Top             =   2640
         Width           =   4605
      End
   End
End
Attribute VB_Name = "frm090801_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/22 改成Form2.0 (無)
'Create by Morgan 2013/4/2
'Modified by Morgan 2019/4/11 +日本案減免資格
Option Explicit

Public m_stAD15 As String, m_stAD16 As String
Public m_stAD02 As String, m_stAD10 As String, m_stCU15 As String 'Added by Morgan 2019/4/11
Dim m_MousePointer As Integer
Dim m_bolActivated As Boolean
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2022/9/1


'Add By Sindy 2022/9/1
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub chkAD15_Click(Index As Integer)
   Dim oCheck As CheckBox
   
   If SSTab2.Enabled = True Then
      If chkAD15(Index).Value = 1 Then
         txtAD16(Index).Enabled = True
         txtAD16(Index).SetFocus
         For Each oCheck In chkAD15
            If oCheck.Index <> Index And oCheck.Value = 1 Then
               oCheck.Value = 0
            End If
         Next
      Else
         txtAD16(Index).Text = ""
         txtAD16(Index).Enabled = False
      End If
   End If
End Sub

Private Sub chkAD15JP1_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP1(Index).Value = 1 Then
      For Each oCheck In chkAD15JP1
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP2_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP2(Index).Value = 1 Then
      For Each oCheck In chkAD15JP2
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP3_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP3(Index).Value = 1 Then
      For Each oCheck In chkAD15JP3
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP4_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP4(Index).Value = 1 Then
      For Each oCheck In chkAD15JP4
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP5_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP5(Index).Value = 1 Then
      For Each oCheck In chkAD15JP5
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP6_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP6(Index).Value = 1 Then
      For Each oCheck In chkAD15JP6
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
      
   Dim bCancel As Boolean, iIdx As Integer
   
   If Index = 0 Then
      If m_stAD02 = "000" Then
         'Modified by Morgan 2020/7/24
         'For iIdx = 1 To 4
         For iIdx = 1 To 6
            If chkAD15(iIdx) = 1 Then
               If txtAD16(iIdx) = "" Then
                  MsgBox "請輸入資本額/員工數！", vbExclamation
                  txtAD16(iIdx).SetFocus
                  Exit Sub
               End If
               
               txtAD16_Validate iIdx, bCancel
               If bCancel = True Then
                  txtAD16(iIdx).SetFocus
                  Exit Sub
               Else
                  m_PrevForm.m_stAD15 = iIdx
                  'Modified by Morgan 2013/6/19 若用貼的會有逗號
                  'm_PrevForm.m_stAD16 = Val(txtAD16(iIdx).Text)
                  m_PrevForm.m_stAD16 = Val(Format(txtAD16(iIdx).Text))
                  Exit For
               End If
            End If
         Next
         'If iIdx > 4 Then
         If iIdx > 6 Then
            'm_PrevForm.m_stAD15 = ""
            'm_PrevForm.m_stAD16 = ""
            MsgBox "請點選資格！", vbExclamation
            Exit Sub
         End If
         
      'Added by Morgan 2019/4/12
      ElseIf m_stAD02 = "011" Then
         iIdx = GetAD15()
         If iIdx > 0 Then
            m_PrevForm.m_stAD10 = SSTab2.Tab + 1 '減免身分
            m_PrevForm.m_stAD15 = iIdx '減免資格
         Else
            MsgBox "請點選資格！", vbExclamation
            Exit Sub
         End If
      End If
   End If
   
   Unload Me
   Screen.MousePointer = m_MousePointer
End Sub

Private Sub Form_Activate()
   If m_bolActivated = False Then
      m_bolActivated = True
      If Val(m_stAD15) > 0 Then
         If m_stAD02 = "000" Then
            If m_stAD15 >= "5" Then 'Added by Morgan 2020/7/27
               chkAD15(Val(m_stAD15)).Value = vbChecked
               txtAD16(Val(m_stAD15)).Text = m_stAD16
            End If 'Added by Morgan 2020/7/27
            
         ElseIf m_stAD02 = "011" Then
            '中小企業
            If m_stAD10 = "1" Then
               SSTab2.Tab = 0
               chkAD15JP1(Val(m_stAD15)).Value = vbChecked
            '獨資企業
            ElseIf m_stAD10 = "2" Then
               SSTab2.Tab = 1
               chkAD15JP2(Val(m_stAD15)).Value = vbChecked
            '小企業
            ElseIf m_stAD10 = "3" Then
               SSTab2.Tab = 2
               chkAD15JP3(Val(m_stAD15)).Value = vbChecked
            '新興企業
            ElseIf m_stAD10 = "4" Then
               SSTab2.Tab = 3
               chkAD15JP4(Val(m_stAD15)).Value = vbChecked
            '學校
            ElseIf m_stAD10 = "5" Then
               SSTab2.Tab = 4
               chkAD15JP5(Val(m_stAD15)).Value = vbChecked
            '個人
            ElseIf m_stAD10 = "6" Then
               SSTab2.Tab = 5
               chkAD15JP6(Val(m_stAD15)).Value = vbChecked
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me, True
   m_MousePointer = Screen.MousePointer
   Screen.MousePointer = vbDefault
   SetTab 'Added by Morgan 2019/4/11
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090801_6 = Nothing
End Sub

Private Sub txtAD16_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtAD16(Index)
End Sub

Private Sub txtAD16_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(",") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtAD16_Validate(Index As Integer, Cancel As Boolean)
   If txtAD16(Index) <> "" Then 'Added by Morgan 2013/6/19
      Select Case Index
         Case 1
            'Modified by Morgan 2013/6/19 若用貼的會有逗號
            'If Val(txtAD16(index)) > 80000000 Then
            If Val(Format(txtAD16(Index))) > 80000000 Then
               MsgBox "金額輸入錯誤，不可高於八千萬！", vbExclamation
               Cancel = True
            ElseIf Val(Format(txtAD16(Index))) < 1000 Then
               MsgBox "金額輸入錯誤，不可低於1000！", vbExclamation
               Cancel = True
            End If
         Case 2
            'Modified by Morgan 2013/6/19 若用貼的會有逗號
            'If Val(txtAD16(index)) > 100000000 Then
            If Val(Format(txtAD16(Index))) > 100000000 Then
               MsgBox "金額輸入錯誤，不可高於一億！", vbExclamation
               Cancel = True
            'Added by Morgan 2013/6/19
            ElseIf Val(Format(txtAD16(Index))) < 1000 Then
               MsgBox "金額輸入錯誤，不可低於1000！", vbExclamation
               Cancel = True
            End If
         Case 3
            If Val(txtAD16(Index)) >= 200 Then
               MsgBox "員工數輸入錯誤，必須小於200！", vbExclamation
               Cancel = True
            'Added by Morgan 2013/6/19
            ElseIf Val(txtAD16(Index)) = 0 Then
               MsgBox "員工數輸入錯誤，必須至少1人！", vbExclamation
               Cancel = True
            End If
         Case 4
            If Val(txtAD16(Index)) >= 100 Then
               MsgBox "員工數輸入錯誤，必須小於100！", vbExclamation
               Cancel = True
            'Added by Morgan 2013/6/19
            ElseIf Val(txtAD16(Index)) = 0 Then
               MsgBox "員工數輸入錯誤，必須至少1人！", vbExclamation
               Cancel = True
            End If
         'Added by Morgan 2020/7/24
         Case 5
            If Val(txtAD16(Index)) > 100000000 Then
               MsgBox "金額輸入錯誤，不可高於一億！", vbExclamation
               Cancel = True
            ElseIf Val(Format(txtAD16(Index))) < 1000 Then
               MsgBox "金額輸入錯誤，不可低於1000！", vbExclamation
               Cancel = True
            End If
         Case 6
            If Val(txtAD16(Index)) >= 200 Then
               MsgBox "員工數輸入錯誤，必須小於200！", vbExclamation
               Cancel = True
            ElseIf Val(txtAD16(Index)) = 0 Then
               MsgBox "員工數輸入錯誤，必須至少1人！", vbExclamation
               Cancel = True
            End If
         'end 2020/7/24
      End Select
   End If
End Sub

'Added by Morgan 2019/4/11
'設定頁籤
Private Sub SetTab()
   Dim ii As Integer
   '台灣
   If m_stAD02 = "000" Then
      For ii = 0 To SSTab2.Tabs - 1
         If ii = 6 Then
            SSTab2.TabVisible(ii) = True
         Else
            SSTab2.TabVisible(ii) = False
         End If
      Next
   '日本
   ElseIf m_stAD02 = "011" Then
      SSTab2.TabVisible(6) = False '台灣頁籤設不可見
      For ii = 0 To 5
         '個人
         If m_stCU15 = "0" Then
            '個人/大學
            If ii = 4 Or ii = 5 Then
               SSTab2.TabVisible(ii) = True
               If ii = 5 Then SSTab2.Tab = 5 'Added by Morgan 2019/4/23 個人預設個人
            Else
               SSTab2.TabVisible(ii) = False
            End If
         '學校
         ElseIf m_stCU15 = "2" Then
            '大學
            If ii = 4 Then
               SSTab2.TabVisible(ii) = True
            Else
               SSTab2.TabVisible(ii) = False
            End If
            
         'Added by Morgan 2021/5/25 未設定身分
         ElseIf m_stCU15 = "" Then
            SSTab2.TabVisible(ii) = True
         'end 2021/5/25
         ElseIf ii = 4 Or ii = 5 Then
            SSTab2.TabVisible(ii) = False
         Else
            SSTab2.TabVisible(ii) = True
         End If
      Next
   End If
End Sub

Private Function GetAD15() As Integer
   Dim oCheck As CheckBox
   '中小企業
   If SSTab2.Tab = 0 Then
      For Each oCheck In chkAD15JP1
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '獨資企業
   ElseIf SSTab2.Tab = 1 Then
      For Each oCheck In chkAD15JP2
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '小企業
   ElseIf SSTab2.Tab = 2 Then
      For Each oCheck In chkAD15JP3
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '新興企業
   ElseIf SSTab2.Tab = 3 Then
      For Each oCheck In chkAD15JP4
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '大學
   ElseIf SSTab2.Tab = 4 Then
      For Each oCheck In chkAD15JP5
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '個人
   ElseIf SSTab2.Tab = 5 Then
      For Each oCheck In chkAD15JP6
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   End If
End Function
