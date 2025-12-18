VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210156 
   BorderStyle     =   1  '單線固定
   Caption         =   "專業部主管分案作業"
   ClientHeight    =   6600
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9168
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9168
   Begin VB.CommandButton cmdOK 
      Caption         =   "接洽單(&M)"
      Height          =   345
      Index           =   9
      Left            =   7022
      TabIndex        =   20
      Top             =   1920
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "專利相關案"
      Height          =   345
      Index           =   7
      Left            =   8040
      TabIndex        =   21
      Top             =   1920
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   30
      TabIndex        =   36
      Top             =   2580
      Width           =   9045
      _ExtentX        =   15960
      _ExtentY        =   7070
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "主管分案（複選不可執行更正和通知補件）"
      TabPicture(0)   =   "frm210156.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRD1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "退智權 ＆ 退程序 (補件)"
      TabPicture(1)   =   "frm210156.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD2"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm210156.frx":0038
         Height          =   3525
         Left            =   60
         TabIndex        =   37
         Top             =   390
         Width           =   8895
         _ExtentX        =   15685
         _ExtentY        =   6223
         _Version        =   393216
         Cols            =   16
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   $"frm210156.frx":004D
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
         _Band(0).Cols   =   16
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD2 
         Bindings        =   "frm210156.frx":00E4
         Height          =   3525
         Left            =   -74970
         TabIndex        =   46
         Top             =   390
         Width           =   8895
         _ExtentX        =   15685
         _ExtentY        =   6223
         _Version        =   393216
         Cols            =   16
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   $"frm210156.frx":00F9
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
         _Band(0).Cols   =   16
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "案件收合"
      Height          =   285
      Left            =   90
      TabIndex        =   35
      Top             =   2310
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "全選"
      Height          =   345
      Index           =   0
      Left            =   6080
      TabIndex        =   19
      Top             =   1920
      Width           =   885
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   360
      Left            =   4200
      TabIndex        =   10
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "設定關連"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   11
      Left            =   9105
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   1920
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "取消關連"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   345
      Index           =   10
      Left            =   10200
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   1920
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "程序補件"
      Height          =   360
      Index           =   3
      Left            =   6300
      TabIndex        =   12
      Top             =   60
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本檔(&B)"
      Height          =   345
      Index           =   5
      Left            =   7022
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   1560
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   345
      Index           =   6
      Left            =   8040
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   1560
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "完整卷宗"
      Height          =   345
      Index           =   4
      Left            =   6080
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   1560
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "智權補件"
      Height          =   360
      Index           =   2
      Left            =   5370
      TabIndex        =   11
      Top             =   60
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "確定分案"
      Height          =   360
      Index           =   1
      Left            =   7230
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   60
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   8
      Left            =   8145
      TabIndex        =   14
      Top             =   60
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Caption         =   "維護區"
      Height          =   1815
      Left            =   60
      TabIndex        =   22
      Top             =   450
      Width           =   5955
      Begin VB.CheckBox Check1 
         Caption         =   "取消相同案"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   4830
         TabIndex        =   51
         Top             =   870
         Width           =   1125
      End
      Begin VB.CommandButton cmdUpdRow 
         Caption         =   "更正"
         Height          =   285
         Left            =   4860
         TabIndex        =   50
         Top             =   120
         Width           =   675
      End
      Begin VB.TextBox txtCRC11 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   3
         Top             =   780
         Width           =   732
      End
      Begin VB.TextBox txtCRC12 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   4
         Top             =   780
         Width           =   732
      End
      Begin VB.TextBox txtCRC10 
         Height          =   285
         Left            =   4350
         MaxLength       =   1
         TabIndex        =   2
         Top             =   480
         Width           =   435
      End
      Begin VB.ComboBox cboCRL68 
         Height          =   300
         ItemData        =   "frm210156.frx":0194
         Left            =   4065
         List            =   "frm210156.frx":0196
         TabIndex        =   9
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cboCRL67 
         Height          =   300
         ItemData        =   "frm210156.frx":0198
         Left            =   4065
         List            =   "frm210156.frx":019A
         TabIndex        =   7
         Top             =   1110
         Width           =   1815
      End
      Begin VB.ComboBox cboAttr 
         Height          =   300
         ItemData        =   "frm210156.frx":019C
         Left            =   1125
         List            =   "frm210156.frx":019E
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cboCRL55 
         Height          =   300
         ItemData        =   "frm210156.frx":01A0
         Left            =   1125
         List            =   "frm210156.frx":01A2
         TabIndex        =   6
         Top             =   1110
         Width           =   1815
      End
      Begin VB.TextBox txtCRC09 
         Height          =   285
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " 新穎性："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   14
         Left            =   3060
         TabIndex        =   52
         Top             =   1620
         Width           =   765
      End
      Begin MSForms.ComboBox cboCRC09 
         Height          =   300
         Left            =   1125
         TabIndex        =   1
         Top             =   450
         Width           =   1815
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3201;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "加乘註記："
         Height          =   180
         Index           =   13
         Left            =   3060
         TabIndex        =   49
         Top             =   810
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "計  件  值："
         Height          =   180
         Index           =   12
         Left            =   90
         TabIndex        =   48
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數：　　　(N:不算)"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   8
         Left            =   3060
         TabIndex        =   42
         Top             =   525
         Width           =   2445
      End
      Begin MSForms.Label lblData 
         Height          =   280
         Index           =   1
         Left            =   3330
         TabIndex        =   41
         Top             =   210
         Width           =   1380
         ForeColor       =   16711680
         VariousPropertyBits=   27
         Caption         =   "lblData(1)"
         Size            =   "2434;494"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "案件性質："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   10
         Left            =   2370
         TabIndex        =   40
         Top             =   210
         Width           =   900
      End
      Begin MSForms.Label lblData 
         Height          =   285
         Index           =   0
         Left            =   1110
         TabIndex        =   39
         Top             =   210
         Width           =   1170
         ForeColor       =   16711680
         VariousPropertyBits=   27
         Caption         =   "lblData(0)"
         Size            =   "2064;503"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   9
         Left            =   90
         TabIndex        =   38
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "擬制喪失"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   5
         Left            =   3090
         TabIndex        =   30
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "一案兩請："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   3060
         TabIndex        =   29
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件屬性："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相同案號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   11
         Left            =   90
         TabIndex        =   25
         Top             =   1170
         Width           =   900
      End
      Begin MSForms.Label lblName 
         Height          =   300
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   930
         VariousPropertyBits=   27
         Caption         =   "lblName"
         Size            =   "1640;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人員："
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   525
         Width           =   1080
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "急件顯示為整列紅色"
      ForeColor       =   &H008080FF&
      Height          =   180
      Index           =   1
      Left            =   7500
      TabIndex        =   47
      Top             =   2340
      Width           =   1620
   End
   Begin MSForms.TextBox txtNote 
      Height          =   885
      Left            =   6060
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   630
      Width           =   2835
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "5001;1561"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   3
      Left            =   3450
      TabIndex        =   44
      Top             =   2340
      Width           =   840
      ForeColor       =   16711680
      VariousPropertyBits=   27
      Caption         =   "lblData(3)"
      Size            =   "1482;503"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   2
      Left            =   1770
      TabIndex        =   43
      Top             =   2310
      Width           =   930
      ForeColor       =   16711680
      VariousPropertyBits=   27
      Caption         =   "lblData(2)"
      Size            =   "1640;503"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(註:複選請先按Ctrl鍵，雙擊開啟接洽單)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   4290
      TabIndex        =   34
      Top             =   2340
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總點數："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   7
      Left            =   2730
      TabIndex        =   33
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總費用："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   32
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "您的意見："
      Height          =   180
      Index           =   6
      Left            =   6090
      TabIndex        =   31
      Top             =   450
      Width           =   990
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   1170
      TabIndex        =   0
      Top             =   60
      Width           =   2025
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3572;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "所　別："
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frm210156"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2022/09/30 專業部主管分案作業
Option Explicit
Const m_F0302 = "3"  '表單類別:接洽單
Dim m_nowF0308 As String '下一處理人員(操作者簽核身份)
Dim m_F0309 As String '目前狀態
Dim m_F0316 As String '智權人員

Dim m_ST06 As String, m_SysNo As String '操作者所別,系統別

Dim m_blnColOrderAsc1 As Boolean, m_blnColOrderAsc2 As Boolean '欄位資料由小到大排序
Dim mSelCount As Integer '勾選Ｘ筆收文
Dim intQ As Integer, intP As Integer
Dim strQuery As String, strQ2 As String
Dim oObj As Object
Dim lngColor As Long
Dim strBCase(0 To 4) As String
Dim blnDBClick As Boolean
Dim intLastTop As Integer '記錄最後勾選位置的第一列 ; 12/16 改成「更正勾選列」移到第一列
Dim m_SelCP09 As String '「更正勾選列」的收文號

Public cmdState As Integer '紀錄作用按鍵
Dim rsAD1 As New ADODB.Recordset
Dim rsQuery As New ADODB.Recordset
Dim strUpdTime As String
Dim m_PrevNo As String '先保留收文號，判斷下一勾選是否要清空
'****維護區*****
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號
Dim m_CP09 As String, m_idX As Integer  '收文號,Row
Dim m_CP31 As String, m_CP10 As String
Dim m_APP01 As String '申請人1
Dim m_CRC09 As String, m_CRC10 As String, m_CRC11 As String, m_CRC12 As String
Dim m_CRL55 As String, m_CRL67 As String, m_CRL68 As String, m_Attr As String
Dim m_CRL65 As String '關連編號=>相同案號(關連)群組顏色
Dim m_CKind As String '專利種類PA08, 商標種類TM08
Dim m_Na01 As String '申請國家
Dim m_CP140 As String '接洽單號
'*******************
Dim mDisplay As Integer, mDisCount As Integer  '案件展開=0 / 收合=1 ; 收合的案件數
Private Const cFixed As Integer = 10  '固定欄位: PUB_MGridGetId("案件性質", mGRID)+1
'GRD1的特定欄位之位置
Dim colCP140_1 As Integer, colCP09_1 As Integer, colCRC09_1 As Integer, colExp_1 As Integer
Dim colCAttr_1 As Integer, colGrpNO_1 As Integer, colCRL55_1 As Integer, colCRL67_1 As Integer, colCRL68_1 As Integer
Dim colCRL90_1 As Integer, colCRL55n_1 As Integer, colCRL67n_1 As Integer, colCRL68n_1 As Integer
Dim colCaseNo_1 As Integer, colCP31_1 As Integer, colCKind_1 As Integer, colNA01_1 As Integer
Dim colCP01_1 As Integer, colCP02_1 As Integer, colCP03_1 As Integer, colCP04_1 As Integer
Dim colCRC10_1 As Integer, colCRC11_1 As Integer, colCRC12_1 As Integer, colCP10_1 As Integer, colCP10n_1 As Integer
Dim colCRC09n_1 As Integer, colCAttrName_1 As Integer, colF0309_1 As Integer, colF0316_1 As Integer
Dim colAPP01_1 As Integer, colCaseName_1 As Integer
Dim bolUpdCP14 As Boolean 'Added by Lydia 2024/11/06 是否變更CP14
Dim colMsg1 As Integer 'Added by Lydia 2025/04/15
Dim colCP13_1 As String 'Added by Morgan 2025/7/3
Dim bolT31xT11 As Boolean, bolDutyT11 As Boolean 'Added by Lydia 2025/07/24
Dim m_Str專利處台北區主管 As String  'Added by Lydia 2025/08/18

Private Sub SetRoleData()
   
   m_ST06 = PUB_GetST06(strUserNum)
   m_SysNo = ""
   Combo1.Clear
   
 '開放分所主管可以切其他所 ---- 經理
'   If m_ST06 = "1" Then  '北所
'       Combo1.AddItem "ALL  全所"
'       Combo1.AddItem "1  北所+分所已分案"
'       Combo1.AddItem "2  中所"
'       Combo1.AddItem "3  南所"
'       Combo1.AddItem "4  高所"
'       Combo1.ListIndex = 1
'   Else
'        Select Case m_ST06
'            Case "1": Combo1.AddItem "1  北所及各所預分"
'            Case "2": Combo1.AddItem "2  中所"
'            Case "3": Combo1.AddItem "3  南所"
'            Case "4": Combo1.AddItem "4  高所"
'        End Select
'        Combo1.ListIndex = 0
'        Combo1.Locked = True
'   End If
'Memo by Lydia 2023/01/04 北所主管可以對全所＋各分所做「確定分案」，分所可以看他所案件但僅可以「更正」不可分案(from 副總)
   If Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "P1" Then
       '專利部主管
       Combo1.AddItem "ALL  全所"
       Combo1.AddItem "1  北所+分所已分案"
       Combo1.AddItem "2  中所"
       Combo1.AddItem "3  南所"
       Combo1.AddItem "4  高所"
       Combo1.ListIndex = Val(m_ST06)
   Else
       '商標部主管
       Combo1.AddItem "ALL  全所"
       Combo1.ListIndex = 0
   End If
   
   If Pub_StrUserSt03 = "M51" Then
        m_nowF0308 = "A6"
        m_SysNo = "P,T"
   Else
        If m_ST06 = "1" Then  '北所
            m_nowF0308 = "A6"
        Else
            m_nowF0308 = "A5"
        End If
        If Left(Pub_StrUserSt03, 2) = "P1" Then
           m_SysNo = "P"
           Call ShowObject(True)
        ElseIf Left(Pub_StrUserSt03, 2) = "P2" Then
           m_SysNo = "T"
           Label1(8).Visible = False:   txtCRC10.Visible = False
           Label1(11).Visible = False: Label1(3).Visible = False: Label1(4).Visible = False: Label1(5).Visible = False
           cboCRL55.Visible = False: cboAttr.Visible = False: cboCRL67.Visible = False: cboCRL68.Visible = False
           Label1(12).Visible = False: Label1(13).Visible = False: txtCRC11.Visible = False: txtCRC12.Visible = False
           Check1.Visible = False
        End If
        
   End If
   If m_SysNo = "T" Then
       Frame1.Height = 885
       Me.SSTab1.TabCaption(0) = "主管分案"
       Label2(0).Caption = "(註:雙擊開啟接洽單)"
   Else
       Frame1.Height = 1815
       Me.SSTab1.TabCaption(0) = "主管分案（複選不可執行更正和通知補件）"
       Label2(0).Caption = "(註:複選請先按Ctrl鍵，雙擊開啟接洽單)"
   End If
   
End Sub

Private Sub cboAttr_Validate(Cancel As Boolean)

   If cboAttr.Enabled = True And cboAttr <> "" And m_CP01 <> "" Then
      '專利設計案案件屬性
      If InStr(m_CP01, "P") > 0 And Left(cboAttr, 1) = "3" Then
        cboAttr = Left(cboAttr, 1) + "." + PUB_GetCaseAttributeName(Left(cboAttr, 1), m_CKind)
      Else
        cboAttr = Left(cboAttr, 1) + "." + PUB_GetCaseAttributeName(Left(cboAttr, 1))
      End If
      If cboAttr = Left(cboAttr, 1) + "." Then
         cboAttr = Left(cboAttr, 1)
         Cancel = True
         cboAttr.SetFocus
      End If
   End If
   cboAttr.Tag = cboAttr.Text
End Sub

Private Sub cboCRL55_Validate(Cancel As Boolean)
   If cboCRL55.Enabled = True And cboCRL55.Tag <> cboCRL55.Text Then
       If cboCRL55.Text <> "" And (m_CP09 = "" Or m_idX = 0) Then
          MsgBox "請先選取資料！", vbInformation
          GoTo EXITSUB
       End If
       If Check1.Visible = True And Check1.Value = 1 And cboCRL55.Text <> "" Then
          MsgBox "已設定取消相同案號！", vbInformation
          cboCRL55.Text = ""
          GoTo EXITSUB
       End If
       '案號依規則P-XXXXX
       strQuery = GetSameInput(cboCRL55.Text)
       If strQuery <> "" Then cboCRL55.Text = strQuery
   End If
   cboCRL55.Tag = cboCRL55.Text
   Exit Sub
   
EXITSUB:
   Cancel = True
   cboCRL55.SetFocus
End Sub

Private Sub cboCRL67_Validate(Cancel As Boolean)
   If cboCRL67.Enabled = True And cboCRL67.Tag <> cboCRL67.Text Then
       If cboCRL67.Text <> "" And (m_CP09 = "" Or m_idX = 0) Then
          MsgBox "請先選取資料！", vbInformation
          GoTo EXITSUB
       End If
       '案號依規則P-XXXXX
       strQuery = GetSameInput(cboCRL67.Text)
       If strQuery <> "" Then cboCRL67.Text = strQuery
   End If
   cboCRL67.Tag = cboCRL67.Text
   Exit Sub
   
EXITSUB:
   Cancel = True
   cboCRL67.SetFocus
End Sub

Private Sub cboCRL68_Validate(Cancel As Boolean)
   If cboCRL68.Enabled = True And cboCRL68.Tag <> cboCRL68.Text Then
       If cboCRL68.Text <> "" And (m_CP09 = "" Or m_idX = 0) Then
          MsgBox "請先選取資料！", vbInformation
          GoTo EXITSUB
       End If
       '案號依規則P-XXXXX
       strQuery = GetSameInput(cboCRL68.Text)
       If strQuery <> "" Then cboCRL68.Text = strQuery
   End If
   cboCRL68.Tag = cboCRL68.Text
   Exit Sub
   
EXITSUB:
   Cancel = True
   cboCRL68.SetFocus
End Sub

'讀取第一筆選取的資料
Private Function GetCurrRow(ByRef mGRID As MSHFlexGrid, ByRef pRow As Integer, ByVal pIdx As Integer) As Boolean
   Me.Enabled = False
   pRow = 0
   If mGRID.Rows > 1 Then
      For intQ = 1 To mGRID.Rows - 1
         mGRID.col = 0
         mGRID.row = intQ
         If mGRID.Text = "V" Then
             pRow = intQ
             GetCurrRow = True
             GoTo JumpToExit
         ElseIf mGRID.Text = "" Then
             If mGRID.CellBackColor = &HFFC0C0 Then
                mGRID.col = cFixed + 1 '取固定欄位後的底色，排除因為全選的變色
                lngColor = mGRID.CellBackColor
                For intP = 0 To cFixed
                   If Not (pIdx = 1 And InStr("," & colCRL55n_1 & "," & colCRL67n_1 & ",", intP) > 0) Then  '排除相同案號
                      mGRID.col = intP
                      mGRID.CellBackColor = lngColor
                   End If
                Next intP
             End If
         End If
      Next intQ
   End If
JumpToExit:
   Me.Enabled = True
End Function

Private Sub Check1_Click()
   If Check1.Value = 1 Then
       cboCRL55.Text = ""
   Else
       cboCRL55.Text = cboCRL55.Tag
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   blnDBClick = False
   Call PubShowNextData
End Sub

Public Sub PubShowNextData()
Dim intA As Integer, intB As Integer
Dim StrTag As String
Dim pGrid As MSHFlexGrid

   If cmdState >= 0 And cmdState <= 3 Then
       If SSTab1.Tab = 1 Then SSTab1.Tab = 0
       'Modified by Lydia 2023/01/17  在意見欄使用「向上鍵」造成SetFocus到其他欄位，暫存接洽單號m_CP140=""
       'If cmdState > 0 And mSelCount = 0 Then
       If cmdState > 0 And (mSelCount = 0 Or (InStr("2,3", cmdState) > 0 And m_CP140 = "")) Then
          MsgBox "請先選取資料！", vbInformation
          Exit Sub
       End If
      If mDisCount > 0 Then
          MsgBox "案件收合狀態不可執行！", vbInformation
          Exit Sub
      End If
   End If
   
   '關閉接洽單: 寫資料的動作
   If InStr("0,1,2,3,8,", cmdState & ",") > 0 Then
      If PUB_CheckFormExist("frm090801_Q") = True Then
          Unload frm090801_Q
      End If
   End If
   
   If SSTab1.Tab = 0 Then
      Set pGrid = GRD1
   Else
      Set pGrid = grd2
   End If
   Select Case cmdState
      Case 0 '全選/取消
         Call GetAllSelType(GRD1, IIf(InStr(cmdOK(cmdState).Caption, "取消") = 0, "1", "0"))
         If InStr(cmdOK(cmdState).Caption, "取消") = 0 Then
            cmdOK(cmdState).Caption = "取消全選"
         Else
            cmdOK(cmdState).Caption = "全選"
         End If
   
      Case 1 '確定分案
           '分所人員可以直接交由北所分案
           If m_nowF0308 = "A5" Then
                intB = 0
                For intA = 1 To pGrid.Rows - 1
                    pGrid.col = 0
                    pGrid.row = intA
                    If Trim(pGrid.Text) = "V" Then
                       If m_CP09 = "" & pGrid.TextMatrix(intA, colCP09_1) And cboCRC09.Text <> "" Then
                       ElseIf "" & pGrid.TextMatrix(intA, colCRC09_1) = "" Then
                          intB = intB + 1
                       End If
                    End If
                Next intA
                If intB > 0 Then
                   If MsgBox("有" & intB & "筆未分案，是否交由北所分案？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                       Exit Sub
                   End If
                End If
           End If
           cmdOK(cmdState).Enabled = False
           For intA = 1 To pGrid.Rows - 1
               pGrid.col = 0
               pGrid.row = intA
               If Trim(pGrid.Text) = "V" Then
                   pGrid.col = cFixed + 1 '取固定欄位後的底色，排除因為全選的變色
                   lngColor = pGrid.CellBackColor
                  '因為分所人員可以直接交由北所分案，所以不預分承辦人也可分案
                  If "" & pGrid.TextMatrix(intA, colCP01_1) <> "" Then
                     If TxtValidate(True) = True Then
                        Call ChkT31xT11(pGrid.TextMatrix(intA, colCP140_1)) 'Added by Lydia 2025/07/24
                        If SaveToCP14(intA, pGrid.TextMatrix(intA, colCP140_1), pGrid.TextMatrix(intA, colCP09_1), pGrid.TextMatrix(intA, colF0316_1), pGrid.TextMatrix(intA, colF0309_1)) = True Then
                            pGrid.TextMatrix(intA, 0) = ""
                            For intB = 0 To cFixed
                               If InStr("," & colCRL55n_1 & "," & colCRL67n_1 & ",", intB) = 0 Then  '排除相同案號
                                   pGrid.col = intB
                                   pGrid.CellBackColor = lngColor
                               End If
                            Next intB
                        Else
                            Call QueryData
                            GoTo JumpToExit
                        End If
                     End If
                  End If
               End If
           Next intA
           Call QueryData
JumpToExit:
           cmdOK(cmdState).Enabled = True
      Case 2 '智權補件
         If SaveToBack(cmdState) = True Then
             Call QueryData
         End If
      Case 3 '程序補件
         If SaveToBack(cmdState) = True Then
             Call QueryData
         End If
      Case 4 '完整卷宗
           Me.Enabled = False
           For intA = 1 To pGrid.Rows - 1
               pGrid.col = 0
               pGrid.row = intA
               If Trim(pGrid.Text) = "V" Then
                  pGrid.col = 0
                  'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                  'pGrid.Text = ""
                  'If SSTab1.Tab = 0 Then Call GetSelState(intA)
                  'end 2023/01/05
                  If "" & pGrid.TextMatrix(intA, colCP01_1) <> "" Then
                     StrTag = pGrid.TextMatrix(intA, colCP01_1) & "-" & pGrid.TextMatrix(intA, colCP02_1) & "-" & pGrid.TextMatrix(intA, colCP03_1) & "-" & pGrid.TextMatrix(intA, colCP04_1)
                     Screen.MousePointer = vbHourglass
                     frm100101_L.m_strKey = StrTag
                     frm100101_L.SetParent Me
                     If frm100101_L.QueryData = True Then
                        frm100101_L.Show
                        Me.Hide
                     End If
                     Screen.MousePointer = vbDefault
                     'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                     'Call GetCurrRow(pGrid, m_idX, IIf(SSTab1.Tab = 0, 1, 2))
                     'If m_idX > 0 Then
                     '    Call UpdateCtrlData(m_idX)
                     'Else
                     '    Call ClearCtrlData
                     'End If
                     'end 2023/01/05
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
           Next intA
           Me.Enabled = True
      Case 5 '基本資料
           Me.Enabled = False
           For intA = 1 To pGrid.Rows - 1
               pGrid.col = 0
               pGrid.row = intA
               If Trim(pGrid.Text) = "V" Then
                  pGrid.col = 0
                  'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                  'pGrid.Text = ""
                  'If SSTab1.Tab = 0 Then Call GetSelState(intA)
                  'end 2023/01/05
                  If "" & pGrid.TextMatrix(intA, colCP01_1) <> "" Then
                     If fnSaveParentForm(Me) = False Then
                         Me.Enabled = True
                         Exit Sub
                     End If
                     StrTag = pGrid.TextMatrix(intA, colCP01_1) & "-" & pGrid.TextMatrix(intA, colCP02_1) & "-" & pGrid.TextMatrix(intA, colCP03_1) & "-" & pGrid.TextMatrix(intA, colCP04_1)
                     Select Case Trim("" & pGrid.TextMatrix(intA, colCP01_1))
                         Case "CFP", "FCP", "P"   '專利
                               Screen.MousePointer = vbHourglass
                               frm100101_3.Show
                               frm100101_3.Tag = StrTag
                               frm100101_3.StrMenu
                               Screen.MousePointer = vbDefault
                         Case "CFT", "FCT", "T", "TF"   '商標
                               Screen.MousePointer = vbHourglass
                               frm100101_4.Show
                               frm100101_4.Tag = StrTag
                               frm100101_4.StrMenu
                               Screen.MousePointer = vbDefault
                         Case "CFL", "FCL", "L", "LIN", "ACS"    '法務
                               Screen.MousePointer = vbHourglass
                               frm100101_5.Show
                               frm100101_5.Tag = StrTag
                               frm100101_5.StrMenu
                               Screen.MousePointer = vbDefault
                         Case "LA"            '顧問
                               Screen.MousePointer = vbHourglass
                               frm100101_6.Show
                               frm100101_6.Tag = StrTag
                               frm100101_6.StrMenu
                               Screen.MousePointer = vbDefault
                         Case Else                  '服務
                              Select Case Trim("" & pGrid.TextMatrix(intA, colCP01_1))
                                  Case "TB"    '條碼
                                     Screen.MousePointer = vbHourglass
                                     frm100101_7.Show
                                     frm100101_7.Tag = StrTag
                                     frm100101_7.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case "TM"
                                     Screen.MousePointer = vbHourglass
                                     frm100101_8.Show
                                     frm100101_8.Tag = StrTag
                                     frm100101_8.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case "TD"
                                     Screen.MousePointer = vbHourglass
                                     frm100101_9.Show
                                     frm100101_9.Tag = StrTag
                                     frm100101_9.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case "TC", "CFC"
                                     Screen.MousePointer = vbHourglass
                                     frm100101_A.Show
                                     frm100101_A.Tag = StrTag
                                     frm100101_A.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case Else
                                     Screen.MousePointer = vbHourglass
                                     frm100101_B.Show
                                     frm100101_B.Tag = StrTag
                                     frm100101_B.StrMenu
                                     Screen.MousePointer = vbDefault
                               End Select
                     End Select
                     'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                     'Call GetCurrRow(pGrid, m_idX, IIf(SSTab1.Tab = 0, 1, 2))
                     'If m_idX > 0 Then
                     '    Call UpdateCtrlData(m_idX)
                     'Else
                     '    Call ClearCtrlData
                     'End If
                     'end 2023/01/05
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
           Next intA
           Me.Enabled = True
      Case 6 '進度檔
           Me.Enabled = False
           For intA = 1 To pGrid.Rows - 1
               pGrid.col = 0
               pGrid.row = intA
               If Trim(pGrid.Text) = "V" Then
                  pGrid.col = 0
                  'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                  'pGrid.Text = ""
                  'If SSTab1.Tab = 0 Then Call GetSelState(intA)
                  'end 2023/01/05
                  If "" & pGrid.TextMatrix(intA, colCP01_1) <> "" Then
                     If fnSaveParentForm(Me) = False Then
                         Me.Enabled = True
                         Exit Sub
                     End If
                     
                     StrTag = pGrid.TextMatrix(intA, colCP01_1) & "-" & pGrid.TextMatrix(intA, colCP02_1) & "-" & pGrid.TextMatrix(intA, colCP03_1) & "-" & pGrid.TextMatrix(intA, colCP04_1)
                     Screen.MousePointer = vbHourglass
                     'Mark by Lydia 2023/01/03 需要看完整進度
                     'frm100101_C.Show
                     'frm100101_C.Tag = StrTag & "=" & pGrid.TextMatrix(intA, colCP09_1)
                     'frm100101_C.StrMenu
                     frm100101_2.Show
                     frm100101_2.Tag = StrTag
                     frm100101_2.StrMenu
                     'end 2023/01/03
                     Screen.MousePointer = vbDefault
                     'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                     'Call GetCurrRow(pGrid, m_idX, IIf(SSTab1.Tab = 0, 1, 2))
                     'If m_idX > 0 Then
                     '    Call UpdateCtrlData(m_idX)
                     'Else
                     '    Call ClearCtrlData
                     'End If
                     'end 2023/01/05
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
           Next intA
           Me.Enabled = True
      
      Case 7 '專利相關案
           Me.Enabled = False
           For intA = 1 To pGrid.Rows - 1
               pGrid.col = 0
               pGrid.row = intA
               If Trim(pGrid.Text) = "V" Then
                  pGrid.col = 0
                  'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                  'pGrid.Text = ""
                  'If SSTab1.Tab = 0 Then Call GetSelState(intA)
                  'end 2023/01/05
                  If "" & pGrid.TextMatrix(intA, colCP01_1) <> "" Then
                     If fnSaveParentForm(Me) = False Then
                         Me.Enabled = True
                         Exit Sub
                     End If
                     StrTag = pGrid.TextMatrix(intA, colCP01_1) & "-" & pGrid.TextMatrix(intA, colCP02_1) & "-" & pGrid.TextMatrix(intA, colCP03_1) & "-" & pGrid.TextMatrix(intA, colCP04_1)
                     Screen.MousePointer = vbHourglass
                     frm100101_h.Show
                     frm100101_h.KeyString = StrTag
                     frm100101_h.SearchKind = "本所案號"
                     frm100101_h.StrMenu
                     Screen.MousePointer = vbDefault
                     'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                     'Call GetCurrRow(pGrid, m_idX, IIf(SSTab1.Tab = 0, 1, 2))
                     'If m_idX > 0 Then
                     '    Call UpdateCtrlData(m_idX)
                     'Else
                     '    Call ClearCtrlData
                     'End If
                     'end 2023/01/05
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
           Next intA
           Me.Enabled = True
      Case 8 '結束
           Unload Me
      Case 9 '接洽單
           Me.Enabled = False
           For intA = 1 To pGrid.Rows - 1
               pGrid.col = 0
               pGrid.row = intA
               If Trim(pGrid.Text) = "V" Then
                  pGrid.col = 0
                  'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                  'pGrid.Text = ""
                  'If SSTab1.Tab = 0 Then Call GetSelState(intA)
                  'end 2023/01/05
                  If "" & pGrid.TextMatrix(intA, colCP140_1) <> "" Then
                     Screen.MousePointer = vbHourglass
                     Call ShowFormCRL(pGrid.TextMatrix(intA, colCP140_1), pGrid.TextMatrix(intA, colCP09_1))
                     Screen.MousePointer = vbDefault
                     'Call GetCurrRow(pGrid, m_idX, IIf(SSTab1.Tab = 0, 1, 2)) 'Mark by Lydia 2023/01/05 不取消勾選 from 李柏翰
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
           Next intA
           Me.Enabled = True
      Case Else
   End Select
End Sub

Private Sub SetGrid(ByRef mGRID As MSHFlexGrid, ByVal pIdx As Integer, Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, intX As Integer
   Dim strTmpGrp As String, intGrp As Integer, strTmpCRL67 As String, intCRL67 As Integer
   Dim intY As Integer
   
'更新欄位在前
'   arrGridHeadText = Array("V", "展", "CRL65", "GRPNO", "承辦人", "CRL55", "相同案號", "本所案號", "案 件 名 稱", "案件性質", "申請國", "智權人", "本所期限", "申  請  人", _
'                                      "屬性", "CRL67", "一案兩請", "CRL68", "擬制喪失新穎性", "算案件", "計件值", "加乘註", "法定期限", "表單狀態", "CP01", "CP02", "CP03", "CP04", _
'                                      "CP31", "CKIND", "CATTR", "F0309", "F0316", "CP140", "CP09", "CP10", "CRC09", "CP12", "CP13", "NA01", "APP01", "CRL90")
'
'   If m_SysNo = "T" Then '商標：隱藏"相同案號", "案件屬性", "一案兩請", "擬制喪失新穎性"
'      arrGridHeadWidth = Array(240, 240, 0, 0, 720, 0, 0, 1000, 1200, 1200, 800, 800, 860, 1200, _
'                                      0, 0, 0, 0, 0, 0, 0, 0, 860, 860, 0, 0, 0, 0, _
'                                      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
'   Else
'      arrGridHeadWidth = Array(240, 240, 0, 0, 720, 0, 1000, 1000, 1200, 1200, 800, 800, 860, 1200, _
'                                    800, 0, 1000, 0, 1000, 700, 700, 700, 860, 860, 0, 0, 0, 0, _
'                                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
'   End If
   'Modified by Lydia 2024/04/18 +sord
   'Modified by Lydia 2025/04/15 +MSG1
   arrGridHeadText = Array("V", "展", "GRPNO", "承辦人", "CRL55", "相同案號", "一案兩請", "本所案號", "案 件 名 稱", "案件性質", "申請國", "智權人", "本所期限", "申  請  人", _
                                      "屬性", "CRL67", "CRL68", "擬制喪失新穎性", "算案件", "計件值", "加乘註", "法定期限", "表單狀態", "CP01", "CP02", "CP03", "CP04", _
                                      "CP31", "CKIND", "CATTR", "F0309", "F0316", "CP140", "CP09", "CP10", "CRC09", "CP12", "CP13", "NA01", "APP01", "CRL90", "SORD", "MSG1")
   If m_SysNo = "T" Then '商標：隱藏"相同案號", "案件屬性", "一案兩請", "擬制喪失新穎"
      arrGridHeadWidth = Array(240, 240, 0, 720, 0, 0, 0, 1000, 1200, 1200, 800, 800, 720, 1200, _
                                      0, 0, 0, 0, 0, 0, 0, 860, 860, 0, 0, 0, 0, _
                                      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   Else
      arrGridHeadWidth = Array(240, 240, 0, 720, 0, 1160, 1160, 1000, 1200, 460, 460, 720, 860, 1200, _
                                    800, 0, 0, 1000, 700, 700, 700, 860, 860, 0, 0, 0, 0, _
                                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   End If
   mGRID.Visible = False
   mGRID.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         mGRID.Clear
         mGRID.Rows = 2
   End If
       
    For iRow = 0 To mGRID.Cols - 1
       mGRID.row = 0
       mGRID.col = iRow
       mGRID.Text = arrGridHeadText(iRow)
       mGRID.CellAlignment = flexAlignCenterCenter
       mGRID.ColWidth(iRow) = arrGridHeadWidth(iRow)
    Next

   'mgrid的特定欄位之位置
   If colCP140_1 = 0 Then
      colCP140_1 = PUB_MGridGetId("CP140", mGRID)   '接洽單號
      colCP09_1 = PUB_MGridGetId("CP09", mGRID)  '收文號
      colCRC09_1 = PUB_MGridGetId("CRC09", mGRID)   '預分承辦人ID
      colCRC09n_1 = PUB_MGridGetId("承辦人", mGRID)   '預分承辦人
      colCAttr_1 = PUB_MGridGetId("CATTR", mGRID)   '案件屬性 PA158
      colCAttrName_1 = PUB_MGridGetId("屬性", mGRID)   '案件屬性
      colGrpNO_1 = PUB_MGridGetId("GRPNO", mGRID)   '群組顏色
      colCRL55n_1 = PUB_MGridGetId("相同案號", mGRID)  '相同案號
      colCRL55_1 = PUB_MGridGetId("CRL55", mGRID)
      colCRL67n_1 = PUB_MGridGetId("一案兩請", mGRID)  '一案兩請
      colCRL67_1 = PUB_MGridGetId("CRL67", mGRID)
      colCRL68n_1 = PUB_MGridGetId("擬制喪失新穎性", mGRID)  '擬制喪失新穎性
      colCRL68_1 = PUB_MGridGetId("CRL68", mGRID)
      colCRL90_1 = PUB_MGridGetId("CRL90", mGRID) '急件
      
      colCaseNo_1 = PUB_MGridGetId("本所案號", mGRID)   '本所案號
      colCP01_1 = PUB_MGridGetId("CP01", mGRID)
      colCP02_1 = PUB_MGridGetId("CP02", mGRID)
      colCP03_1 = PUB_MGridGetId("CP03", mGRID)
      colCP04_1 = PUB_MGridGetId("CP04", mGRID)
      colCP10_1 = PUB_MGridGetId("CP10", mGRID)
      colCaseName_1 = PUB_MGridGetId("案 件 名 稱", mGRID)
      colCP10n_1 = PUB_MGridGetId("案件性質", mGRID)
      colCRC10_1 = PUB_MGridGetId("算案件", mGRID) '是否算案件數CRC10=>CP26
      colCRC11_1 = PUB_MGridGetId("計件值", mGRID) '承辦人計件值CRC11=>CP97
      colCRC12_1 = PUB_MGridGetId("加乘註", mGRID) '承辦人加乘註記CRC12=>CP98
      colCP31_1 = PUB_MGridGetId("CP31", mGRID)
      colCKind_1 = PUB_MGridGetId("CKIND", mGRID) '案件種類PA08,TM08
      colNA01_1 = PUB_MGridGetId("NA01", mGRID)
      colAPP01_1 = PUB_MGridGetId("APP01", mGRID)  '申請人1
      colExp_1 = 1
      colF0316_1 = PUB_MGridGetId("F0316", mGRID)
      colF0309_1 = PUB_MGridGetId("F0309", mGRID)
      colCP13_1 = PUB_MGridGetId("CP13", mGRID) 'Added by Morgan 2025/7/3
      colMsg1 = PUB_MGridGetId("MSG1", mGRID) 'Added by Lydia 2025/04/15 是否為規費調整接洽單，規費有文字就不算CRC13=Y
   End If
   
   If pReset = False Then
      For intX = 1 To mGRID.Rows - 1
           mGRID.row = intX
           For iRow = 0 To mGRID.Cols - 1
              mGRID.col = iRow
              '急件整列紅色
              If "" & mGRID.TextMatrix(intX, colCRL90_1) = "Y" And iRow <> colCRL55n_1 And iRow <> colCRL67n_1 Then
                  mGRID.CellBackColor = &H8080FF  '&H8080FF
              Else
                  mGRID.CellBackColor = &H80000005 '白底
              End If
              '群組顏色: 相同案號CRL55,簡化為兩色
              If iRow = colCRL55n_1 Then
                 If "" & mGRID.TextMatrix(intX, colCRL55_1) <> "" Then
                     If strTmpGrp <> "" & mGRID.TextMatrix(intX, colCRL55_1) Then
                        If strTmpGrp <> "" Then
                            intGrp = intGrp + 1
                        End If
                     End If
                     If intGrp Mod 2 = 0 Then
                         mGRID.CellBackColor = PGColor(0)
                     Else
                         mGRID.CellBackColor = PGColor(1)
                     End If
                     mGRID.CellFontBold = True
                     strTmpGrp = "" & mGRID.TextMatrix(intX, colCRL55_1)
                 ElseIf "" & mGRID.TextMatrix(intX, colCRL90_1) <> "Y" Then
                     mGRID.CellBackColor = &H80000005 '白底
                 End If
              End If
              '群組顏色: 一案兩請CRL67,簡化為兩色
              If iRow = colCRL67n_1 Then
                 If "" & mGRID.TextMatrix(intX, colCRL67_1) <> "" Then
                     If strTmpCRL67 <> "" & mGRID.TextMatrix(intX, colCRL67_1) Then
                        If strTmpCRL67 <> "" Then
                            intCRL67 = intCRL67 + 1
                        End If
                     End If
                     If intCRL67 Mod 2 = 0 Then
                         mGRID.CellBackColor = PGColor(2)
                     Else
                         mGRID.CellBackColor = PGColor(3)
                     End If
                     mGRID.CellFontBold = True
                     strTmpCRL67 = "" & mGRID.TextMatrix(intX, colCRL67_1)
                     '判斷相同案號=一案兩請, 不顯示
                     If "" & mGRID.TextMatrix(intX, colCRL55_1) = "" & mGRID.TextMatrix(intX, colCRL67_1) And "" & mGRID.TextMatrix(intX, colCRL67_1) <> "" Then
                          mGRID.TextMatrix(intX, colCRL55n_1) = ""
                     End If
                 ElseIf "" & mGRID.TextMatrix(intX, colCRL90_1) <> "Y" Then
                     mGRID.CellBackColor = &H80000005 '白底
                 End If
              End If
              mGRID.CellAlignment = flexAlignLeftCenter '內容靠左
           Next iRow
           '專利案: 保持更正的勾選項
           If m_SelCP09 <> "" And InStr(m_SelCP09, "," & mGRID.TextMatrix(intX, colCP09_1)) > 0 Then
               mGRID.TextMatrix(intX, 0) = "V"
               mSelCount = mSelCount + 1
               If intLastTop <= 1 Then intLastTop = intX
               For intY = 0 To cFixed - 1
                   If InStr("," & colCRL55n_1 & "," & colCRL67n_1 & ",", intY) = 0 Then '排除相同案號
                      mGRID.col = intY
                      mGRID.CellBackColor = &HFFC0C0
                   End If
               Next intY
           End If
           
      Next intX
   End If
   
   mGRID.Visible = True
   
End Sub

Public Function QueryData() As Boolean
Dim strMid(1 To 4) As String
   
   If PUB_CheckFormExist("frm090801_Q") = True Then
        Unload frm090801_Q
   End If
   Call ClearCtrlData
   Call CtrlOKbutton(True)
   Frame1.Enabled = False '有勾選一筆才可以輸入維護區的資料，為了避免誤輸，若未勾選或複選直接關閉輸入。
   QueryData = False
'***************************
'--相同案號(與總號CRL55)、一案兩請(案號CRL67)、擬制喪失新穎性(案號CRL68)
'--更正先寫入接洽記錄單案件性質ConsultRecCMP.CRC09，直到確定分案才會寫入CP14。
'***************************
   
    If m_SysNo = "" Then
        MsgBox "無權限！", vbCritical
        Exit Function
    End If
'更新欄位在前
'-----
    For intQ = 1 To 2
        strMid(3) = ""
        If intQ = 1 Then '分案
           strMid(3) = strMid(3) & "and F0308=f0202 and f0207 is null and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "')"
           If Left(Combo1, 1) = "A" Then
               'Modified by Lydia 2023/01/04 北所主管可以對全所＋各分所做「確定分案」，分所可以看他所案件但僅可以「更正」不可分案(from 副總)
               'strMid(3) = strMid(3) & IIf(m_nowF0308 = "A6", " and F0308 in ('A5','A6') ", " and F0308='" & m_nowF0308 & "' ")
               strMid(3) = strMid(3) & " and F0308 in ('A5','A6') "
           Else
               If Left(Combo1, 1) = "1" Then
                   strMid(3) = strMid(3) & " and F0308='A6' "
               Else
                   strMid(3) = strMid(3) & " and F0308='A5' "
               End If
           End If
        Else   '退智權 ＆ 退程序 (補件)
           '因為退回智權後，F0307=智權人員;  F0207=簽核結果6,7
           strMid(3) = strMid(3) & "and f0207 in ('6','7') and f0309 in ('" & Flow_智權補件 & "','" & Flow_程序補件 & "') and instr('A5,A7,'||F0316,F0308) > 0 "
           If Left(Combo1, 1) = "A" Then
               'Modified by Lydia 2023/01/04 北所主管可以對全所＋各分所做「確定分案」，分所可以看他所案件但僅可以「更正」不可分案(from 副總)
               'strMid(3) = strMid(3) & IIf(m_nowF0308 = "A6", " and F0307 in ('A5','A6') ", " and F0307='" & m_nowF0308 & "' ")
               strMid(3) = strMid(3) & " and F0307 in ('A5','A6') "
           Else
               If Left(Combo1, 1) = "1" Then
                   strMid(3) = strMid(3) & " and F0307='A6' "
               Else
                   strMid(3) = strMid(3) & " and F0307='A5' "
               End If
           End If
        End If
 
        strMid(intQ) = ""
        If InStr(m_SysNo, "P") > 0 Then
            '專利:P,CFP
            '相同案號(關連)群組顏色=GRPNO；111/12/13 改成直接用相同案號CRL55設定顏色
            '因為主管只管相同案CRL56=Y，若CRL56=N為ＸＸ案有關預設顯示空白; P.S.CRL74: 排除案源單號(流水號)、CFP英國脫歐案...
            '相同案號CRL55,一案兩請CRL66,擬制喪失新穎性CRL68: 因為子案記錄母案案號，而母案則記錄本身案號所以顯示為空白；為了分辨有無設定，另外抓資料
            'Modified by Lydia 2024/04/18 +sord
            'Modified by Lydia 2025/04/15 +MSG1
            strMid(intQ) = strMid(intQ) & "UNION SELECT '' AS V,'' AS 展開, CRL65 AS GRPNO,S1.ST02 AS CRC09N," & _
                             "DECODE(CRL74||CRL56,'Y',CRL55,'') CRL55, DECODE(CRL74||CRL56,'Y',DECODE(CRL55,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04),NULL,CRL55),NULL) AS 相同案號," & _
                             "DECODE(CRL67,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04),NULL,CRL67) AS 一案兩請,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04) AS CASENO,NVL(PA05,NVL(PA06,PA07)) AS CASENAME,DECODE(NA01,'000',CPM03,CPM04) AS CP10N, " & _
                             "NA03,S2.ST02 AS CP13N, SQLDATET(CP06) AS CP06T,NVL(CU04,NVL(CU05,CU06)) APP01N, DECODE(PA08,'3', DECODE(PA158,'1','整體','2','部分','3','圖像','4','組成',PA158 ),DECODE(PA158,'1','機械','2','電子','3','化學生醫',PA158 )) AS 案件屬性," & _
                             "CRL67,CRL68, DECODE(CRL68,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04),NULL,CRL68) AS 擬制喪失新穎性," & _
                             "NVL(CP26,CRC10) CRC10,CRC11,CRC12, SQLDATET(CP07) AS CP07T, DECODE(F0309, " & ShowFlow表單狀態中文 & ", F0309) AS 目前表單狀態," & _
                             "CP01,CP02,CP03,CP04,CP31,PA08 AS CKIND,PA158 AS CATTR,F0309,F0316 ,CP140,CP09,CP10,CRC09,CP12,CP13, NA01,PA26 AS APP01,CRL90 " & _
                             ", decode(crl65,null,null,crl65||nvl(crl67,cp01||'-'||cp02||decode(cp03||cp04,'000',null,'-'||cp03||'-'||cp04))||nvl(crl55,cp01||'-'||cp02||decode(cp03||cp04,'000',null,'-'||cp03||'-'||cp04))) as sord " & _
                             ", DECODE(CRC13||DECODE(TRANSLATE(CRC05,'/0123456789','/'),NULL,'','N'),'Y','Y',NULL) MSG1 " & _
                             "FROM CASEPROGRESS, CONSULTRECORDLIST, CONSULTRECCMP, STAFF S1, STAFF S2 ,PATENT, CUSTOMER, NATION,CASEPROPERTYMAP, FLOW003, FLOW002,STAFF S3 "
            '高所的機械案歸南所分案
            'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉AND NVL(CPM35,'0')<>'2'
            strMid(intQ) = strMid(intQ) & "WHERE CP140 IS NOT NULL AND CP158=0 AND CP159=0 AND CP157 IS NULL AND PA01 IN ('P','CFP') AND CP140=CRL01(+) " & _
                             " AND CRL01=CRC01 AND CRC08=CP09 AND CRC09=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                             "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND PA09=NA01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                             "AND CP140=F0301 AND F0302='" & m_F0302 & "' AND CP140=F0201 AND F0316=S3.ST01(+) " & strMid(3)
            '高所的機械案歸南所分案
            If Left(Combo1, 1) = "3" Then
               'Modified by Lydia 2024/06/03 測試銷卷/閉卷案件，遇見pa158=null無法正確判斷；PA158='1'=> NVL(PA158,'N')='1'
               strMid(intQ) = strMid(intQ) & " AND ((S3.ST06='3') or (S3.ST06='4' AND PA08 in ('1','2') AND NVL(PA158,'N')='1' )) "
            ElseIf Left(Combo1, 1) = "4" Then
               'Modified by Lydia 2024/06/03 測試銷卷/閉卷案件，遇見pa158=null無法正確判斷；PA158='1'=> NVL(PA158,'N')='1'
               strMid(intQ) = strMid(intQ) & " AND S3.ST06='4' AND NOT (PA08 in ('1','2') AND NVL(PA158,'N')='1' ) "
            ElseIf Left(Combo1, 1) = "2" Then
               strMid(intQ) = strMid(intQ) & " AND S3.ST06='" & Left(Combo1.Text, 1) & "'"
            End If
            '服務:PS,CPS
            'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉AND NVL(CPM35,'0')<>'2'
            'Modified by Lydia 2024/04/18 +sord
            'Modified by Lydia 2025/04/15 +MSG1
            strMid(intQ) = strMid(intQ) & "UNION SELECT '' AS V,'' AS 展開, CRL65 AS GRPNO,S1.ST02 AS CRC09N," & _
                             "DECODE(CRL74||CRL56,'Y',CRL55,'') CRL55, DECODE(CRL74||CRL56,'Y',DECODE(CRL55,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04),NULL,CRL55),NULL) AS 相同案號," & _
                             "DECODE(CRL67,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04),NULL,CRL67) AS 一案兩請,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04) AS CASENO,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,DECODE(NA01,'000',CPM03,CPM04) AS CP10N, " & _
                             "NA03,S2.ST02 AS CP13N, SQLDATET(CP06) AS CP06T,NVL(CU04,NVL(CU05,CU06)) APP01N, '' AS 案件屬性," & _
                             "CRL67, CRL68, DECODE(CRL68,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04),NULL,CRL68) AS 擬制喪失新穎性," & _
                             "NVL(CP26,CRC10) CRC10,CRC11,CRC12, SQLDATET(CP07) AS CP07T, DECODE(F0309, " & ShowFlow表單狀態中文 & ", F0309) AS 目前表單狀態," & _
                             "CP01,CP02,CP03,CP04,CP31,'' AS CKIND,'' AS CATTR,F0309,F0316 ,CP140,CP09,CP10,CRC09,CP12,CP13, NA01,SP08 AS APP01,CRL90 " & _
                             ", decode(crl65,null,null,crl65||nvl(crl67,cp01||'-'||cp02||decode(cp03||cp04,'000',null,'-'||cp03||'-'||cp04))||nvl(crl55,cp01||'-'||cp02||decode(cp03||cp04,'000',null,'-'||cp03||'-'||cp04))) as sord " & _
                             ", '' AS MSG1 FROM CASEPROGRESS, CONSULTRECORDLIST, CONSULTRECCMP, STAFF S1, STAFF S2 ,SERVICEPRACTICE, CUSTOMER, NATION,CASEPROPERTYMAP, FLOW003, FLOW002,STAFF S3 "
            strMid(intQ) = strMid(intQ) & "WHERE CP140 IS NOT NULL AND CP158=0 AND CP159=0 AND CP157 IS NULL AND SP01 IN ('PS','CPS') AND CP140=CRL01(+) " & _
                             " AND CRL01=CRC01 AND CRC08=CP09 AND CRC09=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
                             "AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND SP09=NA01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                             "AND CP140=F0301 AND F0302='" & m_F0302 & "' AND CP140=F0201 AND F0316=S3.ST01(+) " & strMid(3) & _
                             IIf(InStr("A,1,", Left(Combo1, 1)) = 0, " AND S3.ST06='" & Left(Combo1.Text, 1) & "' ", "")
        End If

        If InStr(m_SysNo, "T") > 0 Then
             '商標: T案和FCT爭議案，有Flow的商標案
            'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉AND NVL(CPM35,'0')<>'2'
            'Modified by Lydia 2024/04/18 +sord
            'Modified by Lydia 2025/04/15 +MSG1
            strMid(intQ) = strMid(intQ) & "UNION SELECT '' AS V,'' AS 展開,CRL65 AS GRPNO,S1.ST02 AS CRC09N," & _
                                "DECODE(CRL74||CRL56,'Y',CRL55,'') CRL55,'' AS 相同案號,'' AS 一案兩請,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04) AS CASENO,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(NA01,'000',CPM03,CPM04) AS CP10N," & _
                                "NA03,S2.ST02 AS CP13N, SQLDATET(CP06) AS CP06T,NVL(CU04,NVL(CU05,CU06)) APP01N, DECODE(TM08,'1','商標','7','證明標章','8','團體標章','9','團體商標',TM08) AS 商標種類 ,CRL67, CRL68, '' AS 擬制喪失新穎性," & _
                                "NVL(CP26,CRC10) CRC10,CRC11,CRC12, SQLDATET(CP07) AS CP07T, DECODE(F0309, " & ShowFlow表單狀態中文 & ", F0309) AS 目前表單狀態," & _
                                "CP01,CP02,CP03,CP04,CP31,'' AS CKIND,'' AS CATTR,F0309,F0316 ,CP140,CP09,CP10,CRC09,CP12,CP13, NA01,TM23 AS APP01,CRL90,'0' as sord " & _
                                ", '' AS MSG1 FROM CASEPROGRESS, ConsultRecordList, ConsultRecCMP, STAFF S1, STAFF S2 ,TRADEMARK, CUSTOMER, NATION,CasePropertyMap, FLOW003, FLOW002,STAFF S3 "
            strMid(intQ) = strMid(intQ) & "WHERE CP140 IS NOT NULL AND CP158=0 AND CP159=0 AND CP157 IS NULL AND TM01 IS NOT NULL AND CP140=CRL01(+) AND CRL01=CRC01 " & _
                                "AND CRC08=CP09 AND CRC09=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
                                "AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND TM10=NA01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                                "AND CP140=F0301 AND F0302='" & m_F0302 & "' AND CP140=F0201 AND F0316=S3.ST01(+) " & strMid(3)
            '服務案: SK02=6
            'Modified by Lydia 2024/04/18 +sord
            'Modified by Lydia 2025/04/15 +MSG1
            strMid(intQ) = strMid(intQ) & "UNION SELECT '' AS V,'' AS 展開,CRL65 AS GRPNO,S1.ST02 AS CRC09N," & _
                                "DECODE(CRL74||CRL56,'Y',CRL55,'') CRL55,'' AS 相同案號,'' AS 一案兩請,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04) AS CASENO,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,DECODE(NA01,'000',CPM03,CPM04) AS CP10N," & _
                                "NA03,S2.ST02 AS CP13N, SQLDATET(CP06) AS CP06T,NVL(CU04,NVL(CU05,CU06)) APP01N, '' AS 商標種類 ,CRL67, CRL68, '' AS 擬制喪失新穎性," & _
                                "NVL(CP26,CRC10) CRC10,CRC11,CRC12, SQLDATET(CP07) AS CP07T, DECODE(F0309, " & ShowFlow表單狀態中文 & ", F0309) AS 目前表單狀態," & _
                                "CP01,CP02,CP03,CP04,CP31,'' AS CKIND,'' AS CATTR,F0309,F0316 ,CP140,CP09,CP10,CRC09,CP12,CP13, NA01,SP08 AS APP01,CRL90,'0' as sord " & _
                                ", '' AS MSG1 FROM CASEPROGRESS, ConsultRecordList, ConsultRecCMP, STAFF S1, STAFF S2 ,SERVICEPRACTICE,SYSTEMKIND, CUSTOMER, NATION,CasePropertyMap, FLOW003, FLOW002,STAFF S3 "
            'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉AND NVL(CPM35,'0')<>'2'
            strMid(intQ) = strMid(intQ) & "WHERE CP140 IS NOT NULL AND CP158=0 AND CP159=0 AND CP157 IS NULL AND CP01=SK01 AND SK02='6' AND SP01 IS NOT NULL AND CP140=CRL01(+) AND CRL01=CRC01 " & _
                                "AND CRC08=CP09 AND CRC09=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
                                "AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND SP09=NA01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                                "AND CP140=F0301 AND F0302='" & m_F0302 & "' AND CP140=F0201 AND F0316=S3.ST01(+) " & strMid(3)
        End If
        If InStr(m_SysNo, "P") > 0 Then '專利
            '排序:改成依相同案號>一案兩請>擬制喪失新穎性>接洽單號
            'strMid(intQ) = strMid(intQ) & "ORDER BY GRPNO,CP12,CP13, CASENO,CP09 "
            'Modified by Lydia 2024/04/18 同一組(相同案號+一案兩請+擬制喪失新穎性)的所有案件一定要全部在一起; 以關聯代號為主做排序
            'strMid(intQ) = strMid(intQ) & "ORDER BY CRL55 ASC, CRL67 ASC, CRL68 ASC, CP140 ASC ,CP09 ASC "
            strMid(intQ) = strMid(intQ) & "ORDER BY sord ASC, CP140 ASC ,CP09 ASC "
        Else
            strMid(intQ) = strMid(intQ) & "ORDER BY NA01,CASENO,CP09"
        End If
    Next intQ

    cmdOK(0).Caption = "全選"

    mDisplay = 0 '預設模式：展開
    mDisCount = 0
    mSelCount = 0 '勾選X筆
    '待分案
    strMid(1) = Mid(strMid(1), 7)
    intQ = 1
    Set rsAD1 = ClsLawReadRstMsg(intQ, strMid(1))
    Call SetGrid(GRD1, 1, True)
    If intQ = 1 Then
         GRD1.FixedCols = 0
         Set GRD1.Recordset = rsAD1
         Call SetGrid(GRD1, 1, False)
         GRD1.FixedCols = cFixed
         '移動到上一次勾選的Row捲軸 ; 12/16 改成「更正勾選列」移到第一列 (SelGrid讀取m_SelCP09取得intLastTop)
         If GRD1.Rows >= intLastTop - 1 And intLastTop > 1 Then
            GRD1.TopRow = intLastTop
         End If
         intLastTop = 1
    ElseIf intQ = 0 Then
         MsgBox "目前無分案資料！", vbInformation
    End If
    
    '退智權 ＆ 退程序 (補件)
    strMid(2) = Mid(strMid(2), 7)
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strMid(2))
    Call SetGrid(grd2, 2, True)
    If intQ = 1 Then
         grd2.FixedCols = 0
         Set grd2.Recordset = rsQuery
         Call SetGrid(grd2, 2, False)
         grd2.FixedCols = cFixed
    End If
    QueryData = True
End Function

Private Sub cmdQuery_Click()
   
   cmdQuery.Enabled = False
   If QueryData = True Then
   End If
   cmdQuery.Enabled = True
End Sub

Private Sub cmdUpdRow_Click()
Dim intRow As Integer

   cmdUpdRow.Enabled = False
   If mSelCount > 0 Then
       If TxtValidate = True Then
           'intLastTop = GRD1.TopRow
           For intRow = 1 To GRD1.Rows - 1
              If "" & GRD1.TextMatrix(intRow, 0) = "V" And "" & GRD1.TextMatrix(intRow, colCP09_1) <> "" Then
                 'Added by Lydia 2024/11/06 更正時詢問是否變更CP14; T-247443在11/5同時收301變更+401訴願，在產生收文時301變更已預設CP14+CRC09，但是要經過確認才能分承辦人，所以在清空接洽單預分承辦人詢問是否一併將承辦人清空---嘉雯
                 bolUpdCP14 = False
                 If m_CRC09 <> "" And Trim(cboCRC09.Text) = "" Then
                    strExc(1) = "select cp14,st02 from caseprogress,staff where cp09='" & GRD1.TextMatrix(intRow, colCP09_1) & "' and cp14=st01(+) "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
                    If intI = 1 Then
                       If "" & RsTemp.Fields("CP14") <> "" Then
                          intI = MsgBox("案件進度已設定承辦人 " & RsTemp.Fields("CP14") & " " & RsTemp.Fields("st02") & "，是否清空承辦人？" & vbCrLf & _
                                  "選「是」會清空案件進度承辦人，" & vbCrLf & "選「否」不清空，" & vbCrLf & "選「取消」會取消本次更正。", vbYesNoCancel + vbDefaultButton3)
                          If intI = 6 Then  'Yes
                             bolUpdCP14 = True
                          ElseIf intI = 2 Then  'Cancel
                             cboCRC09.Text = m_CRC09 & " " & GetStaffName(m_CRC09, True)
                             cmdUpdRow.Enabled = True
                             Exit Sub
                          End If
                       End If
                    End If
                 End If
                 'end 2024/11/06
                 Call ChkT31xT11(GRD1.TextMatrix(intRow, colCP140_1)) 'Added by Lydia 2025/07/24
                 'Modified by Lydia 2023/07/26 + False
                 If SaveUpdRow("1", "" & GRD1.TextMatrix(intRow, colCP140_1), "" & GRD1.TextMatrix(intRow, colCP09_1), False) = False Then
                     GoTo EXITSUB
                 End If
                 If InStr(m_SysNo, "P") > 0 Then
                     m_SelCP09 = m_SelCP09 & "," & GRD1.TextMatrix(intRow, colCP09_1)
                 End If
                 GRD1.TextMatrix(intRow, 0) = ""
                 '商標案可複選，先清除預設
                 m_CRC09 = ""
                 m_CRC10 = ""
                 m_CRC11 = ""
                 m_CRC12 = ""
              End If
           Next intRow
       Else
           cmdUpdRow.Enabled = True
           Exit Sub
       End If
       
       '預設：畫面更新
       cmdUpdRow.Enabled = True
       Call QueryData
       'intLasTop=記錄最後勾選位置的第一列 ; 12/16 改成「更正勾選列」移到第一列=>清空記錄
       m_SelCP09 = ""
   End If
   cmdUpdRow.Enabled = True
   
EXITSUB:
   If mDisplay = 1 Then '預設：展開
       Call Command1_Click
   End If

End Sub
'判斷有無變更資料
Private Function DiffValidate() As Boolean
   
   DiffValidate = True
   If Trim(Left(cboCRC09.Text, 6)) <> m_CRC09 Or Trim(txtCRC10.Text) <> m_CRC10 Or Trim(txtNote) <> "" Or Trim(txtCRC11.Text) <> m_CRC11 Or Trim(txtCRC12.Text) <> m_CRC12 Then
        Exit Function
   End If
   '相同案號: 考慮母案被變動的可能, 不限制m_CRL55 <> lblData(0).Caption(母案顯示為空白)
   If cboCRL55.Enabled = True And Trim(cboCRL55.Text) <> m_CRL55 Then
        Exit Function
   End If
   If cboCRL67.Enabled = True And Trim(cboCRL67.Text) <> m_CRL67 Then
        Exit Function
   End If
   If cboCRL68.Enabled = True And Trim(cboCRL68.Text) <> m_CRL68 Then
        Exit Function
   End If
   If cboAttr.Enabled = True And Trim(cboAttr.Text) <> m_Attr Then
        Exit Function
   End If
   If Check1.Visible = True And Check1.Value = 1 Then
       Exit Function
   End If
   
   DiffValidate = False
End Function

'Modified by Lydia 2023/02/18
Private Function TxtValidate(Optional ByVal bolPass As Boolean) As Boolean
Dim tmpBol As Boolean

   TxtValidate = False
   
   If bolPass = False Then 'Added by Lydia 2023/02/18
      If DiffValidate = False Then
         Exit Function
      End If
   End If   'Added by Lydia 2023/02/18
   
   '承辦人
   If cboCRC09.Enabled = True And Trim(cboCRC09.Text) <> "" Then
      If ClsPDGetStaff(Trim(Left(cboCRC09.Text, 6)), strQuery) = False Then
         cboCRC09.SetFocus
         Exit Function
      End If
      'Added by Lydia 2024/06/19 特定大陸案分案給程序人員跳提醒
      If PUB_GetST03(Trim(Left(cboCRC09.Text, 6))) = "P12" And m_CP01 = "P" And m_Na01 = "020" And InStr("101,102", m_CP10) > 0 Then
         strExc(1) = Pub_GetField(" ConsultRecordList ", " CRL01='" & m_CP140 & "' ", "CRL138") '是否會簡體版
         '案件性質為發明申請(101)及實用新型(102)，申請國為中國(020)，是否會簡體版為是，若承辦人員為程序人員，則請於按下確定分案時，跳出提醒【本案要會簡體版，請確認是否仍給程序人員辦理】。若選是，則續行分案。若選否，則關閉提醒視窗。
         If strExc(1) = "Y" Then
            If MsgBox("本案要會簡體版，請確認是否仍給程序人員辦理？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               cboCRC09.SetFocus
               Exit Function
            End If
         Else
            '案件性質為新型(102)，申請國為中國(020)，案件屬性為2.電子電機，若承辦人員為程序人員，則請於按下確定分案時，跳出提醒【本案為大陸新型電子案，請確認是否仍給程序人員辦理】。若選是，則續行分案。若選否，則關閉提醒視窗。
            If m_CP10 = "102" And Trim(Left(cboAttr.Text, 1)) = "2" Then
               If MsgBox("本案為大陸新型電子案，請確認是否仍給程序人員辦理？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  cboCRC09.SetFocus
                  Exit Function
               End If
            End If
         End If
      End If
      'end 2024/06/19
   End If
   
   '相同案號
   If cboCRL55.Enabled = True And Trim(cboCRL55.Text) <> "" Then
      'Modified by Lydia 2023/10/26 +傳入判斷的關聯名稱
      If ChkInputCase(cboCRL55.Text, strQuery, , "相同案號-案號檢查") = False Then
          cboCRL55.SetFocus
          Exit Function
      End If
   End If
   '一案兩請
   If cboCRL67.Enabled = True And Trim(cboCRL67.Text) <> "" Then
      'Added by Lydia 2023/02/18
      If InStr(cboCRL67.Text, ",") > 0 Then
         MsgBox "不可輸入複數案號！"
         cboCRL67.SetFocus
         Exit Function
      Else
      'end 2023/02/18
         'Modified by Lydia 2023/10/26 +傳入判斷的關聯名稱
         If ChkInputCase(cboCRL67.Text, strQuery, "Y", "一案兩請-案號檢查") = False Then
             cboCRL67.SetFocus
             Exit Function
         End If
      End If 'Added by Lydia 2023/02/18
   End If
   '擬制喪失新穎性案號
   If cboCRL68.Enabled = True And Trim(cboCRL68.Text) <> "" Then
      'Added by Lydia 2023/02/18
      If InStr(cboCRL68.Text, ",") > 0 Then
         MsgBox "不可輸入複數案號！"
         cboCRL68.SetFocus
         Exit Function
      Else
      'end 2023/02/18
         'Modified by Lydia 2023/10/26 +傳入判斷的關聯名稱
         If ChkInputCase(cboCRL68.Text, strQuery, "擬制喪失新穎性-案號檢查") = False Then
             cboCRL68.SetFocus
             Exit Function
         End If
      End If 'Added by Lydia 2023/02/18
   End If
   '案件屬性
   If cboAttr.Enabled = True And Trim(cboAttr.Text) <> "" Then
       Call cboAttr_Validate(tmpBol)
      If tmpBol = True Then
          Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Private Sub Combo1_Change()
   If m_SysNo <> "" Then
       If ProcGetLock(Me.Name & "=" & Left(m_SysNo, 1) & "-" & Left(Combo1.Text, 1)) = True Then
          '分所人員代理北所
          'Mark by Lydia 2023/01/04 分所人員不可代理北所
          'If m_ST06 <> "1" Then
          '   If InStr("1,A", Left(Combo1.Text, 1)) > 0 Then
          '      m_nowF0308 = "A6"
          '   Else
           '     m_nowF0308 = "A5"
          '   End If
          'End If
          'Added by Lydia 2025/10/20 分所主管代北所主管分案
          If m_ST06 <> "1" Then
             If InStr("1,A", Left(Combo1.Text, 1)) > 0 And InStr(m_Str專利處台北區主管, strUserNum) > 0 Then
                m_nowF0308 = "A6"
             Else
                m_nowF0308 = "A5"
             End If
          End If
          'end 2025/10/20
          intLastTop = 1
          Call QueryData
          'Call SetCboCRC09
       End If
   End If
End Sub

Private Sub Command1_Click()
   
   '案件展開=0 / 收合=1
   If mDisplay = 0 Then
      Call SetExplore(1)
      If mDisplay = 1 Then '有案件被收合
         Command1.Caption = "案件展開"  '下一次
         Call CtrlOKbutton(False)
      End If
   ElseIf mDisplay = 1 Then
      Call SetExplore(0)
      Command1.Caption = "案件收合"  '下一次
      If mSelCount < 2 Then
          Call CtrlOKbutton(True)
      End If
   End If
End Sub

Private Sub Form_Load()
 
   '載入前次結束時的大小及位置
   PUB_SetPdfForm Me, False
   
   Call SetRoleData
   Call QueryData
   'Call SetCboCRC09
   '不限制第4碼為9的員工
   strQuery = "SELECT ST01, ST02 FROM STAFF WHERE ST04='1' AND ST01>'6' AND SUBSTR(ST01,1,1)<'G' AND ST01 NOT LIKE 'F%' "
   If m_SysNo = "T" Then
       strQuery = strQuery & "AND ST03='P21' "
   Else
       strQuery = strQuery & "AND (ST03='P11' or ST03='P10') "
   End If
   strQuery = strQuery & " ORDER BY ST01 DESC "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   cboCRC09.Clear
   If intQ = 1 Then
       With rsQuery
            .MoveFirst
            Do While Not .EOF
                If "" & .Fields("ST01") <> "" And "" & .Fields("ST02") <> "" Then
                    cboCRC09.AddItem convForm("" & .Fields("ST01"), 6) & " " & .Fields("ST02"), 0
                End If
                .MoveNext
            Loop
       End With
       cboCRC09.AddItem String(10, " "), 0
       '不預設
       'cboCRC09.ListIndex = 0
       'cboCRC09.Tag = cboCRC09.Text
   End If
   
   Me.SSTab1.Tab = 0
   If ProcGetLock(Me.Name & "=" & Left(m_SysNo, 1) & "-" & Left(Combo1.Text, 1)) = True Then
   End If
   
   'Added by Lydia 2025/07/24
   bolDutyT11 = False
   If Pub_StrUserSt93 = "T31" Then
      strExc(0) = Pub_GetSpecMan("內商爭議案程序主管")
      If strExc(0) <> "" Then
         strExc(0) = Replace(Replace(Left(strExc(0), 6), ",", ""), ";", "")
         strExc(1) = GetCaseDutyAgent(strExc(0), "", False, , True, "A")
         If strExc(1) <> "" Then
            bolDutyT11 = True
         End If
      End If
   End If
   'end 2025/07/24
   
   m_Str專利處台北區主管 = Pub_GetSpecMan("A2") 'Added by Lydia 2025/08/18
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Call ProcGetLock("A" & Me.Name)
   
   PUB_SavePdfForm Me '紀錄視窗最後的大小及位置
   Set rsAD1 = Nothing
   Set rsQuery = Nothing
   Unload Me
End Sub

Private Sub Grd1_Click()
'12/15 改成Ctrl為複選
'Dim intRow As Integer, intCol As Integer
'
'   With GRD1
'       intCol = .MouseCol
'       If .MouseRow > 0 Then
' '         intLastTop = GRD1.TopRow
'       End If
'       If .MouseRow > 0 And intCol = 0 And blnDBClick = False Then
'          intRow = .MouseRow
'          .row = intRow
'          .col = cFixed + 1 '取固定欄位後的底色，排除因為全選的變色
'          lngColor = .CellBackColor
'          GridClick GRD1, intRow, 0, 0, cFixed, "V", lngColor, colCRL55n_1 & "," & colCRL67n_1
'          If "" & GRD1.TextMatrix(intRow, colCaseNo_1) <> "" Then
'               Call GetSelState(intRow)
'               If "" & GRD1.TextMatrix(intRow, 0) = "V" Then
'                   If PUB_CheckFormExist("frm090801_Q") = True Then
'                       Unload frm090801_Q
'                   End If
'               End If
'          End If
'       End If
'   End With
'
'   blnDBClick = False
End Sub

Private Sub grd2_Click()
'12/15 改成Ctrl為複選
'Dim intRow As Integer, intCol As Integer
'
'   With GRD2
'       intCol = .MouseCol
'       If .MouseRow > 0 And intCol = 0 And blnDBClick = False Then
'          intRow = .MouseRow
'          .row = intRow
'          .col = cFixed + 1 '取固定欄位後的底色，排除因為全選的變色
'          lngColor = .CellBackColor
'          GridClick GRD2, intRow, 0, 0, cFixed, "V", lngColor, colCRL55n_1 & "," & colCRL67n_1
'          If "" & GRD2.TextMatrix(intRow, colCaseNo_1) <> "" Then
'               If "" & GRD2.TextMatrix(intRow, 0) = "V" Then
'                   If PUB_CheckFormExist("frm090801_Q") = True Then
'                       Unload frm090801_Q
'                   End If
'               End If
'          End If
'       End If
'   End With
'
'   blnDBClick = False
End Sub
'Modified by Lydia 2023/01/04 傳入收文號strCP09
Private Sub ShowFormCRL(ByVal strNo As String, ByVal strCP09 As String)

   If PUB_CheckFormExist("frm090801_Q") = True Then
        Unload frm090801_Q
   End If
   frm090801_Q.SetParent Me
   frm090801_Q.m_blnCallPrint = True
   frm090801_Q.Text5 = strNo
   Call frm090801_Q.cmdok_Click(4)
   frm090801_Q.Show
   
End Sub

Private Sub GRD1_DblClick()
Dim intRow As Integer, intCol As Integer
    
   intRow = GRD1.MouseRow
   intCol = GRD1.MouseCol
   If intRow > 0 Then
      '12/15 改成Ctrl為複選
      'If intCol = 0 Or intCol = 1 Then 'V的Dbclick改成指定勾選, 展開也列入
      '    If intCol = 0 Then
      '       If "" & GRD1.TextMatrix(intRow, colCaseNo_1) <> "" Then
      '            Call GetAllSelType(GRD1, "0", intRow)
      '            blnDBClick = True
      '       End If
      '    End If
      'Else
      'Added by Lydia 2023/01/06 增加雙擊「相同案號」欄位開啟"相同案號"的進度檔畫面
      If intCol = colCRL55n_1 Then
         If GRD1.TextMatrix(intRow, colCRL55_1) <> "" Then
             Call GetAllSelType(GRD1, "0", intRow)
             If ShowCPgrid(GRD1.TextMatrix(intRow, colCRL55_1)) = True Then
             End If
         End If
      Else
      'end 2023/01/06
         If GRD1.TextMatrix(intRow, colCP140_1) <> "" Then
            Screen.MousePointer = vbHourglass
                Call GetAllSelType(GRD1, "0", intRow)
                Call ShowFormCRL(GRD1.TextMatrix(intRow, colCP140_1), GRD1.TextMatrix(intRow, colCP09_1))
                blnDBClick = True
            Screen.MousePointer = vbDefault
         End If
      End If 'Added by Lydia 2023/01/06
      'End If '12/15 改成Ctrl為複選
   End If
End Sub

Private Sub grd2_DblClick()
Dim intRow As Integer

   intRow = grd2.MouseRow
   If intRow > 0 Then
      If grd2.TextMatrix(intRow, colCP140_1) <> "" Then
         Screen.MousePointer = vbHourglass
             Call GetAllSelType(grd2, "0", intRow)
             Call ShowFormCRL(grd2.TextMatrix(intRow, colCP140_1), grd2.TextMatrix(intRow, colCP09_1))
             blnDBClick = True
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

'更新維護區
Private Sub UpdateCtrlData(ByVal pRow As Integer)
   
   'Modified by Lydia 2023/01/05 不取消勾選和輸入未更正的內容 from 李柏翰
   'If pRow >= 1 Then
   If pRow >= 1 And m_PrevNo <> m_CP09 Then
       Call ClearCtrlData
       m_CP140 = "" & GRD1.TextMatrix(pRow, colCP140_1)
       m_CP09 = "" & GRD1.TextMatrix(pRow, colCP09_1)
       m_idX = pRow
       cboCRC09.Text = "" & GRD1.TextMatrix(pRow, colCRC09_1)
       If cboCRC09.Text <> "" Then
           Call cboCRC09_Validate(False)
       End If
       m_CRC09 = "" & GRD1.TextMatrix(pRow, colCRC09_1)
       
       m_F0316 = "" & GRD1.TextMatrix(pRow, colF0316_1) '智權人員
       m_F0309 = "" & GRD1.TextMatrix(pRow, colF0309_1)
       '是否算案件數CRC10=>CP26
       txtCRC10.Text = "" & GRD1.TextMatrix(pRow, colCRC10_1)
       m_CRC10 = txtCRC10.Text
       '承辦人計件值CRC11=>CP97
       txtCRC11.Text = "" & GRD1.TextMatrix(pRow, colCRC11_1)
       m_CRC11 = txtCRC11.Text
       '承辦人加乘註記CRC12=>CP98
       txtCRC12.Text = "" & GRD1.TextMatrix(pRow, colCRC12_1)
       m_CRC12 = txtCRC12.Text
       
       m_CP01 = "" & GRD1.TextMatrix(pRow, colCP01_1)
       m_CP02 = "" & GRD1.TextMatrix(pRow, colCP02_1)
       m_CP03 = "" & GRD1.TextMatrix(pRow, colCP03_1)
       m_CP04 = "" & GRD1.TextMatrix(pRow, colCP04_1)
       m_CP10 = "" & GRD1.TextMatrix(pRow, colCP10_1)
       m_CP31 = "" & GRD1.TextMatrix(pRow, colCP31_1)
       m_CKind = "" & GRD1.TextMatrix(pRow, colCKind_1)
       m_Na01 = "" & GRD1.TextMatrix(pRow, colNA01_1)
       m_APP01 = "" & GRD1.TextMatrix(pRow, colAPP01_1)
       m_CRL65 = "" & GRD1.TextMatrix(pRow, colGrpNO_1)
       '相同案號CRL65,一案兩請CRL66,擬制喪失新穎性CRL68: 因為子案記錄母案案號，而母案則記錄本身案號所以顯示為空白；為了分辨有無設定，另外抓資料
       cboCRL55.Text = "" & GRD1.TextMatrix(pRow, colCRL55n_1)
       cboCRL67.Text = "" & GRD1.TextMatrix(pRow, colCRL67n_1)
       cboCRL68.Text = "" & GRD1.TextMatrix(pRow, colCRL68n_1)
       m_CRL55 = "" & GRD1.TextMatrix(pRow, colCRL55_1)
       m_CRL67 = "" & GRD1.TextMatrix(pRow, colCRL67_1)
       m_CRL68 = "" & GRD1.TextMatrix(pRow, colCRL68_1)
       '商標案不用, 新案才能更改「案件屬性」，舊案若要更改由程序人員負責
       'Modified by Lydia 2023/10/30 因為專利可能先有專利調查才收發明申請,所以改成判斷新案性質; ex.CFP-034069
       'If InStr(m_CP01, "T") > 0 Or m_CP31 <> "Y" Or InStr(NewCasePtyList, m_CP10) = 0 Then
       If InStr(m_CP01, "T") > 0 Or (InStr(m_CP01, "T") = 0 And InStr(NewCasePtyList, m_CP10) = 0) Then
          cboCRL55.Text = "" '不用顯示
          cboAttr.Enabled = False
          cboCRL55.Enabled = False
          cboCRL67.Enabled = False
          cboCRL68.Enabled = False
          Call ShowObject(False)
       Else
          If (m_CP01 <> "P" And m_CP01 <> "CFP") Or (m_CP01 = "P" And m_CP01 = "CFP" And m_Na01 <> "000") Then '商標和專利非台灣案不可設定案件屬性
             cboAttr.Enabled = False
          Else
             cboAttr.Enabled = True
          End If
          cboCRL55.Enabled = True
          cboCRL67.Enabled = True
          cboCRL68.Enabled = True
          Call ShowObject(True)
       End If
       '一案兩請欄位要鎖住不可以輸入
       'Modified by Lydia 2023/03/07 開放發明和新型案都可以輸入; ex.3/6收的P-131075,P-131076發明案和2/14收文P-130986,P130987設一案兩請
       'If cboCRL67.Enabled = True And m_CP10 <> "102" Then
       If cboCRL67.Enabled = True And InStr("101,102", m_CP10) = 0 Then
           cboCRL67.Enabled = False
       End If
       Check1.Value = False
       If cboCRL67.Enabled = True And cboCRL67.Visible = True And m_CRL55 <> "" And m_CRL67 <> "" Then
          Check1.Visible = True
       Else
          Check1.Visible = False
       End If
       
       '擬制喪失新穎性欄位要鎖住不可以輸入; 111/11/17 擬制喪失新穎性不限制103
       'If cboCRL68.Enabled = True And m_CP10 <> "103" Then
       '    cboCRL68.Enabled = False
       'End If
       '相同案號若為PS,CPS案的967智財協作鎖定不可修改
       If (m_CP01 = "PS" And m_CP10 = "967") Or (m_CP01 = "CPS" And m_CP10 = "967") Then
           cboCRL55.Enabled = False
       End If
       
       If m_CRL55 <> "" & GRD1.TextMatrix(pRow, colCaseNo_1) Then Call SetCboCRL(cboCRL55, m_CRL55)
       If m_CRL67 <> "" & GRD1.TextMatrix(pRow, colCaseNo_1) Then Call SetCboCRL(cboCRL67, m_CRL67)
       If m_CRL68 <> "" & GRD1.TextMatrix(pRow, colCaseNo_1) Then Call SetCboCRL(cboCRL68, m_CRL68)

       'PA158案件屬性
       If cboAttr.Enabled = True Then
           Call PUB_AddCaseAttributeCombo(cboAttr, m_CKind)
           If "" & GRD1.TextMatrix(pRow, colCAttr_1) <> "" Then
              cboAttr.Text = "" & GRD1.TextMatrix(pRow, colCAttr_1) + "." + PUB_GetCaseAttributeName(GRD1.TextMatrix(pRow, colCAttr_1), m_CKind)
           End If
           m_Attr = cboAttr.Text
       End If
       
       cboCRL55.Tag = cboCRL55.Text:         cboCRL67.Tag = cboCRL67.Text:        cboCRL68.Tag = cboCRL68.Text:        cboAttr.Tag = cboAttr.Text
       
       lblData(0) = GRD1.TextMatrix(pRow, colCaseNo_1)   '本所案號
       lblData(1) = "" & GRD1.TextMatrix(pRow, colCP10n_1) '案件性質名稱
       '總費用,總點數
       Call GetTotFeeDot(m_CP01, m_CP02, m_CP03, m_CP04)
       
       'Added by Morgan 2025/7/3 --柏翰
       '1.智權人員為郭雅娟79075
       '2.案件性質為發明申請101或新型申請102
       '案件屬性為1.機械時，加乘註記0.4
       '案件屬性為2.電子電機或3.化學生醫時，加乘註記0.3
       If txtCRC12 = "" And "" & GRD1.TextMatrix(pRow, colCP13_1) = "79075" And m_CP01 = "P" And (m_CP10 = "101" Or m_CP10 = "102") Then
         If GRD1.TextMatrix(pRow, colCAttr_1) = "1" Then
            txtCRC12 = "0.4"
         Else
            txtCRC12 = "0.3"
         End If
       End If
       'end 2025/7/3
   End If
End Sub

Private Sub GetTotFeeDot(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String)

  If pCP01 <> "" And pCP02 <> "" Then
      '只抓目前尚未分案
      'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉and nvl(cpm35,'0')<>'2'
      strQuery = "select sum(nvl(cp16,0)) totfee,sum(nvl(cp18,0)) totdot from caseprogress, CasePropertyMap" & _
                 " where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' " & _
                 "and cp158=0 and cp159=0 and cp157 is null and cp01=cpm01(+) and cp10=cpm02 "
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
      If intQ = 1 Then
          lblData(2) = Val("" & rsQuery.Fields("totfee"))
          lblData(3) = Val("" & rsQuery.Fields("totdot"))
      End If
  End If
End Sub

Private Sub ClearCtrlData()
    m_CP01 = ""
    m_CP02 = ""
    m_CP03 = ""
    m_CP04 = ""
    m_CP09 = ""
    m_idX = -1
    m_CP140 = ""
    cboCRC09.Text = "": cboCRC09.Tag = "": m_CRC09 = ""
    txtCRC10 = "":  m_CRC10 = ""
    txtCRC11 = "":  m_CRC11 = ""
    txtCRC12 = "":  m_CRC12 = ""
    m_CP10 = ""
    m_CP31 = ""
    m_CKind = ""
    m_Na01 = ""
    m_APP01 = ""
    m_CRL65 = ""
    lblName.Caption = ""
    cboCRL55.Text = "":      cboCRL67.Text = "":     cboCRL68.Text = "":     cboAttr.Text = ""
    cboCRL55.Tag = "":     cboCRL67.Tag = "":     cboCRL68.Tag = "":    cboAttr.Tag = ""
    m_CRL55 = "":     m_CRL67 = "":     m_CRL68 = "":     m_Attr = ""
    cboAttr.Enabled = True
    cboCRL55.Enabled = True
    cboCRL67.Enabled = True
    cboCRL68.Enabled = True
    For Each oObj In lblData
        oObj.Caption = ""
    Next
    txtNote.Text = ""
    blnDBClick = False
    Check1.Value = False
    
End Sub

Private Sub CtrlOKbutton(ByVal bEnabled As Boolean)
    'Added by Lydia 2023/01/04 北所主管可以對全所＋各分所做「確定分案」，分所可以看他所案件但僅可以「更正」不可分案(from 副總)
    'Modified by Lydia 2025/08/18 林柄佑協理有權限對對全所＋各分所做「確定分案」
    'If InStr(m_SysNo, "P") > 0 And m_ST06 <> "1" And Left(Combo1, 1) <> m_ST06  Then
    If InStr(m_SysNo, "P") > 0 And m_ST06 <> "1" And Left(Combo1, 1) <> m_ST06 And InStr(m_Str專利處台北區主管, strUserNum) = 0 Then
        cmdUpdRow.Enabled = bEnabled
        cmdOK(0).Enabled = bEnabled
        cmdOK(1).Enabled = False
        cmdOK(2).Enabled = bEnabled
        cmdOK(3).Enabled = bEnabled
    Else
'    'end 2023/01/04
        cmdUpdRow.Enabled = bEnabled
        cmdOK(0).Enabled = bEnabled
        cmdOK(1).Enabled = bEnabled
        cmdOK(2).Enabled = bEnabled
        cmdOK(3).Enabled = bEnabled
    End If
End Sub
Private Sub SetExplore(ByVal pDisplay As Integer)
   Dim stType As String
   Dim intS As Integer
      
   blnDBClick = False
   mDisCount = 0 '收合的案件數
   '案件展開=0 / 收合=1
   With GRD1
      If .Rows > 1 Then
          .Visible = False
          If pDisplay = 0 Then '展開
              For intP = 1 To .Rows - 1
                 .RowHeight(intP) = 255
                 .TextMatrix(intP, colExp_1) = ""
              Next intP
              mDisplay = 0
          ElseIf pDisplay = 1 Then '收合
              For intP = 1 To .Rows - 1
                 If stType <> "" & .TextMatrix(intP, colCaseNo_1) Then
                     .RowHeight(intP) = 255
                     stType = "" & .TextMatrix(intP, colCaseNo_1)
                     intS = intP '記錄第一筆收文
                 Else
                     .RowHeight(intP) = 0
                     If "" & .TextMatrix(intS, colExp_1) = "" Then '加上收合註記
                        .TextMatrix(intS, colExp_1) = "＋"
                        mDisCount = mDisCount + 1 '收合的案件數
                     End If
                 End If
              Next intP
              If mDisCount > 0 Then
                  mDisplay = 1
              End If
          End If
          .TopRow = 1
          .Visible = True
       End If
   End With

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   blnDBClick = False
End Sub

Private Sub txtCRC10_GotFocus()
   TextInverse txtCRC10
End Sub

Private Sub txtCRC10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
   End If
End Sub

Private Sub cboCRL55_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboCRL67_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboCRL68_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboAttr_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'檢查案號的格式
Private Function GetSameInput(ByVal pNowNo As String) As String
Dim tmpArr As Variant
   GetSameInput = ""
   If pNowNo <> "" Then
      tmpArr = Split(pNowNo, ",")
      For intQ = 0 To UBound(tmpArr)
          strBCase(0) = Replace(Trim(tmpArr(intQ)), "-", "")
          Call ChgCaseNo(strBCase(0), strBCase)
          If strBCase(1) = "" Or strBCase(2) = "" Then
              MsgBox "錯誤案號：" & pNowNo & vbCrLf & "請輸入正確本所案號！"
              Exit Function
          Else
              GetSameInput = GetSameInput & "," & strBCase(1) & "-" & strBCase(2) & IIf(strBCase(4) <> "00", "-" & strBCase(3) & "-" & strBCase(4), IIf(strBCase(3) <> "0", "-" & strBCase(3), ""))
          End If
      Next intQ
   End If
   If GetSameInput <> "" Then GetSameInput = Mid(GetSameInput, 2)
   
End Function

'檢查多案的案號輸入
'Modified by Lydia 2023/10/26 +傳入判斷的關聯名稱pMsgTitle
Private Function ChkInputCase(ByVal pCaseNo As String, ByRef pNewNo As String, Optional ByVal pKind As String = "N", Optional ByVal pMsgTitle As String) As Boolean
Dim bolCaseIsExists As Boolean
Dim tmpArr  As Variant
Dim strA1 As String

    ChkInputCase = False
    pNewNo = ""
    If Trim(pCaseNo) <> "" Then
       tmpArr = Split(pCaseNo, ",")
       For intQ = 0 To UBound(tmpArr)
           strBCase(0) = Replace(Pub_RplStr(Trim(tmpArr(intQ))), "-", "")
           If strBCase(0) <> "" Then
              Call ChgCaseNo(strBCase(0), strBCase)
              If strBCase(1) = "" Or strBCase(2) = "" Then
                 'Modified by Lydia 2023/10/26 +pMsgTitle
                 MsgBox "錯誤案號：" & tmpArr(intQ) & vbCrLf & "請輸入正確本所案號！", , pMsgTitle
                 Exit Function
              ElseIf strBCase(1) & strBCase(2) & strBCase(3) & strBCase(4) = m_CP01 & m_CP02 & m_CP03 & m_CP04 Then
                 MsgBox "錯誤案號：" & tmpArr(intQ) & vbCrLf & "不可輸入相同號！", , pMsgTitle
                 Exit Function
              ElseIf InStr(pNewNo & ",", strBCase(1) & "-" & strBCase(2) & IIf(strBCase(4) <> "00", "-" & strBCase(3) & "-" & strBCase(4), IIf(strBCase(3) <> "0", "-" & strBCase(3), ""))) > 0 Then
                 MsgBox "錯誤案號：" & tmpArr(intQ) & vbCrLf & "不可輸入重複案號！", , pMsgTitle
                 Exit Function
              Else
                 strExc(0) = "select pa26,nvl(cu04,nvl(cu05,cu06)) pa26n, pa57, pa08, pa09,nvl(pa05,nvl(pa06,pa07)) casename from patent,customer " & _
                                  "where pa01='" & strBCase(1) & "' and pa02='" & strBCase(2) & "' and pa03='" & strBCase(3) & "' and pa04='" & strBCase(4) & "' " & _
                                  "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                 If intI = 0 Then
                     MsgBox "錯誤案號：無此案件基本資料，請檢查是否案號輸入錯誤！", , pMsgTitle
                     Exit Function
                 Else
                     If "" & RsTemp.Fields("PA57") <> "" Then
                        'Modified by Lydia 2023/12/08 改成彈提醒---李柏翰
                        'MsgBox "錯誤案號：此案件已閉卷！", , pMsgTitle
                        If MsgBox(strBCase(0) & "此案件已閉卷，是否繼續作業？", vbInformation + vbYesNo + vbDefaultButton2, pMsgTitle) = vbNo Then
                           Exit Function
                        End If 'Added by Lydia 2023/12/08
                     End If
                     If "" & RsTemp.Fields("casename") <> "" & GRD1.TextMatrix(m_idX, colCaseName_1) Then
                        If MsgBox(strBCase(0) & " " & RsTemp.Fields("casename") & vbCrLf & "案件名稱與本案不同，是否繼續作業？", vbInformation + vbYesNo + vbDefaultButton2, pMsgTitle) = vbNo Then
                           Exit Function
                        End If
                     End If
                     '一案兩請的檢查
                     If pKind = "Y" Then
                        If "" & RsTemp.Fields("PA26") <> m_APP01 Then
                            MsgBox "錯誤案號：申請人不同！" & vbCrLf & RsTemp.Fields("pa26") & "：" & RsTemp.Fields("pa26n"), , pMsgTitle
                            Exit Function
                        End If
                        If strBCase(1) <> "" & m_CP01 Then
                            MsgBox "錯誤案號：系統別不同！", , pMsgTitle
                            Exit Function
                        End If
                        If m_Na01 <> "" & RsTemp.Fields("pa09") Then
                            MsgBox "錯誤案號：申請國家不同！", , pMsgTitle
                            Exit Function
                        End If
                        If InStr("000,020,231", m_Na01) = 0 Then
                            MsgBox "錯誤案號：一案兩請之申請國家只有臺灣、大陸、德國！", , pMsgTitle
                            Exit Function
                        End If
                        If m_CKind = "" & RsTemp.Fields("pa08") Then
                           MsgBox "錯誤案號：請輸入" & IIf(m_CKind = "1", "新型案之案號", "發明案之案號") & "！", , pMsgTitle
                           Exit Function
                        End If
                     End If
                 End If
                 pNewNo = pNewNo & "," & strBCase(1) & "-" & strBCase(2) & IIf(strBCase(4) <> "00", "-" & strBCase(3) & "-" & strBCase(4), IIf(strBCase(3) <> "0", "-" & strBCase(3), ""))

              End If
           End If
       Next intQ
    End If
    If pNewNo <> "" Then pNewNo = Mid(pNewNo, 2)
    
    ChkInputCase = True
End Function

Private Sub txtCRC10_Validate(Cancel As Boolean)
     If txtCRC10 <> "" And (m_CP09 = "" Or m_idX = 0) Then
        MsgBox "請先選取資料！", vbInformation
        Cancel = True
        txtCRC10.SetFocus
        Call txtCRC10_GotFocus
     End If
End Sub

Private Sub txtCRC09_GotFocus()
    TextInverse txtCRC09
End Sub

Private Sub txtCRC09_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCRC09_Validate(Cancel As Boolean)
    
    If Trim(txtCRC09) = "" Then
        lblName = ""
    ElseIf txtCRC09.Tag <> txtCRC09.Text Then
        lblName = ""
        'Mark by 有可能先輸入再點選
        'If m_CP09 = "" Or m_idX = 0 Then
        '  MsgBox "請先選取資料！", vbInformation
        '  GoTo EXITSUB
        'End If
        If ClsPDGetStaff(txtCRC09, strQuery) = False Then
            GoTo EXITSUB
        Else
            lblName = strQuery
        End If
    End If
    txtCRC09.Tag = txtCRC09.Text
    
    Exit Sub
    
EXITSUB:
    Cancel = True
    txtCRC09.SetFocus
    Call txtCRC09_GotFocus
End Sub

Private Sub txtNote_Change()
   PUB_RefreshText txtNote
End Sub

Private Sub txtNote_GotFocus()
   TextInverse txtNote
End Sub

'Modified by Lydia 2023/07/26 bolSpace承辦人是否空白
Private Function SaveUpdRow(ByVal mStatus As String, ByVal pCP140 As String, ByVal pCP09 As String, ByVal bolSpace As Boolean) As Boolean
'mStatus: 1-更正，2-確認分案
'Added by Lydia 2024/04/18
Dim strSNo As String, strB1 As String, strB2 As String, strNoList As String, strCrl65List As String
Dim intB As Integer, rsBD As New ADODB.Recordset
Dim pCRL65 As String, pCRL55 As String, pCRL67 As String, pCRL68 As String
Dim cCRL65 As String, cCRL55 As String, cCRL67 As String, cCRL68 As String

    SaveUpdRow = False
    If DiffValidate = True Then '判斷有變動才進入
    
On Error GoTo ErrHandle
       cnnConnection.BeginTrans
          '預分承辦人=承辦人
          If cboCRC09.Enabled = True And Trim(Left(cboCRC09.Text, 6)) <> m_CRC09 Then
             'Modified by Lydia 2023/07/26 判斷承辦人是否空白+bolSpace
             'Modified by Lydia 2025/06/17 商申ST93=T31排除商爭案性質 IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "")
             strSql = "UPDATE ConsultRecCMP SET CRC09='" & Trim(Left(cboCRC09.Text, 6)) & "' WHERE CRC01='" & pCP140 & "' AND CRC08='" & pCP09 & "' " & IIf(bolSpace = True, "AND CRC09 IS NULL ", "") & IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "")
             cnnConnection.Execute strSql
             '同一張接洽單的案件性質，一併更新同人員（前題是預分承辦人是空白的），但可再修改。
             'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉AND NVL(CPM35,'0')<>'2'
             'Modified by Lydia 2025/06/17 商申ST93=T31排除商爭案性質 IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "")
             strSql = "UPDATE ConsultRecCMP SET CRC09='" & Trim(Left(cboCRC09.Text, 6)) & "' WHERE CRC01='" & pCP140 & "' AND CRC09 IS NULL AND CRC08 IN (" & _
                          "SELECT CRC08 FROM ConsultRecordList,ConsultRecCMP,CasePropertyMap WHERE CRL01='" & pCP140 & "' AND CRL01=CRC01(+) AND CRL07=CPM01(+) AND CRC03=CPM02(+) " & _
                          IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "") & ") "
             cnnConnection.Execute strSql
             'Added by Lydia 2024/11/06 更正時詢問是否變更CP14; T-247443在11/5同時收301變更+401訴願，在產生收文時301變更已預設CP14+CRC09，但是要經過確認才能分承辦人，所以在清空接洽單預分承辦人詢問是否一併將承辦人清空---嘉雯
             If bolUpdCP14 = True Then
                'Modified by Lydia 2025/06/17 商申ST93=T31排除商爭案性質 IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',cp10)=0 ", "")
                strSql = "Update Caseprogress set cp14=" & CNULL(Trim(Left(cboCRC09.Text, 6))) & " where cp09='" & pCP09 & "' " & IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',cp10)=0 ", "")
                Pub_SeekTbLog strSql
                cnnConnection.Execute strSql
             End If
             'end 2024/11/06
             
               'Modified by Lydia 2023/01/06 專利案分所主管分案先不上進度檔的承辦人，要等到北所主管執行「確定分案」
               'If mStatus = "2" Then '確定分案
               'Modified by Lydia 2025/10/20 改成判斷權限
               'If mStatus = "2" And m_ST06 = "1" Then
               If mStatus = "2" And m_nowF0308 = "A6" Then
                  'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉AND NVL(CPM35,'0')<>'2'
                  'Modified by Lydia 2023/01/17 debug: CFP-32906 同時收422+417,承辦人分別設工程師和程序
                  'strSql = "Update CaseProgress set CP14='" & Trim(Left(cboCRC09.Text, 6)) & "' Where CP140='" & pCP140 & "' AND CP158=0 AND CP159=0 AND CP14 IS NULL AND CP09 IN (" & _
                              "SELECT CRC08 FROM ConsultRecordList,ConsultRecCMP,CasePropertyMap WHERE CRL01='" & pCP140 & "' AND CRL01=CRC01(+) AND CRL07=CPM01(+) AND CRC03=CPM02(+) ) "
                  'Modified by Lydia 2023/05/19 一律以主管分案為準AND CP159=0 AND CP14 IS NULL => AND CP159=0;ex.P-131540(AB2018840)因為主管指示程序退回重新分案換工程師，但是沒有回寫新的承辦人，經過與Sindy討論決定都以主管為主
                  'Modified by Lydia 2025/06/17 商申ST93=T31排除商爭案性質 IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "")
                  strSql = "Update CaseProgress set CP14=(select crc09 from ConsultRecordList,ConsultRecCMP,CasePropertyMap " & _
                               "where CRL01='" & pCP140 & "' AND CRL01=CRC01(+) AND CRL07=CPM01(+) AND CRC03=CPM02(+) and crc08=cp09 and crc09 is not null " & IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "") & ") " & _
                               "where cp140='" & pCP140 & "' AND CP158=0 AND CP159=0 "
                  cnnConnection.Execute strSql
               End If
             GRD1.TextMatrix(m_idX, colCRC09_1) = Trim(Left(cboCRC09.Text, 6))
             GRD1.TextMatrix(m_idX, colCRC09n_1) = Trim(Mid(cboCRC09.Text, 7))
             'm_CRC09 = Trim(Left(cboCRC09.Text, 6)) '為了SaveGrpData的判斷, 改到最後面
          End If
          
          'Added by Lydia 2024/04/18
          pCRL55 = m_CRL55
          pCRL65 = m_CRL65
          pCRL67 = m_CRL67
          pCRL68 = m_CRL68
          '母號=>本所案號
          If cboCRL55.Text <> "" Then
             cCRL55 = cboCRL55.Text
          ElseIf pCRL55 <> "" Then
             strB1 = GetSameInput(m_CP01 & m_CP02 & m_CP03 & m_CP04)
             If pCRL55 = strB1 Then
                cCRL55 = strB1
             End If
          End If
          If cboCRL67.Text <> "" Then
             cCRL67 = cboCRL67.Text
          ElseIf pCRL67 <> "" Then
             strB1 = GetSameInput(m_CP01 & m_CP02 & m_CP03 & m_CP04)
             If pCRL67 = strB1 Then
                cCRL67 = strB1
             End If
          End If
          If cboCRL68.Text <> "" Then
             cCRL68 = cboCRL68.Text
          ElseIf pCRL68 <> "" Then
             strB1 = GetSameInput(m_CP01 & m_CP02 & m_CP03 & m_CP04)
             If pCRL68 = strB1 Then
                cCRL68 = strB1
             End If
          End If
          'end 2024/04/18
          
          If m_CP01 = "P" Or m_CP01 = "CFP" Then
             '相同案號：主管只管相同案CRL56=Y
             If SaveGrpData(mStatus, cboCRL55.Enabled, pCP140, "CRL55", cboCRL55.Text, m_CRL55, Trim(Left(cboCRC09.Text, 6)), m_CRL65, m_CRL67) = True Then
             End If
             '一案兩請
             If SaveGrpData(mStatus, cboCRL67.Enabled, pCP140, "CRL67", cboCRL67.Text, m_CRL67, Trim(Left(cboCRC09.Text, 6)), m_CRL65, m_CRL67) = True Then
             End If
             '擬制喪失新穎性
             If SaveGrpData(mStatus, cboCRL68.Enabled, pCP140, "CRL68", cboCRL68.Text, m_CRL68, Trim(Left(cboCRC09.Text, 6)), m_CRL65, m_CRL67) = True Then
             End If
          End If
          
          'Added by Lydia 2024/04/18 同一組(相同案號+一案兩請+擬制喪失新穎性)的所有案件一定要全部在一起; 以關聯代號為主做排序
          strB2 = ""
          If Trim(pCRL55 & pCRL67 & pCRL68) <> Trim(cCRL55 & cCRL67 & cCRL68) Or (Check1.Value = 1 And Check1.Visible = True) Then
             If Trim(cCRL55 & cCRL67 & cCRL68) <> "" Then
                '統一CRL65的最小流水號
                If InStr(strB2 & ",", Trim(cCRL55)) = 0 And Trim(cCRL55) <> "" Then
                   strB2 = strB2 & " OR INSTR(CRL55||','||CRL67||','||CRL68,'" & Trim(cCRL55) & "') > 0"
                End If
                If InStr(strB2 & ",", Trim(cCRL67)) = 0 And Trim(cCRL67) <> "" Then
                   strB2 = strB2 & " OR INSTR(CRL55||','||CRL67||','||CRL68,'" & Trim(cCRL67) & "') > 0"
                End If
                If InStr(strB2 & ",", Trim(cCRL68)) = 0 And Trim(cCRL68) <> "" Then
                   strB2 = strB2 & " OR INSTR(CRL55||','||CRL67||','||CRL68,'" & Trim(cCRL68) & "') > 0"
                End If
                strB1 = " select crl01,crl07,crl08,crl09,crl10,crl55,crl56,crl65,crl67,crl68" & _
                        " from consultrecordlist, flow003 where crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) " & _
                        " and (" & Mid(strB2, 5) & ") group by crl01,crl07,crl08,crl09,crl10,crl55,crl56,crl65,crl67,crl68 order by crl01 "
                intB = 1
                Set rsBD = ClsLawReadRstMsg(intB, strB1)
                If intB = 1 Then
                   rsBD.MoveFirst
                   strB2 = ""
                   Do While Not rsBD.EOF
                      If InStr(strB2, "CRL55='" & rsBD.Fields("crl55") & "'") = 0 And "" & rsBD.Fields("crl55") <> "" Then
                         strB2 = strB2 & " OR CRL55='" & rsBD.Fields("crl55") & "'"
                      End If
                      If InStr(strB2, "CRL67='" & rsBD.Fields("crl67") & "'") = 0 And "" & rsBD.Fields("crl67") <> "" Then
                         strB2 = strB2 & " OR CRL67='" & rsBD.Fields("crl67") & "'"
                      End If
                      If InStr(strB2, "CRL68='" & rsBD.Fields("crl68") & "'") = 0 And "" & rsBD.Fields("crl68") <> "" Then
                         strB2 = strB2 & " OR CRL68='" & rsBD.Fields("crl68") & "'"
                      End If
                      rsBD.MoveNext
                   Loop
                   strB1 = " select crl01,crl07,crl08,crl09,crl10,crl55,crl56,crl65,crl67,crl68" & _
                          " from consultrecordlist, flow003 where crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) " & _
                          " and (" & Mid(strB2, 5) & ") group by crl01,crl07,crl08,crl09,crl10,crl55,crl56,crl65,crl67,crl68 order by crl01 "
                   intB = 1
                   Set rsBD = ClsLawReadRstMsg(intB, strB1)
                   If intB = 1 Then
                      rsBD.MoveFirst
                      strSNo = "" & rsBD.Fields("crl01")
                      Do While Not rsBD.EOF
                         strSql = "Update ConsultRecordList Set CRL65='" & strSNo & "' Where crl01='" & rsBD.Fields("crl01") & "' "
                         cnnConnection.Execute strSql
                         If InStr(strNoList & ",", rsBD.Fields("crl01") & ",") = 0 And "" & rsBD.Fields("crl01") <> "" Then
                            strNoList = strNoList & "," & rsBD.Fields("crl01")
                         End If
                         If InStr(strCrl65List & ",", rsBD.Fields("crl65") & ",") = 0 And "" & rsBD.Fields("crl65") <> "" And "" & rsBD.Fields("crl65") <> strSNo Then
                            strCrl65List = strCrl65List & "," & rsBD.Fields("crl65")
                         End If
                         rsBD.MoveNext
                      Loop
                   End If
                End If
               '排除沒有設定的接洽單
               If strNoList <> "" Then
                  strSql = "Update ConsultRecordList Set CRL65=null Where crl65='" & strSNo & "' and instr('" & Mid(strNoList, 2) & "',crl01)=0 "
                  cnnConnection.Execute strSql, intI
               End If
               '重新整理其他群組
               If strCrl65List <> "" Then
                  strB1 = "select crl65, count(crl01) as cnt, min(crl01) as mno from consultrecordlist, flow003 where crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) " & _
                          " and instr('" & Mid(strCrl65List, 2) & "',crl65) > 0 group by crl65"
                  intB = 1
                  Set rsBD = ClsLawReadRstMsg(intB, strB1)
                  If intB = 1 Then
                     rsBD.MoveFirst
                     Do While Not rsBD.EOF
                        If Val("" & rsBD.Fields("cnt")) = 1 Then
                            strSql = "Update ConsultRecordList Set CRL65=null Where crl65='" & "" & rsBD.Fields("crl65") & "' "
                            cnnConnection.Execute strSql
                        ElseIf Val("" & rsBD.Fields("cnt")) > 1 And "" & rsBD.Fields("mno") <> "" & rsBD.Fields("crl65") Then
                            strSql = "Update ConsultRecordList Set CRL65='" & "" & rsBD.Fields("mno") & "' Where crl65='" & "" & rsBD.Fields("crl65") & "' "
                            cnnConnection.Execute strSql
                        End If
                        rsBD.MoveNext
                     Loop
                  End If
               End If
             Else
                strSql = "Update ConsultRecordList Set CRL65=null Where crl01='" & pCP140 & "' "
                cnnConnection.Execute strSql
                strB1 = "select count(crl01) cnt from ConsultRecordList,flow003 where CRL65='" & pCRL65 & "' " & _
                        " and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) "
                intB = 1
                Set rsBD = ClsLawReadRstMsg(intB, strB1)
                If intB = 1 Then
                   If Val("" & rsBD.Fields("cnt")) = 1 Then
                      strSql = "Update ConsultRecordList Set CRL65=null Where CRL65='" & pCRL65 & "' "
                      cnnConnection.Execute strSql
                   End If
                End If
             End If
          End If
          Set rsBD = Nothing
          'end 2024/04/18
          
          m_CRC09 = Trim(Left(cboCRC09.Text, 6)) '為了SaveGrpData的判斷, 改到最後面
          
          '案件屬性
          If cboAttr.Enabled = True And cboAttr.Text <> m_Attr Then
              strSql = "Update ConsultRecordList set CRL81='" & Left(cboAttr.Text, 1) & "' Where CRL01='" & pCP140 & "' "
              cnnConnection.Execute strSql
              strSql = "Update Patent set PA158='" & Left(cboAttr.Text, 1) & "' WHERE PA01='" & m_CP01 & "' AND PA02='" & m_CP02 & "' AND PA03='" & m_CP03 & "' AND PA04='" & m_CP04 & "' "
              Pub_SeekTbLog strSql
              cnnConnection.Execute strSql
              GRD1.TextMatrix(m_idX, colCAttr_1) = Trim(Left(cboAttr, 1))
              GRD1.TextMatrix(m_idX, colCAttrName_1) = Trim(Mid(cboAttr, 3))
              m_Attr = cboAttr.Text
          End If
          
          '是否算案件數
          If txtCRC10.Enabled = True And txtCRC10.Text <> m_CRC10 Then
              strSql = "Update ConsultRecCMP set CRC10='" & txtCRC10.Text & "' Where CRC01='" & pCP140 & "' AND CRC08='" & pCP09 & "' "
              cnnConnection.Execute strSql
              GRD1.TextMatrix(m_idX, colCRC10_1) = Trim(txtCRC10.Text)
              m_CRC10 = txtCRC10
          End If
          '承辦人計件值
          If txtCRC11.Enabled = True And txtCRC11.Text <> m_CRC11 Then
             strSql = "Update ConsultRecCMP set CRC11='" & txtCRC11.Text & "' Where CRC01='" & pCP140 & "' AND CRC08='" & pCP09 & "' "
             cnnConnection.Execute strSql
             GRD1.TextMatrix(m_idX, colCRC11_1) = Trim(txtCRC11.Text)
             m_CRC11 = txtCRC11
          End If
          '承辦人加乘註記
          If txtCRC12.Enabled = True And txtCRC12.Text <> m_CRC12 Then
             strSql = "Update ConsultRecCMP set CRC12='" & txtCRC12.Text & "' Where CRC01='" & pCP140 & "' AND CRC08='" & pCP09 & "' "
             cnnConnection.Execute strSql
             GRD1.TextMatrix(m_idX, colCRC12_1) = Trim(txtCRC12.Text)
             m_CRC12 = txtCRC12
          End If
          
          '流程=>意見
          If Trim(txtNote) <> "" Then '統一狀態= Flow_待分案，不用傳入F0408,F0409
             strSql = GetInsertFLOW004Sql(pCP140, strUserNum, strSrvDate(1), Format(ServerTime, "000000"), Flow_待分案, ChgSQL(Trim(txtNote.Text)))
             cnnConnection.Execute strSql
          End If
       cnnConnection.CommitTrans
    End If
    
    SaveUpdRow = True
    Exit Function
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "SaveUpdRow(" & IIf(mStatus = "1", "更正", "確定分案") & ")"
        cnnConnection.RollbackTrans
    End If
End Function

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim intRow As Integer, intCol As Integer

   With GRD1
       intCol = .MouseCol
'       If .MouseRow > 0 Then
'          intLastTop = GRD1.TopRow
'       End If

       'If .MouseRow > 0 And blnDBClick = False Then '12/15 改成Ctrl為複選
       If .MouseRow > 0 Then
          intRow = .MouseRow
          .row = intRow
          .col = cFixed + 1 '取固定欄位後的底色，排除因為全選的變色
          lngColor = .CellBackColor
          m_PrevNo = "" & GRD1.TextMatrix(intRow, colCP09_1)
          If InStr(m_SysNo, "P") > 0 And Shift <> 2 Then  ''12/15 改成Ctrl為複選: 專利部沒按Ctrl，皆為單選
             If "" & GRD1.TextMatrix(intRow, colCaseNo_1) <> "" Then
                  If "" & GRD1.TextMatrix(intRow, 0) = "V" Then
                      GoTo JumpToSel01
                  Else
                      Call GetAllSelType(GRD1, "0", intRow)
                  End If
             End If
          Else
JumpToSel01:
             GridClick GRD1, intRow, 0, 0, cFixed, "V", lngColor, colCRL55n_1 & "," & colCRL67n_1
             If "" & GRD1.TextMatrix(intRow, colCaseNo_1) <> "" Then
                  Call GetSelState(intRow)
                  If "" & GRD1.TextMatrix(intRow, 0) = "V" Then
                      If PUB_CheckFormExist("frm090801_Q") = True Then
                          Unload frm090801_Q
                      End If
                  End If
             End If
          End If
       End If
   End With
   
   'blnDBClick = False
   
'-------------------------------------------
'保留：為了不弄亂群組，先不開放此功能
'Dim nCol As Long, nRow As Long
'   getGrdColRow GRD1, X, Y, nCol, nRow
'   If nCol < 0 Or nRow < 0 Then Exit Sub
'   GRD1.col = nCol
'   GRD1.row = nRow
'   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
'      If mDisCount > 0 Then
'          MsgBox "案件收合狀態不可重新排序！", vbInformation
'          Exit Sub
'      End If
'      If InStr("相似度", Me.GRD1.Text) > 0 Then  '保留
'         If m_blnColOrderAsc1 = True Then
'            Me.GRD1.Sort = 3  '數值昇冪
'            m_blnColOrderAsc1 = False
'         Else
'            Me.GRD1.Sort = 4 '數值降冪
'            m_blnColOrderAsc1 = True
'         End If
'      Else
'         If m_blnColOrderAsc1 = True Then
'            Me.GRD1.Sort = 5 '字串昇冪
'            m_blnColOrderAsc1 = False
'         Else
'            Me.GRD1.Sort = 6 '字串降冪
'            m_blnColOrderAsc1 = True
'         End If
'      End If
'   End If
End Sub

Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim intRow As Integer, intCol As Integer

   With grd2
       intCol = .MouseCol
       If .MouseRow > 0 Then
          intRow = .MouseRow
          .row = intRow
          .col = cFixed + 1 '取固定欄位後的底色，排除因為全選的變色
          lngColor = .CellBackColor
          If InStr(m_SysNo, "P") > 0 And Shift <> 2 Then  '專利部沒按Ctrl，皆為單選
             If "" & grd2.TextMatrix(intRow, colCaseNo_1) <> "" Then
                  If "" & grd2.TextMatrix(intRow, 0) = "V" Then
                      GoTo JumpToSel02
                  Else
                      Call GetAllSelType(grd2, "0", intRow)
                  End If
             End If
          Else
JumpToSel02:
             GridClick grd2, intRow, 0, 0, cFixed, "V", lngColor, colCRL55n_1 & "," & colCRL67n_1
             If "" & grd2.TextMatrix(intRow, colCaseNo_1) <> "" Then
                  If "" & grd2.TextMatrix(intRow, 0) = "V" Then
                      If PUB_CheckFormExist("frm090801_Q") = True Then
                          Unload frm090801_Q
                      End If
                  End If
             End If
          End If
       End If
   End With

'-------------------------------------------
'保留：為了不弄亂群組，先不開放此功能
'Dim nCol As Long, nRow As Long
'
'   getGrdColRow GRD2, x, y, nCol, nRow
'   If nCol < 0 Or nRow < 0 Then Exit Sub
'   GRD2.col = nCol
'   GRD2.row = nRow
'   If Me.GRD2.row < 1 And Me.GRD2.Text <> "V" Then
'      If InStr("相似度", Me.GRD2.Text) > 0 Then  '保留
'         If m_blnColOrderAsc2 = True Then
'            Me.GRD2.Sort = 3  '數值昇冪
'            m_blnColOrderAsc2 = False
'         Else
'            Me.GRD2.Sort = 4 '數值降冪
'            m_blnColOrderAsc2 = True
'         End If
'      Else
'         If m_blnColOrderAsc2 = True Then
'            Me.GRD2.Sort = 5 '字串昇冪
'            m_blnColOrderAsc2 = False
'         Else
'            Me.GRD2.Sort = 6 '字串降冪
'            m_blnColOrderAsc2 = True
'         End If
'      End If
'   End If
End Sub

Private Function SaveToBack(ByVal mStatus As String) As Boolean
Dim strContent As String, strSubject As String
    
    SaveToBack = False
    If Trim(txtNote) = "" Then
        MsgBox "您的意見不可空白！", vbInformation
        txtNote.SetFocus
        txtNote_GotFocus
        Exit Function
    End If
    
    strUpdTime = Format(ServerTime, "000000")
    
On Error GoTo ErrHandle
    cnnConnection.BeginTrans

       '簽核檔: F0207=簽核結果6,7
       If m_nowF0308 = "A6" Then '一併處理分所
          strSql = "update FLOW002 set F0205='" & strSrvDate(1) & "' ,F0206='" & strUpdTime & "',F0207=" & CNULL(IIf(mStatus = "2", "6", "7")) & ", F0204='" & strUserNum & "' " & _
                  " where F0201='" & m_CP140 & "' and F0202='A5' and F0207 is null "
          cnnConnection.Execute strSql
       End If
       strSql = "update FLOW002 set F0205='" & strSrvDate(1) & "' ,F0206='" & strUpdTime & "',F0207=" & CNULL(IIf(mStatus = "2", "6", "7")) & ", F0204='" & strUserNum & "' " & _
               " where F0201='" & m_CP140 & "' and F0202='" & m_nowF0308 & "' and F0207 is null "
       cnnConnection.Execute strSql
   
       '表單主檔: 2=智權, 3=程序
       strSql = "update FLOW003 set " & _
                "F0307='" & m_nowF0308 & "' " & _
                ",F0308='" & IIf(mStatus = "2", m_F0316, "A7") & "'" & _
                ",F0309='" & IIf(mStatus = "2", Flow_智權補件, Flow_程序補件) & "'" & _
                " where F0301='" & m_CP140 & "' And F0302='" & m_F0302 & "' "
       cnnConnection.Execute strSql
   
       '流程備註檔
       If Trim(txtNote.Text) <> "" Then
          strSql = GetInsertFLOW004Sql(m_CP140, strUserNum, strSrvDate(1), strUpdTime, _
                     IIf(mStatus = "2", Flow_智權補件, Flow_程序補件), ChgSQL(Trim(txtNote.Text)), _
                     m_nowF0308, IIf(mStatus = "2", m_F0316, "A7"))
          cnnConnection.Execute strSql
       End If
       
       '程序+智權補件：再新增待簽核的記錄，只需新增自己所別
       Call SetConultRecPrePerson_Flow002(Me.Name, m_CP140, IIf(mStatus = "2", "A0", "A7"), m_F0316) '給智權人員/程序
       Call SetConultRecPrePerson_Flow002(Me.Name, m_CP140, m_nowF0308)  '給所別
       
    cnnConnection.CommitTrans
    
       '發E-Mail通知當事人: 退程序=A7不用發email通知由人工自行查看，退智權發email
       If mStatus = "2" Then
          strContent = GetEMailContent_Flow(m_CP140, strSubject)
          If Trim(txtNote.Text) <> "" Then
             'Modified by Lydia 2023/05/03 "；退回原因：" =>"；原因："; 智權人員會誤解意思by 秀玲
             strSubject = strSubject & "；原因：" & Trim(txtNote.Text)
          End If
          PUB_SendMail strUserNum, m_F0316, "", strSubject, strContent
       End If
    
    SaveToBack = True
    Exit Function
    
ErrHandle:
    If Err.Number <> 0 Then
       cnnConnection.RollbackTrans
       MsgBox Err.Description, vbCritical, IIf(mStatus = "2", "智權補件", "程序補件") & "更新失敗"
    End If
End Function

'取得/釋放資料異動權 p_NewKey:欲取得資料鍵值
Private Function ProcGetLock(ByVal p_NewKey As String) As Boolean

On Error GoTo ErrHand
   
   If Left(p_NewKey, 1) = "A" Then
      strSql = "Delete from LockRec where LR02='" & strUserNum & "' and LR01 like '" & Mid(p_NewKey, 2) & "%' "
      cnnConnection.Execute strSql
   Else
      strSql = "Delete from LockRec where LR02 = '" & strUserNum & "' and LR01 like '" & Mid(p_NewKey, 1, Len(p_NewKey) - 1) & "%' "
      cnnConnection.Execute strSql
      strQuery = "select st02 from LockRec,staff where LR01 like '" & p_NewKey & "%' and st01(+)=LR02 and st01 <> '" & strUserNum & "' "
      intQ = 1
      Set rsAD1 = ClsLawReadRstMsg(intQ, strQuery)
      If intQ = 1 Then
          strSql = "" & rsAD1.GetString(adClipString, , , ",")
          MsgBox "【" & Mid(strSql, 1, Len(strSql) - 1) & "】同時使用【" & IIf(Left(m_SysNo, 1) = "P", "專利", "商標") & "】之" & IIf(Left(m_SysNo, 1) = "P", Mid(Combo1.Text, 4, 2), "") & "分案！", vbInformation
      End If
      strSql = "Insert into LockRec(LR01,LR02,LR03) values ('" & p_NewKey & "-" & strUserNum & "','" & strUserNum & "',to_char(sysdate,'YYYYMMDDHH24MISS'))"
      cnnConnection.Execute strSql
   End If
   ProcGetLock = True
   Exit Function
   
ErrHand:
   If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "取得/釋放資料異動"
   End If
End Function

Private Function SaveToCP14(ByVal pIdx As Integer, ByVal pCP140 As String, ByVal pCP09 As String, ByVal pF0316 As String, ByVal pF0309 As String) As Boolean
Dim NowCP14 As String, NowCRC10 As String, NowCRC11 As String, NowCRC12 As String
Dim bolConn As Boolean
Dim updF0308 As String

On Error GoTo ErrHandle
     
    '先完成更正，再判斷確定分案;　商標可複選，需要一併更正
    If (m_SysNo <> "T" And pCP09 = m_CP09) Or (m_SysNo = "T" And cboCRC09 <> "") Then
        m_CRC09 = "" '強制更新CP14
        'Modified by Lydia 2023/07/26 商標案承辦人需為空白才會一併更正; T133524延展和TS-001968查名逐一輸入承辦人和按下更正，之後兩筆一起送出，最後承辦人都掛TS案
        'If SaveUpdRow("2", pCP140, pCP09) = False Then
        If SaveUpdRow("2", pCP140, pCP09, IIf(m_SysNo = "T", True, False)) = False Then
            GoTo ErrHandle
        End If
    End If
            
    '抓目前DB的預分承辦人
    If pCP09 = m_CP09 Then
        NowCP14 = Trim(Left(cboCRC09, 6))
        NowCRC10 = txtCRC10
        NowCRC11 = txtCRC11
        NowCRC12 = txtCRC12
    Else
        strSql = "SELECT CRC09,ST02,CRL06,CRL07,DECODE(CRL87,'3', decode(CRL81,'1','整體','2','部分','3','圖像','4','組成',CRL81 ),decode(CRL81,'1','機械','2','電子','3','化學生醫',CRL81 )) as CAttr," & _
                    "DECODE(CRL56,'Y',CRL55, NULL) CRL55,CRL74,CRL67,CRL68,CRC10,CRC11,CRC12 FROM ConsultRecCMP,ConsultRecordList,Staff " & _
                    "where crc01='" & pCP140 & "' and crc08='" & pCP09 & "' and crc01=crl01 and crc09=st01(+)"
        intQ = 1
        Set rsAD1 = ClsLawReadRstMsg(intQ, strSql)
        If intQ = 1 Then
            NowCP14 = "" & rsAD1.Fields("crc09")
            NowCRC10 = "" & rsAD1.Fields("crc10")
            NowCRC11 = "" & rsAD1.Fields("crc11")
            NowCRC12 = "" & rsAD1.Fields("crc12")
            '更新Grid
            If NowCP14 <> "" And NowCP14 <> GRD1.TextMatrix(pIdx, colCRC09_1) Then
               GRD1.TextMatrix(pIdx, colCRC09_1) = NowCP14
               GRD1.TextMatrix(pIdx, colCRC09n_1) = "" & rsAD1.Fields("ST02")
            End If
            If NowCRC10 <> "" And NowCRC10 <> GRD1.TextMatrix(pIdx, colCRC10_1) Then
               GRD1.TextMatrix(pIdx, colCRC10_1) = NowCRC10
            End If
            If NowCRC11 <> "" And NowCRC11 <> GRD1.TextMatrix(pIdx, colCRC11_1) Then
               GRD1.TextMatrix(pIdx, colCRC11_1) = NowCRC11
            End If
            If NowCRC12 <> "" And NowCRC12 <> GRD1.TextMatrix(pIdx, colCRC12_1) Then
               GRD1.TextMatrix(pIdx, colCRC12_1) = NowCRC12
            End If
            '非新案, 商標案不用
            If InStr("," & rsAD1.Fields("CRL07"), "T") = 0 And "" & rsAD1.Fields("CRL06") = "Y" Then
               If "" & rsAD1.Fields("CRL55") <> GRD1.TextMatrix(pIdx, colCRL55_1) Then
                   GRD1.TextMatrix(pIdx, colCRL55_1) = "" & rsAD1.Fields("CRL55")
               End If
               If "" & rsAD1.Fields("CRL67") <> GRD1.TextMatrix(pIdx, colCRL67_1) Then
                   GRD1.TextMatrix(pIdx, colCRL67_1) = "" & rsAD1.Fields("CRL67")
               End If
               If "" & rsAD1.Fields("CRL68") <> GRD1.TextMatrix(pIdx, colCRL68_1) Then
                   GRD1.TextMatrix(pIdx, colCRL68_1) = "" & rsAD1.Fields("CRL68")
               End If
               If "" & rsAD1.Fields("CRL07") = "P" Or "" & rsAD1.Fields("CRL07") = "CFP" Then
                   GRD1.TextMatrix(pIdx, colCAttrName_1) = "" & rsAD1.Fields("CATTR")
               End If
            End If
        End If
    End If
    If m_nowF0308 = "A6" And NowCP14 = "" Then
        MsgBox GRD1.TextMatrix(pIdx, colCaseNo_1) & "尚未分承辦人，不可確定分案！", vbInformation, "分案檢查"
        Call GetCurrRow(GRD1, m_idX, 1)
        If m_idX > 0 Then
            Call UpdateCtrlData(m_idX)
        Else
            Call ClearCtrlData
        End If
        Exit Function
    End If
    
    '檢查整張接洽單是否全部分案
    If m_nowF0308 = "A6" Then
      'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉and nvl(cpm35,'0')<>'2'
      strSql = "SELECT CRC08 FROM ConsultRecCMP,ConsultRecordList,CasePropertyMap where crc01='" & pCP140 & "' and crc01=crl01(+) " & _
                  "and crc09 is null and crl07=cpm01(+) and crc03=cpm02(+) "
      intQ = 1
      Set rsAD1 = ClsLawReadRstMsg(intQ, strSql)
      If intQ = 1 Then
          MsgBox GRD1.TextMatrix(pIdx, colCaseNo_1) & "尚未分承辦人，不可確定分案！", vbInformation, "分案檢查"
            Call GetCurrRow(GRD1, m_idX, 1)
          If m_idX > 0 Then
              Call UpdateCtrlData(m_idX)
          Else
              Call ClearCtrlData
          End If
          Exit Function
       End If
    End If

    '因為分所人員可以直接交由北所分案，所以判斷分所人員不預分承辦人也可分案
    If (m_nowF0308 = "A6" And NowCP14 <> "") Or m_nowF0308 = "A5" Then
        strUpdTime = Format(ServerTime, "000000")
        cnnConnection.BeginTrans
        bolConn = True
            'Modified by Lydia 2025/10/20 改成判斷權限
            'If m_ST06 = "1" Then  'Added by Lydia 2023/01/06 專利案分所主管分案先不上進度檔的承辦人，要等到北所主管執行「確定分案」
            If m_nowF0308 = "A6" Then
                'Modified by Lydia 2023/01/17 debug: CFP-32906 同時收422+417,承辦人分別設工程師和程序
                'strSql = "Update CaseProgress set CP14=(select crc09 from ConsultRecCMP where crc01='" & pCP140 & "' and crc08='" & pCP09 & "' ) where cp09='" & pCP09 & "' and cp14 is null "
                'Modified by Lydia 2023/05/19 一律以主管分案為準AND CP159=0 AND CP14 IS NULL => AND CP159=0; ex.P-131540(AB2018840)因為主管指示程序退回重新分案換工程師，但是沒有回寫新的承辦人，經過與Sindy討論決定都以主管為主
                strSql = "Update CaseProgress set CP14=(select crc09 from ConsultRecordList,ConsultRecCMP,CasePropertyMap " & _
                            "where CRL01='" & pCP140 & "' AND CRL01=CRC01(+) AND CRL07=CPM01(+) AND CRC03=CPM02(+) and crc08=cp09 and crc09 is not null) " & _
                            "where cp140='" & pCP140 & "' AND CP158=0 AND CP159=0 "
                cnnConnection.Execute strSql
                'Added by Lydia 2025/10/22 T案收308分割會直接用同一接洽單另外產生新案；T分割的接洽單一般都會預分承辦人,但現在若承辦人請假時, 接洽單就會進入主管分案作業,當主管針對母案操作分案時, 請一併把子案也掛上同承辦人。
                                          'Ex.母案: T-250772 子案: T-256736,T-256737
                If m_SysNo = "T" Then
                   strSql = "UPDATE caseprogress SET cp14=(SELECT crc09  FROM consultreccmp WHERE crc01='" & pCP140 & "' AND crc03='308' AND nvl(crc09,'N')<>'N') " & _
                            " WHERE cp01='T' and cp14 IS NULL AND cp140=(SELECT crc01 FROM consultreccmp WHERE crc01='" & pCP140 & "' AND crc03='308' AND nvl(crc09,'N')<>'N') "
                   cnnConnection.Execute strSql
                End If
                'end 2025/10/22
            End If
            strQuery = ""
            If m_nowF0308 = "A6" And NowCP14 <> "" Then '北所主管分案
              '是否算案件數CRC10=>CP26
               If Trim(NowCRC10) <> "" Then strQuery = strQuery & ", CP26=" & CNULL(NowCRC10)
               '承辦人計件值CRC11=>CP97
               If Trim(NowCRC11) <> "" Then
                  If ExistCheck("EXVALUE", "EV01", pCP09, strExc(0), False) = True Then
                      strSql = "UPDATE EXVALUE Set EV02=" & Val(NowCRC11) & " Where EV01='" & pCP09 & "' "
                  Else
                      strSql = "INSERT INTO EXVALUE (EV01,EV02) VALUES ('" & pCP09 & "'," & Val(NowCRC11) & ") "
                  End If
                  Pub_SeekTbLog strSql '記錄修改log
                  cnnConnection.Execute strSql
                  strQuery = strQuery & ", CP97=" & CNULL(NowCRC11, True)
               End If
               '承辦人加乘註記CRC12=>CP98
               If Trim(NowCRC12) <> "" Then
                  strQuery = strQuery & ", CP98=" & CNULL(NowCRC12, True) & ", CP99=TO_CHAR(SYSDATE,'YYYYMMDD')||" & CNULL("主管分案(" & strUserNum & ");") & "||CP99"
               End If
               If strQuery <> "" Then
                   strSql = "Update CaseProgress Set " & Mid(strQuery, 2) & " Where CP09=" & CNULL(pCP09)
                   Pub_SeekTbLog strSql '記錄修改log
                   cnnConnection.Execute strSql
               End If
            End If '北所主管分案
            
            '簽核檔
            If m_nowF0308 = "A6" Then '一併處理分所
                strSql = "update FLOW002 set F0205='" & strSrvDate(1) & "' ,F0206='" & strUpdTime & "',F0207='1', F0204='" & strUserNum & "' " & _
                            " where F0201='" & pCP140 & "' and F0202='A5' and F0207 is null "
                cnnConnection.Execute strSql
            End If
            strSql = "update FLOW002 set F0205='" & strSrvDate(1) & "' ,F0206='" & strUpdTime & "',F0207='1', F0204='" & strUserNum & "' " & _
                        " where F0201='" & pCP140 & "' and F0202='" & m_nowF0308 & "' and F0207 is null "
            cnnConnection.Execute strSql
            '讀取下一處理人員updF0308, 而不是變更操作人員m_nowF0308
            If GetNextProPerson_Flow(pCP140, pF0316, updF0308, pF0309) = False Then GoTo ErrHandle
        cnnConnection.CommitTrans
    End If
    SaveToCP14 = True
    Exit Function
    
ErrHandle:
    If Err.Number <> 0 Then
        If bolConn = True Then cnnConnection.RollbackTrans
        MsgBox Err.Description, vbCritical, "分案失敗"
    End If

End Function

Private Sub GetSelState(ByVal pIdx As Integer)
'勾選單筆收文提供「更正、智權補件、程序補件、確定分案」的功能；
'複選收文僅提供「確定分案」的功能，其他"更正、智權補件、程序補件"按鈕不可點選。
     
   If "" & GRD1.TextMatrix(pIdx, 0) = "V" Then
       mSelCount = mSelCount + 1
   ElseIf mSelCount > 0 Then
       mSelCount = mSelCount - 1
   End If
   
   '專利主管分案（複選不可執行更正和通知補件）
   If (mSelCount > 1 Or (mSelCount > 0 And mDisplay = 1 And mDisCount > 0)) And InStr(m_SysNo, "P") > 0 Then
       'Added by Lydia 2023/01/04 北所主管可以對全所＋各分所做「確定分案」，分所可以看他所案件但僅可以「更正」不可分案(from 副總)
'''       If InStr(m_SysNo, "P") > 0 And m_ST06 <> "1" And Left(Combo1, 1) <> m_ST06 Then
'''           cmdOK(1).Enabled = False
'''       Else
'''           cmdOK(1).Enabled = True
'''       End If
'end 2023/01/04
       cmdUpdRow.Enabled = False
       cmdOK(2).Enabled = False
       cmdOK(3).Enabled = False
       Call ShowObject(False)
       If mSelCount = 2 Then
           Call ClearCtrlData
       End If
       Frame1.Enabled = False '有勾選一筆才可以輸入維護區的資料，為了避免誤輸，若未勾選或複選直接關閉輸入。
       lblData(2).Visible = False: lblData(3).Visible = False
   Else
       If mSelCount = 1 Then
          'Modified by Lydia 2023/01/04 北所主管可以對全所＋各分所做「確定分案」，分所可以看他所案件但僅可以「更正」不可分案(from 副總)
             'cmdUpdRow.Enabled = True
             'cmdOK(1).Enabled = True
             'cmdOK(2).Enabled = True
             'cmdOK(3).Enabled = True
          'Modified by Lydia 2025/08/18 林柄佑協理有權限對對全所＋各分所做「確定分案」
          'If InStr(m_SysNo, "P") > 0 And m_ST06 <> "1" And Left(Combo1, 1) <> m_ST06 Then
          If InStr(m_SysNo, "P") > 0 And m_ST06 <> "1" And Left(Combo1, 1) <> m_ST06 And InStr(m_Str專利處台北區主管, strUserNum) = 0 Then
             cmdOK(1).Enabled = False
          Else
             cmdOK(1).Enabled = True
          End If
          cmdUpdRow.Enabled = True
          cmdOK(2).Enabled = True
          cmdOK(3).Enabled = True
          'end 2023/01/04
          Call GetCurrRow(GRD1, m_idX, 1)
          If m_idX > 0 Then
            Frame1.Enabled = True '先開放，不然在欄位設定無法判斷
            Call UpdateCtrlData(m_idX)
          End If
          lblData(2).Visible = True: lblData(3).Visible = True
       ElseIf mSelCount = 0 Then
          cmdUpdRow.Enabled = False
          cmdOK(2).Enabled = False
          cmdOK(3).Enabled = False
          Frame1.Enabled = False '有勾選一筆才可以輸入維護區的資料，為了避免誤輸，若未勾選或複選直接關閉輸入。
          lblData(2).Visible = False: lblData(3).Visible = False
       End If
   End If

End Sub

Private Sub SetCboCRL(ByRef pCmb As ComboBox, ByVal pCaseNo As String)
Dim intA As Integer
Dim arrTmp As Variant
    
    pCmb.Clear
    If pCmb.Enabled = True Then
        If UCase(pCmb.Name) = "CBOCRL67" And InStr("000,020,231", m_Na01) > 0 Then
           '預設符合主管分案+同一系統+同一申請國+同一申請人+101,102的案件;
           strExc(0) = "SELECT PA01||'-'||PA02||DECODE(PA03||PA04,'000',NULL,'-'||PA03||'-'||PA04) AS CASENO FROM CONSULTRECORDLIST,CONSULTRECCMP,PATENT " & _
                           "WHERE CRL01 IN (SELECT CRL01 FROM CONSULTRECORDLIST,FLOW003 WHERE CRL01=F0301 AND F0309 IN ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "') " & _
                           ") AND CRL01=CRC01(+) AND CRL07='" & m_CP01 & "' AND CRL08||CRL09||CRL10<>'" & m_CP02 & m_CP03 & m_CP04 & "' AND CRL15='" & m_Na01 & "' AND CRC03='" & IIf(m_CP10 = "101", "102", "101") & "' AND CRL01=CRC01(+) " & _
                           "AND CRL07=PA01(+) AND CRL08=PA02(+) AND CRL09=PA03(+) AND CRL10=PA04(+) AND PA26='" & m_APP01 & "' ORDER BY 1 "
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
              RsTemp.MoveFirst
              Do While Not RsTemp.EOF
                 pCmb.AddItem "" & RsTemp.Fields("caseno")
                 RsTemp.MoveNext
              Loop
           End If
        Else
           If UCase(pCmb.Name) = "CBOCRL55" Then
               If m_CRL55 = m_CRL67 And m_CRL55 <> "" Then
                   pCaseNo = ""
               End If
           End If
           If pCaseNo <> "" And pCaseNo <> lblData(0) Then
               arrTmp = Split(pCaseNo, ",")
               For intA = 0 To UBound(arrTmp)
                   If Trim(arrTmp(intA)) <> "" Then
                       pCmb.AddItem Trim(arrTmp(intA))
                   End If
               Next intA
           End If
        End If
        pCmb.Text = pCaseNo
    End If
End Sub

Private Sub ShowObject(ByVal bolShow As Boolean)
    '商標預設不顯示”相同案號、一案兩請....”
    If m_SysNo <> "T" Then
        Label1(11).Visible = bolShow: Label1(3).Visible = bolShow: Label1(4).Visible = bolShow: Label1(5).Visible = bolShow
        cboCRL55.Visible = bolShow: cboAttr.Visible = bolShow: cboCRL67.Visible = bolShow: cboCRL68.Visible = bolShow
        Check1.Visible = bolShow
    End If
End Sub

'針對Grid的選取控制
Private Sub GetAllSelType(ByRef gGrid As MSHFlexGrid, ByVal gType As String, Optional ByVal gIDX As Integer = 0)
'gType:  0=取消, 1=選取
'gIDX: 指定選取
Dim intX As Integer, intY As Integer

     gGrid.Visible = False
     If UCase(gGrid.Name) = "GRD1" Then mSelCount = 0
     If gGrid.Rows > 1 Then
        For intX = 1 To gGrid.Rows - 1
           gGrid.row = intX
           gGrid.col = cFixed + 1  '取固定欄位後的底色，排除因為全選的變色
           lngColor = gGrid.CellBackColor
           gGrid.col = 0
           gGrid.row = intX
           If gType = "1" Or (gType = "0" And intX = gIDX) Then  '全選
               gGrid.Text = "V"
               If UCase(gGrid.Name) = "GRD1" Then
                  Call GetSelState(intX)
               End If
               For intY = 0 To cFixed - 1
                   If InStr("," & colCRL55n_1 & "," & colCRL67n_1 & ",", intY) = 0 Then '排除相同案號
                      gGrid.col = intY
                      gGrid.CellBackColor = &HFFC0C0
                   End If
               Next intY
           ElseIf gType = "0" Then '取消
               gGrid.Text = ""
               For intY = 0 To cFixed - 1
                   If InStr("," & colCRL55n_1 & "," & colCRL67n_1 & ",", intY) = 0 Then '排除相同案號
                      gGrid.col = intY
                      gGrid.CellBackColor = lngColor
                   End If
               Next intY
           End If
        Next intX
     End If
     gGrid.Visible = True
        
     'Added by Lydia 2025/04/15 檢查是否為規費調整接洽單，彈訊息"為規費調整接洽單！"警示，規費有文字就不算CRC13=Y
     If "" & gGrid.TextMatrix(gIDX, 0) = "V" And "" & gGrid.TextMatrix(gIDX, colMsg1) = "Y" Then
         MsgBox gGrid.TextMatrix(gIDX, colCaseNo_1) & "為規費調整接洽單！", vbInformation + vbOKOnly
     End If
     'end 2025/04/15
End Sub

Private Sub txtCRC11_GotFocus()
    TextInverse txtCRC11
End Sub

Private Sub txtCRC11_Validate(Cancel As Boolean)

   If Trim(txtCRC11) = "" Then Exit Sub
   
      If Not IsNumeric(txtCRC11) Then
        MsgBox "承辦人計件值輸入錯誤！", vbExclamation
        Cancel = True
      End If

End Sub

Private Sub txtCRC12_GotFocus()
   TextInverse txtCRC12
End Sub

Private Sub txtCRC12_Validate(Cancel As Boolean)
   Dim iMax As Integer
   iMax = 3

   If Trim(txtCRC12) = "" Then Exit Sub
   
   If bolNewPromoterRule Then
      iMax = 9
   End If
   
   If Not IsNumeric(txtCRC12) Then
      MsgBox "資料輸入錯誤！", vbExclamation
      Cancel = True
   ElseIf Val(txtCRC12) > iMax Then
      MsgBox "資料輸入錯誤！", vbExclamation
      Cancel = True
   End If
End Sub

Private Sub cboCRC09_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboCRC09_Validate(Cancel As Boolean)
Dim intX As Integer
   
   intX = -1
   If Trim(cboCRC09.Text) <> "" And cboCRC09.Tag <> cboCRC09.Text Then
         For intQ = 0 To cboCRC09.ListCount - 1
             If InStr(cboCRC09.List(intQ), Trim(cboCRC09.Text)) > 0 Then
                 intX = intQ
                 Exit For
             End If
        Next intQ
        If intX = -1 Then
             If ByInputGetST01or02(Trim(Left(cboCRC09.Text, 6)), strQuery, strQ2) = False Then
                 cboCRC09.SetFocus
                 cboCRC09.Tag = cboCRC09.Text
                 Cancel = True
                 Exit Sub
             Else
                 cboCRC09.Text = convForm(strQuery, 6) & "  " & strQ2
             End If
        Else
             cboCRC09.ListIndex = intX
        End If
   End If
   cboCRC09.Tag = cboCRC09.Text
End Sub

'Added by Lydia 2023/01/06 顯示共同查詢之進度檔
Private Function ShowCPgrid(ByVal pCaseNo As String) As Boolean
    
    ShowCPgrid = False
    If Trim(pCaseNo) <> "" Then
        strBCase(0) = Replace(Trim(pCaseNo), "-", "")
        Call ChgCaseNo(strBCase(0), strBCase)
        If strBCase(1) <> "" And strBCase(2) <> "" Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Function
            End If
            strBCase(0) = strBCase(1) & "-" & strBCase(2) & "-" & strBCase(3) & "-" & strBCase(4)
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = strBCase(0)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            ShowCPgrid = True
        End If
    End If
End Function

'Added by Lydia 2023/02/18 複數母案
Private Function SaveGrpData(ByVal pStatus As String, ByVal pEnabled As Boolean, ByVal pCP140 As String, ByVal pTName As String, ByRef pNewData As String, ByRef pOldData As String, ByVal pDefCP14 As String, ByVal pGrpNo As String, ByVal pCRL67 As String) As Boolean
'pStatus: 1-更正，2-確認分案
'pEnabled: 物件可維護 (不限制物件可維護，因為有可能在非新案或子案的收文上承辦人)
'pTName: 物件名稱
'pNewData=更新後；　pOldData=更新前
Dim strCon1 As String
Dim intG As Integer
Dim tmpArrNew As Variant
Dim strNCase(0 To 4) As String

    SaveGrpData = False
    
On Error GoTo ErrHandle
    
    'Added by Lydia 2024/04/18
    Dim strDataP As String
    strDataP = pOldData
    'end 2024/04/18
    If pEnabled = False Or pNewData = pOldData Or (pOldData = lblData(0).Caption And pNewData = "") Then
       '關聯沒有變更
    Else
       
'--------變更Key值：原本為母案顯示為空白 + 子案變更母號
       'If (pOldData = lblData(0).Caption And pNewData <> "") Or (pOldData <> lblData(0).Caption And pNewData <> "") Then
       If pNewData <> "" Then
           strBCase(0) = pNewData
           If InStr(pNewData, ",") = 0 Then
               Call ChgCaseNo(Replace(pNewData, "-", ""), strBCase)
               strCon1 = pTName & "=" & CNULL(strBCase(1) & "-" & strBCase(2) & IIf(strBCase(3) & strBCase(4) <> "000", "-" & strBCase(3) & "-" & strBCase(4), "")) & IIf(pTName = "CRL55", ",CRL56='Y'", "")
           Else
               strCon1 = pTName & "=" & CNULL(strBCase(0)) & IIf(pTName = "CRL55", ",CRL56='Y'", "")
           End If
           tmpArrNew = Empty
           tmpArrNew = Split(strBCase(0), ",")
           strSql = "Update ConsultRecordList set " & strCon1 & " Where CRL01='" & pCP140 & "' "
           cnnConnection.Execute strSql
           
           For intG = 0 To UBound(tmpArrNew) '統一用迴圈
               strNCase(0) = Trim(tmpArrNew(intG))
               Call ChgCaseNo(Replace(strNCase(0), "-", ""), strNCase)
               '更新母案
               'Modified by Lydia 2023/01/30 排除母案為其他案之子案 and crl55 is null
               'Modified by Lydia 2023/10/26 調整:同時有相同案號+一案兩請的關聯 ex.P132430+P132431
               'strSql = "Update ConsultRecordList SET " & strCon1 & _
                           " where CRL01 IN (select crl01 from ConsultRecordList,flow003 where crl07='" & strNCase(1) & "' and crl08='" & strNCase(2) & "' and crl09='" & strNCase(3) & "' and crl10='" & strNCase(4) & "' " & _
                           "and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' )) and crl55 is null "
               strSql = "Update ConsultRecordList SET " & strCon1 & _
                           " where CRL01 IN (select crl01 from ConsultRecordList,flow003 where crl07='" & strNCase(1) & "' and crl08='" & strNCase(2) & "' and crl09='" & strNCase(3) & "' and crl10='" & strNCase(4) & "' " & _
                           "and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' )) and " & pTName & " is null "
               cnnConnection.Execute strSql
               '母案變為子案
               If pOldData = lblData(0).Caption And pNewData <> "" Then
                  strSql = "Update ConsultRecordList SET " & strCon1 & _
                              " where CRL01 IN (select crl01 from ConsultRecordList,flow003 where " & pTName & "='" & pOldData & "' " & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                              "and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "')) "
                  cnnConnection.Execute strSql
               End If
               
               'CRL55更新關連代號CRL65
               If pTName = "CRL55" Then
                   '變更相同案號若同時有一案兩請，同時變更另一案件CRL55
                   If pNewData <> "" And pCRL67 <> "" Then
                       'Modified by Lydia 2023/01/30 排除母案為其他案之子案 and crl55 is null
                       strSql = "Update ConsultRecordList SET " & strCon1 & _
                                   " where CRL01 IN (select crl01 from ConsultRecordList,flow003 where crl67='" & pCRL67 & "' " & _
                                   "and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "')) and crl55 is null "
                       cnnConnection.Execute strSql
                       'Move by Lydia 2024/04/18 移到下面 >>檢查沒有子號，母號一併清除設定
                   End If   'If pNewData <> "" And pCrl67 <> "" Then
                   'Mark by Lydia 2024/04/18 同一組(相同案號+一案兩請+擬制喪失新穎性)的所有案件一定要全部在一起; 以關聯代號為主做排序
                   ''統一CRL65的最小流水號
                   'strQuery = "select min(nvl(crl65,crl01)) minno from ConsultRecordList,flow003 where " & pTName & "=" & CNULL(pNewData) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                   '         " and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) "
                   'intP = 1
                   'Set rsAD1 = ClsLawReadRstMsg(intP, strQuery)
                   'If intP = 1 Then
                   '   strSql = "Update ConsultRecordList Set CRL65='" & rsAD1.Fields("minno") & "' Where CRL01 IN (select CRL01 from ConsultRecordList,flow003 where " & pTName & "=" & CNULL(pNewData) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                   '         " and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' )) "
                   '   cnnConnection.Execute strSql
                   'End If '統一CRL65的最小流水號
                   'end 2024/04/18
               End If
               'Move by Lydia 2024/04/18 移到下面
               '檢查沒有子號，母號一併清除設定
               If pOldData <> "" Then
                   Call SaveGrpDataChk(pStatus, pEnabled, pCP140, pTName, pNewData, pOldData, pDefCP14, pGrpNo, pCRL67)
               End If
           Next intG 'For intG = 0 To UBound(tmpArrNew) '統一用迴圈
       End If
       
'--------刪除Key值: 一案兩請同時有多國案號,在多國案號欄位內不顯示一案兩請
       If (pTName = "CRL55" And pNewData = "" And pOldData <> lblData(0).Caption And pOldData <> pCRL67) Or (pTName <> "CRL55" And pNewData = "" And pOldData <> lblData(0).Caption) Then
           'CRL55更新關連代號CRL65
           strSql = "Update ConsultRecordList set " & pTName & "=null " & IIf(pTName = "CRL55", ",CRL56=NULL, CRL65=NULL", "") & " Where CRL01='" & pCP140 & "' "
           cnnConnection.Execute strSql
           '檢查沒有子號，母號一併清除設定
           'Modified by Lydia 2024/04/18
           'If pOldData <> "" Then
           If Not (pTName = "CRL55" And pNewData = "" And pCRL67 <> "" And Check1.Value = 1) Then
               Call SaveGrpDataChk(pStatus, pEnabled, pCP140, pTName, pNewData, pOldData, pDefCP14, pGrpNo, pCRL67)
           End If 'mark by Lydia 2024/04/18
           
           'CRL55更新關連代號CRL65=>檢查沒有子號，母號一併清除設定
           'Mark by Lydia 2024/04/18 同一組(相同案號+一案兩請+擬制喪失新穎性)的所有案件一定要全部在一起; 以關聯代號為主做排序
           'If pTName = "CRL55" Then
           '    strQuery = "select count(crl01) cnt from ConsultRecordList,flow003 where CRL65='" & pGrpNo & "' " & _
           '             " and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) "
           '    intP = 1
           '    Set rsAD1 = ClsLawReadRstMsg(intP, strQuery)
           '    If intP = 1 Then
           '       If Val(rsAD1.Fields("cnt")) = 1 Then
           '          strSql = "Update ConsultRecordList Set CRL65=null Where CRL65=CRL01 AND CRL01='" & pGrpNo & "' "
           '          cnnConnection.Execute strSql
           '       End If
           '    End If
           'End If
           'end 2024/04/18
       End If
       'GRD1.TextMatrix(m_idX, colCRL55_1) = Trim(pNewData)
       'Mark by Lydia 2024/04/18 取消相同案號的設定>>移到外面
       '有一案兩請+相同案號的關聯，只要取消相同案號的設定
       'If pTName = "CRL55" And pNewData = "" And pCRL67 <> "" And Check1.Value = 1 Then
       '    strSql = "Update ConsultRecordList set CRL55=null, CRL56=null Where CRL01 IN (" & _
       '                 "SELECT CRL01 FROM ConsultRecordList,FLOW003 WHERE CRL67=" & CNULL(pCRL67) & _
       '                 " AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "')) "
       '    cnnConnection.Execute strSql
       'End If
       'end 2024/04/18
       If pOldData <> lblData(0).Caption Then
          pOldData = pNewData
       End If
    End If
    'Added by Lydia 2024/04/18 有一案兩請+相同案號的關聯，只要取消相同案號的設定
    If pEnabled = True And pTName = "CRL55" And pNewData = "" And pCRL67 <> "" And Check1.Value = 1 Then
       strSql = "Update ConsultRecordList set CRL55=null, CRL56=null Where CRL01 IN (" & _
                    "SELECT CRL01 FROM ConsultRecordList,FLOW003 WHERE CRL67=" & CNULL(pCRL67) & _
                    " AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "')) "
       cnnConnection.Execute strSql
       
       '檢查沒有子號，母號一併清除設定
       Call SaveGrpDataChk(pStatus, pEnabled, pCP140, pTName, "", strDataP, pDefCP14, pGrpNo, pCRL67)
    End If
    'end 2024/04/18
    
JumpToUpd:
    '最後一併更新承辦人
    If pDefCP14 <> m_CRC09 Or Trim(pNewData) <> "" Or (pOldData = lblData(0).Caption And pNewData = "") Then
       strBCase(0) = IIf(pNewData = "", pOldData, pNewData)
       If strBCase(0) <> "" Then
           tmpArrNew = Empty
           tmpArrNew = Split(strBCase(0), ",")
           For intG = 0 To UBound(tmpArrNew)
               Call ChgCaseNo(Replace(tmpArrNew(intG), "-", ""), strBCase)
               '同一關連CRL55,CRL67,CRL68+同一系統別
               If pDefCP14 = "" And m_CRC09 = pDefCP14 And strBCase(1) = m_CP01 And strBCase(2) <> "" Then '抓最新母案/子案的承辦人
                  'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉and nvl(cpm35,'0')<>'2'
                  strQuery = "select crc09 FROM ConsultRecCMP,ConsultRecordList,FLOW003,CasePropertyMap where crc01=crl01 and crl07='" & strBCase(1) & "' and crl08='" & strBCase(2) & "' and crl09='" & strBCase(3) & "' and crl10='" & strBCase(4) & "' " & _
                          "and crc09 is not null AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "') and crl07=cpm01(+) and crc03=cpm02(+) order by crc02 "
                  intP = 1
                  Set rsAD1 = ClsLawReadRstMsg(intP, strQuery)
                  If intP = 1 Then
                      pDefCP14 = "" & rsAD1.Fields("crc09")
                  End If
               End If
           Next intG
       End If
       '複數案號: 相同案號為先前收的案號 ; ex.CFP33679的相同案號為P-129313,P-130476
       If pDefCP14 = "" And InStr(pNewData, ",") > 0 Then
            strQuery = "select crc09 FROM ConsultRecCMP,ConsultRecordList,FLOW003,CasePropertyMap where crc01=crl01 and " & pTName & "=" & CNULL(pNewData) & _
                    "and crc09 is not null AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "') and crl07=cpm01(+) and crc03=cpm02(+) order by crc02 "
            intP = 1
            Set rsAD1 = ClsLawReadRstMsg(intP, strQuery)
            If intP = 1 Then
                pDefCP14 = "" & rsAD1.Fields("crc09")
            End If
       End If
       'Modified by Lydia 2024/04/23 排除現在為「大陸案、日本案」;因為大陸案有超過一半以上會改分給品薇,日本案很常會分給冠智F5717 (其實就是給外面的人翻)
       'If pDefCP14 <> "" And strBCase(0) <> "" Then
       If pDefCP14 <> "" And strBCase(0) <> "" And m_Na01 <> "020" And m_Na01 <> "011" Then
           '同一關連(CRL55,CRL67,CRL68)+同一系統別CRL07才一併更新
           'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉and nvl(cpm35,'0')<>'2'
           'Modified by Lydia 2024/04/23 排除「大陸案、日本案」
           'strSql = "Update ConsultRecCMP set CRC09='" & pDefCP14 & "' Where CRC09 IS NULL AND (CRC01,CRC03) IN (" & _
                       "SELECT CRL01,CRC03 FROM ConsultRecCMP,ConsultRecordList,FLOW003,CasePropertyMap WHERE crc01=crl01 and " & pTName & "=" & CNULL(strBCase(0)) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                       " AND CRL07='" & m_CP01 & "' and crl07=cpm01(+) and crc03=cpm02(+) AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "')) "
           'Modified by Lydia 2025/06/17 商申ST93=T31排除商爭案性質 IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "")
           strSql = "Update ConsultRecCMP set CRC09='" & pDefCP14 & "' Where CRC09 IS NULL " & IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',crc03)=0 ", "") & " AND (CRC01,CRC03) IN (" & _
                       "SELECT CRL01,CRC03 FROM ConsultRecCMP,ConsultRecordList,FLOW003,CasePropertyMap,Patent WHERE crc01=crl01 and " & pTName & "=" & CNULL(strBCase(0)) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                       " AND CRL07='" & m_CP01 & "' and crl07=cpm01(+) and crc03=cpm02(+) AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "') " & _
                       " AND CRL07=PA01(+) AND CRL08=PA02(+) AND CRL09=PA03(+) AND CRL10=PA04(+) AND PA09 NOT IN ('020','011') ) "
           cnnConnection.Execute strSql
           'Modified by Lydia 2023/01/06 專利案分所主管分案先不上進度檔的承辦人，要等到北所主管執行「確定分案」
           'If pStatus = "2" Then '確定分案
           'Modified by Lydia 2025/10/20 改成判斷權限
           'If pStatus = "2" And m_ST06 = "1" Then
           If pStatus = "2" And m_nowF0308 = "A6" Then
              'Modified by Lydia 2023/01/11 經過討論不限制案件性質,拿掉and nvl(cpm35,'0')<>'2'
              'Modified by Lydia 2023/05/19 一律以主管分案為準AND CP159=0 AND CP14 IS NULL => AND CP159=0;ex.P-131540(AB2018840)因為主管指示程序退回重新分案換工程師，但是沒有回寫新的承辦人，經過與Sindy討論決定都以主管為主
              'Modified by Lydia 2024/04/23 排除「大陸案、日本案」
              'strSql = "Update CaseProgress set CP14='" & pDefCP14 & "' Where CP158=0 AND CP159=0 AND (CP140,CP09) IN (" & _
                       "SELECT CRL01,CRC03 FROM ConsultRecCMP,ConsultRecordList,FLOW003,CasePropertyMap WHERE crc01=crl01 and " & pTName & "=" & CNULL(strBCase(0)) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                       " AND CRL07='" & m_CP01 & "' and crl07=cpm01(+) and crc03=cpm02(+) AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "')) "
              'Modified by Lydia 2025/06/17 商申ST93=T31排除商爭案性質 IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',cp10)=0 ", "")
              strSql = "Update CaseProgress set CP14='" & pDefCP14 & "' Where CP158=0 AND CP159=0 " & IIf(bolT31xT11 = True, " and instr('" & TMdebate & "',cp10)=0 ", "") & "AND (CP140,CP09) IN (" & _
                       "SELECT CRL01,CRC03 FROM ConsultRecCMP,ConsultRecordList,FLOW003,CasePropertyMap,Patent WHERE crc01=crl01 and " & pTName & "=" & CNULL(strBCase(0)) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                       " AND CRL07='" & m_CP01 & "' and crl07=cpm01(+) and crc03=cpm02(+) AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "') " & _
                       " AND CRL07=PA01(+) AND CRL08=PA02(+) AND CRL09=PA03(+) AND CRL10=PA04(+) AND PA09 NOT IN ('020','011') ) "
              cnnConnection.Execute strSql
           End If
       End If
       
    End If
    SaveGrpData = True
    Exit Function
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "更新" & pTName & "發生失敗"
    End If
End Function

'Added by Lydia 2023/02/18 檢查沒有子號，母號一併清除設定
Private Sub SaveGrpDataChk(ByVal pStatus As String, ByVal pEnabled As Boolean, ByVal pCP140 As String, ByVal pTName As String, ByRef pNewData As String, ByRef pOldData As String, ByVal pDefCP14 As String, ByVal pGrpNo As String, ByVal pCRL67 As String)
Dim tmpArrOld As Variant
Dim strOCase(0 To 4) As String
Dim intH As Integer, intCnt As Integer
       
   tmpArrOld = Empty
   tmpArrOld = Split(pOldData)

   intCnt = 0
   For intH = 0 To UBound(tmpArrOld)
      strOCase(0) = Trim(tmpArrOld(intH))
      Call ChgCaseNo(Replace(strOCase(0), "-", ""), strOCase)
      'Modified by Lydia 2024/04/18 改成子母號加起來
      'strQuery = "select crl01 from ConsultRecordList,flow003 where " & pTName & "=" & CNULL(pOldData) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                " and crl07||crl08||crl09||crl10<>" & CNULL(strOCase(1) & strOCase(2) & strOCase(3) & strOCase(4)) & _
                " and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) "
      strQuery = "select crl01 from ConsultRecordList,flow003 where " & pTName & "=" & CNULL(pOldData) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
                " and crl01=f0301 and f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "' ) " & _
                " group by crl01"
      intP = 1
      Set rsAD1 = ClsLawReadRstMsg(intP, strQuery)
      If intP = 1 Then
          intCnt = intCnt + rsAD1.RecordCount
      End If
   Next intH
   'Modified by Lydia 2024/04/18
   'If intCnt = 0 Then
   If intCnt < 2 Then
      strSql = "Update ConsultRecordList set " & pTName & "=null " & IIf(pTName = "CRL55", ",CRL56=NULL", "") & " Where CRL01 IN (" & _
               "SELECT CRL01 FROM ConsultRecordList,FLOW003 WHERE " & pTName & "=" & CNULL(pOldData) & IIf(pTName = "CRL55", " AND CRL56='Y' ", "") & _
               " AND CRL01=F0301 AND f0309 in ('" & Flow_待分案 & "','" & Flow_補件完成 & "','" & Flow_智權補件 & "','" & Flow_程序補件 & "')) "
      cnnConnection.Execute strSql
   End If

End Sub

'Added by Lydia 2025/07/24 增加對商申主管的權限判斷，若分案的性質保含「商爭性質」並且商爭主管當時請假，彈訊息選擇是否要一併為商爭分案。
Private Sub ChkT31xT11(ByVal sCP140 As String)
Dim intS As Integer, strS1 As String
Dim rsSD As New ADODB.Recordset

   If Pub_StrUserSt93 = "T31" Then
      bolT31xT11 = True
      If bolDutyT11 = True Then  '商爭主管休假
         strS1 = "select cp01||'-'||cp02||decode(cp03||cp04,000,null,'-'||cp03||'-'||cp04) as caseno, crc08 from consultreccmp,caseprogress where crc08=cp09(+) and crc01='" & sCP140 & "' and crc09 is null and instr('" & TMdebate & "',crc03)>0 "
         intS = 1
         Set rsSD = ClsLawReadRstMsg(intS, strS1)
         If intS = 1 Then
             If MsgBox(rsSD.Fields("caseno") & "是否要一併為商爭分案？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                bolT31xT11 = False
             End If
         End If
      End If
   Else
      bolT31xT11 = False
   End If
End Sub
