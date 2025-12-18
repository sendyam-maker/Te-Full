VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010613 
   BorderStyle     =   1  '單線固定
   Caption         =   "信件記錄查詢"
   ClientHeight    =   6870
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9080
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "V"
      Height          =   225
      Left            =   8730
      Style           =   1  '圖片外觀
      TabIndex        =   54
      Top             =   450
      Width           =   285
   End
   Begin VB.Frame Frame2 
      Caption         =   "電腦中心使用的"
      Height          =   2025
      Left            =   4560
      TabIndex        =   52
      Top             =   1290
      Width           =   4455
      Begin VB.OptionButton OptPi15 
         Caption         =   "包含"
         Height          =   225
         Index           =   1
         Left            =   3660
         TabIndex        =   27
         Top             =   870
         Width           =   735
      End
      Begin VB.OptionButton OptPi15 
         Caption         =   "排除"
         Height          =   225
         Index           =   0
         Left            =   3000
         TabIndex        =   26
         Top             =   870
         Width           =   735
      End
      Begin VB.TextBox TxtPi15 
         Height          =   285
         Left            =   1140
         TabIndex        =   25
         Top             =   810
         Width           =   1845
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Caption         =   "X"
         Height          =   225
         Left            =   3840
         Style           =   1  '圖片外觀
         TabIndex        =   53
         Top             =   0
         Width           =   285
      End
      Begin VB.TextBox txtTime_E 
         Height          =   285
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   24
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox txtTime_S 
         Height          =   285
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   23
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox txtII03 
         Height          =   285
         Left            =   1140
         MaxLength       =   5
         TabIndex        =   22
         Top             =   210
         Width           =   765
      End
      Begin VB.CommandButton cmdM51 
         BackColor       =   &H00C0C0FF&
         Caption         =   "沖銷信件,上確認註記"
         Height          =   360
         Left            =   2310
         Style           =   1  '圖片外觀
         TabIndex        =   29
         Top             =   1470
         Width           =   1905
      End
      Begin VB.CommandButton cmdM51c 
         BackColor       =   &H00C0FFC0&
         Caption         =   "恢復為無處理狀態"
         Height          =   360
         Left            =   270
         Style           =   1  '圖片外觀
         TabIndex        =   28
         Top             =   1470
         Width           =   1905
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "轉入編號："
         Height          =   180
         Left            =   210
         TabIndex        =   60
         Top             =   270
         Width           =   900
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4350
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "執行功能："
         Height          =   180
         Left            =   210
         TabIndex        =   58
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "系統記錄："
         Height          =   180
         Left            =   210
         TabIndex        =   57
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "轉入時間："
         Height          =   180
         Left            =   210
         TabIndex        =   55
         Top             =   570
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1890
         X2              =   2280
         Y1              =   660
         Y2              =   660
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   345
      Index           =   5
      Left            =   4280
      Style           =   1  '圖片外觀
      TabIndex        =   62
      Top             =   2340
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Height          =   345
      Index           =   4
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   61
      Top             =   2340
      Width           =   1080
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   0
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1260
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1260
      Width           =   855
   End
   Begin VB.TextBox txtTime_E2 
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtTime_S2 
      Height          =   285
      Left            =   1050
      MaxLength       =   6
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtPI20 
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   16
      Top             =   2130
      Width           =   255
   End
   Begin VB.TextBox txtPI18 
      Height          =   270
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   14
      Top             =   2130
      Width           =   495
   End
   Begin VB.TextBox txtPI19 
      Height          =   270
      Left            =   1545
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2130
      Width           =   855
   End
   Begin VB.TextBox txtPI21 
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   17
      Top             =   2130
      Width           =   375
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   260
      ItemData        =   "frm06010613.frx":0000
      Left            =   990
      List            =   "frm06010613.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   0
      Width           =   1710
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H000000C0&
      Height          =   300
      ItemData        =   "frm06010613.frx":0025
      Left            =   2790
      List            =   "frm06010613.frx":002F
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   300
      Width           =   600
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   285
      Left            =   120
      TabIndex        =   42
      Top             =   1860
      Width           =   2805
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   7
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   13
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Index           =   6
         Left            =   930
         MaxLength       =   7
         TabIndex        =   12
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "刪除日期："
         Height          =   180
         Left            =   0
         TabIndex        =   43
         Top             =   60
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1680
         X2              =   2070
         Y1              =   150
         Y2              =   150
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含刪除未轉寄資料"
      Height          =   240
      Left            =   6210
      TabIndex        =   21
      Top             =   930
      Width           =   2355
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "信件狀況"
      Height          =   360
      Left            =   7200
      TabIndex        =   31
      Top             =   60
      Width           =   915
   End
   Begin VB.ComboBox cboII05 
      Height          =   260
      ItemData        =   "frm06010613.frx":003B
      Left            =   3780
      List            =   "frm06010613.frx":0057
      Style           =   2  '單純下拉式
      TabIndex        =   18
      Top             =   900
      Width           =   1860
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   5
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   5
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   4
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   4
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   3
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   2
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   360
      Left            =   6360
      TabIndex        =   30
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8160
      TabIndex        =   32
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm06010613.frx":0095
      Height          =   3705
      Left            =   60
      TabIndex        =   44
      Top             =   2700
      Width           =   8955
      _ExtentX        =   15804
      _ExtentY        =   6526
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "轉入日期時間|主旨|分類|收受者|轉寄日期時間|刪除日期時間|收信日期時間"
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
      _Band(0).Cols   =   7
   End
   Begin VB.ComboBox Combo5 
      Height          =   300
      ItemData        =   "frm06010613.frx":00AA
      Left            =   7410
      List            =   "frm06010613.frx":00AC
      Style           =   2  '單純下拉式
      TabIndex        =   51
      Top             =   7080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSForms.ComboBox cboIR13 
      Height          =   285
      Left            =   4260
      TabIndex        =   3
      Top             =   300
      Width           =   1710
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3016;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboII06 
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   300
      Width           =   1710
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3016;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtII11 
      Height          =   300
      Left            =   3780
      TabIndex        =   19
      Top             =   1200
      Width           =   5240
      VariousPropertyBits=   679495707
      Size            =   "9234;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtII17 
      Height          =   810
      Left            =   3780
      TabIndex        =   20
      Top             =   1500
      Width           =   5240
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "9243;1429"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1800
      X2              =   2190
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "轉入日期："
      Height          =   180
      Left            =   120
      TabIndex        =   59
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "信件時間："
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   1020
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1800
      X2              =   2190
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label13 
      Caption         =   "本所案號："
      Height          =   225
      Left            =   120
      TabIndex        =   50
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label12 
      Caption         =   "共用信箱："
      Height          =   180
      Left            =   60
      TabIndex        =   49
      Top             =   60
      Width           =   945
   End
   Begin VB.Label Label11 
      Caption         =   "＊代表已處理完"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   150
      TabIndex        =   48
      Top             =   2430
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "　　105/11/17 IPDept信件停止在系統內轉寄,一律用Outlook轉寄給個人;只有專利處除外"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   47
      Top             =   7140
      Width           =   7785
   End
   Begin VB.Label Label1 
      Caption         =   "下列備註只有電腦中心人員才看的到："
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   46
      Top             =   6900
      Width           =   7785
   End
   Begin VB.Label Label1 
      Caption         =   "註：105/6/22及105/6/23上午(信件未入系統) 請至IPDept信箱查看"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   45
      Top             =   6660
      Width           =   7785
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   150
      X2              =   6040
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Label Label10 
      Caption         =   "轉寄者："
      Height          =   180
      Left            =   3480
      TabIndex        =   41
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "寄件者："
      Height          =   180
      Left            =   3000
      TabIndex        =   40
      Top             =   1200
      Width           =   740
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1770
      X2              =   2160
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label Label4 
      Caption         =   "主　旨："
      Height          =   180
      Left            =   3000
      TabIndex        =   39
      Top             =   1530
      Width           =   740
   End
   Begin VB.Label Label7 
      Caption         =   "收受者："
      Height          =   180
      Left            =   225
      TabIndex        =   38
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "信件日期："
      Height          =   180
      Left            =   120
      TabIndex        =   37
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "轉寄日期："
      Height          =   180
      Left            =   120
      TabIndex        =   36
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "分　類："
      Height          =   180
      Left            =   3000
      TabIndex        =   35
      Top             =   960
      Width           =   740
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1740
      X2              =   2160
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Label LblTotCnt 
      Caption         =   "總筆數:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7470
      TabIndex        =   34
      Top             =   6420
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "備註：1.雙擊”主旨”開啟信件"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   33
      Top             =   6420
      Width           =   2505
   End
End
Attribute VB_Name = "frm06010613"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/16 Form2.0已修改
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Dim m_AttachPath As String
Dim nCol As Long, nRow As Long
Public m_WorkType As String '0.信箱主檔 1.轉寄
Public m_MailUsernum As String
Dim m_PrevForm As Form '前一畫面
Dim m_MailListIndex As Integer '記錄人員原權限可以查看的信箱為何
Public cmdState As Integer '紀錄作用按鍵


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cboII06_Validate(Cancel As Boolean)
   If cboII06.Text <> "" Then
      Call cboII06_LostFocus
'      '檢查人員是否存在或離職
'      If ChkStaffST04(Left(cboII06, 5)) = True Then
'         cboII06.Text = ""
'         cboII06.SetFocus
'         Call cboII06_GotFocus
'         Exit Sub
'      End If
   End If
'   If Len(Trim(cboII06.Text)) = 5 Then
'      cboII06.Text = Left(cboII06.Text, 5) & " " & GetStaffName(Left(cboII06.Text, 5), True)
'   End If
   'Add By Sindy 2017/12/19
   'Modify By Sindy 2019/4/15
   'Modify By Sindy 2019/6/4 mark
'   If UCase(cboII06.Text) = UCase("patent") Or _
'      UCase(cboII06.Text) = UCase("tm") Then
'      Combo2.ListIndex = 0
'   Else
      If Combo2.Enabled = False Then
         Combo2.ListIndex = m_MailListIndex
      End If
'   End If
   '2017/12/19 END
End Sub

Private Sub SetCombo2Limit()
   '決定信箱名稱
   If Pub_StrUserSt03 = "M51" Then
      Combo2.Enabled = True
      m_MailListIndex = 1 '記錄人員原權限可以查看的信箱
   Else
      Me.Caption = "信件記錄查詢 - 信件收受者"
      Combo2.Enabled = True 'False Modify By Sindy 2022/5/23 經理說不用鎖
      'Modify By Sindy 2019/6/28 特殊人員
      If strUserNum = "67002" Or strUserNum = "98020" Then '商標處
         m_MailListIndex = 2 '記錄人員原權限可以查看的信箱
      '2019/6/28 END
      ElseIf Left(Pub_StrUserSt03, 1) = "F" Then '國外部
         m_MailListIndex = 0 '記錄人員原權限可以查看的信箱
         Me.Caption = "信件記錄查詢 - 信件收受者（此處只可查詢到分信至Patent的信件）"
      ElseIf Left(Pub_StrUserSt03, 2) = "P1" Then '專利處
         m_MailListIndex = 1 '記錄人員原權限可以查看的信箱
      ElseIf Left(Pub_StrUserSt03, 2) = "P2" Then  '商標處
         m_MailListIndex = 2 '記錄人員原權限可以查看的信箱
      End If
   End If
   Combo2.ListIndex = m_MailListIndex
   Me.Caption = "信件記錄查詢(主檔)"
   If m_WorkType = "1" Then '轉寄
      Check1.Visible = False
      Me.Caption = "信件記錄查詢(個人收件區)"
      cboII06.Text = IIf(Len(m_MailUsernum) <= 5, m_MailUsernum & " " & GetStaffName(m_MailUsernum, True), m_MailUsernum)
      cboII06.Tag = frm06010613.cboII06.Text
      cboIR13.Text = IIf(Len(m_MailUsernum) <= 5, m_MailUsernum & " " & GetStaffName(m_MailUsernum, True), m_MailUsernum)
      cboIR13.Tag = frm06010613.cboII06.Text
      txtDate(2) = CompWorkDay(2, strSrvDate(1), 1) - 19110000 'Modify By Sindy 2024/6/17
      txtDate(3) = strSrvDate(2)
      Call cboII06_LostFocus
   Else '信箱主檔
      Combo1.Enabled = False
      Me.Caption = "信件記錄查詢 - 信件主檔"
      cboIR13 = "" 'strUserNum & " " & GetPrjSalesNM(strUserNum)
      cboIR13.Tag = frm06010613.cboIR13.Text
      txtDate(0) = CompWorkDay(2, strSrvDate(1), 1) - 19110000 'Modify By Sindy 2024/6/17
      txtDate(1) = strSrvDate(2) 'Modify By Sindy 2022/8/11
      'txtDate(4) = strSrvDate(2) 'Add By Sindy 2020/8/25
      'txtDate(5) = strSrvDate(2) 'Add By Sindy 2020/8/25
   End If
'   If Pub_StrUserSt03 = "M51" Then
'      txtDate(2) = ""
'      txtDate(3) = ""
'   End If
End Sub

Private Sub SetFrame1()
   'Frame1.Visible = False
   txtDate(6).Enabled = False
   txtDate(7).Enabled = False
   If Check1.Visible = True Then
      'Frame1.Visible = True
      txtDate(6).Enabled = True
      txtDate(7).Enabled = True
   'ElseIf cboII06.Text <> "" And cboIR13.Text = "" Then
   ElseIf cboII06.Text <> "" Then
      'Frame1.Visible = True
      txtDate(6).Enabled = True
      txtDate(7).Enabled = True
   End If
End Sub

'Add By Sindy 2016/5/18
Private Sub cboII06_GotFocus()
   If cboII06.Text = "" Then
      cboII06.Text = cboII06.Tag
   End If
   Call SetFrame1
   cboII06.SelStart = 0
   cboII06.SelLength = Len(cboII06.Text)
End Sub
Private Sub cboII06_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cboII06_LostFocus()
Dim strText As String
   
   Call SetFrame1
   If cboII06.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboII06.Text)
      If strText <> "" Then
         cboII06.Text = strText & " " & cboII06.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboII06.Text, 5))
         If strText <> "" Then
            cboII06.Text = Left(cboII06.Text, 5) & " " & strText
         End If
      End If
   End If
End Sub
'2016/5/18 END

Private Sub cboIr13_Validate(Cancel As Boolean)
   If cboIR13.Text <> "" Then
      Call cboIr13_LostFocus
'      '檢查人員是否存在或離職
'      If ChkStaffST04(Left(cboIR13, 5)) = True Then
'         cboIR13.Text = ""
'         cboIR13.SetFocus
'         Call cboIr13_GotFocus
'         Exit Sub
'      End If
   End If
'   If Len(Trim(cboIr13.Text)) = 5 Then
'      cboIr13.Text = Left(cboIr13.Text, 5) & " " & GetStaffName(Left(cboIr13.Text, 5), True)
'   End If
End Sub
'Add By Sindy 2016/5/18
Private Sub cboIr13_GotFocus()
   If Me.Check1.Visible = True Then
      If UCase(cboIR13.List(0)) = UCase("ipdept") Then cboIR13.RemoveItem (0)
   End If
   If cboIR13.Text = "" Then
      cboIR13.Text = cboIR13.Tag
   End If
   Call SetFrame1
   cboIR13.SelStart = 0
   cboIR13.SelLength = Len(cboIR13.Text)
End Sub
Private Sub cboIr13_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cboIr13_LostFocus()
Dim strText As String
   
   Call SetFrame1
   'If cboIR13.Text <> "" Then
   If cboIR13.Text <> "" And Left(cboIR13.Text, 5) <> "QPGMR" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboIR13.Text)
      If strText <> "" Then
         cboIR13.Text = strText & " " & cboIR13.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboIR13.Text, 5))
         If strText <> "" Then
            cboIR13.Text = Left(cboIR13.Text, 5) & " " & strText
         End If
      End If
   End If
End Sub
'2016/5/18 END

Private Sub cmdDetail_Click()
   Call PubShowNextData_2
End Sub

'Add By Sindy 2016/5/19
Public Function PubShowNextData_2() As Boolean
   PubShowNextData_2 = False
   'If dblPrevRow > 0 Then
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" And GRD1.TextMatrix(i, 14) <> "" Then
         PubShowNextData_2 = True
         '明細資料
         frm06010613_1.m_II01 = GRD1.TextMatrix(i, 9)
         frm06010613_1.m_II02 = GRD1.TextMatrix(i, 10)
         frm06010613_1.m_II03 = GRD1.TextMatrix(i, 14)
         frm06010613_1.m_II19 = GRD1.TextMatrix(i, 11)
         Call CancelRowColor(i)
         frm06010613_1.CmdNext.Enabled = False
         For j = i To GRD1.Rows - 1
            If GRD1.TextMatrix(j, 0) = "V" And GRD1.TextMatrix(j, 14) <> "" Then
               frm06010613_1.CmdNext.Enabled = True
               Exit For
            End If
         Next j
         Call frm06010613_1.SetParent(Me)
         frm06010613_1.Show
         frm06010613_1.QueryData
         Me.Hide
         Exit Function
      End If
   Next i
   'End If
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

'國外部信箱
Private Function ReadIpDeptData(ByVal strUserId As String) As String
Dim strCon As String
   
   strCon = ""
   '轉入日期
   If txtDate(0) <> "" And txtDate(1) <> "" Then
      strCon = strCon & " and ii01>=" & DBDATE(txtDate(0)) & " and ii01<=" & DBDATE(txtDate(1))
      
      'Add By Sindy 2017/12/22
      If txtII03 <> "" And Frame2.Visible = True Then
         strCon = strCon & " and ii03='" & txtII03 & "'"
      End If
      '2017/12/22 END
   End If
   
   'Add By Sindy 2020/7/28
   '轉入時間
   If txtTime_S <> "" And txtTime_E <> "" And Frame2.Visible = True Then
      strCon = strCon & " and ii02>=" & txtTime_S & " and ii02<=" & txtTime_E
   End If
   '2020/7/28 END
   
   '信件日期
   If txtDate(4) <> "" And txtDate(5) <> "" Then
      strCon = strCon & " and ((ii12>=" & DBDATE(txtDate(4)) & " and ii12<=" & DBDATE(txtDate(5)) & ")"
      'Add by Sindy 2020/11/16 未傳遞信件:無收信日期
      strCon = strCon & " or (ii01>=" & DBDATE(txtDate(4)) & " and ii01<=" & DBDATE(txtDate(5)) & " and ii12=0)"
      '2020/11/16 END
      strCon = strCon & ")"
   End If
   
   'Add By Sindy 2020/7/28
   '信件時間
   If txtTime_S2 <> "" And txtTime_E2 <> "" Then
      strCon = strCon & " and ii13>=" & txtTime_S2 & " and ii13<=" & txtTime_E2
   End If
   '2020/7/28 END
   
'   'Add By Sindy 2018/6/28 本所案號
'   If txtPI18 <> "" And txtPI19 <> "" Then
'      strCon = strCon & " and cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "'"
'   End If
'   '2018/6/28 END
   'Modify By Sindy 2022/7/21 本所案號
   If txtPI18 <> "" And txtPI19 <> "" Then
      strCon = strCon & " and ii23='" & txtPI18 & "' and ii24='" & txtPI19 & "' and ii25='" & txtPI20 & "' and ii26='" & txtPI21 & "'"
   'Modify By Sindy 2023/5/11
   ElseIf txtPI18 <> "" And txtPI19 = "" Then
      strCon = strCon & " and ii23='" & txtPI18 & "'"
      '2023/5/11 END
   End If
   '2022/7/21 END
   '分類
   If cboII05 <> "" Then
      strCon = strCon & " and ii05='" & Trim(Left(cboII05, 2)) & "'"
   End If
   '主旨
   If Trim(txtII17) <> "" Then
      strCon = strCon & " and instr(upper(ii17),upper('" & ChgSQL(Trim(txtII17)) & "'))>0"
   End If
   '寄件者
   If txtII11 <> "" Then
      strCon = strCon & " and instr(upper(ii11),upper('" & ChgSQL(txtII11) & "'))>0"
   End If
   
   'Add By Sindy 2020/11/9
   '系統記錄
   If TxtPi15 <> "" And Frame2.Visible = True Then
      If OptPi15(0).Value = True Then '排除
         strCon = strCon & " and instr(upper(Ii18),upper('" & ChgSQL(TxtPi15) & "'))=0"
      Else '包含
         strCon = strCon & " and instr(upper(Ii18),upper('" & ChgSQL(TxtPi15) & "'))>0"
      End If
   End If
   '2020/11/9 END
   
   '個人收件區:刪除日期放個人的刪除日期
   '不會抓到轉寄新知的信件資料
   '有轉入日期(20160525)不等於轉寄日期(20160526)
   If Check1.Visible = False Or Trim(cboII06) <> "" Or Trim(cboIR13) <> "" Then
      '轉寄日期
      If txtDate(2) <> "" And txtDate(3) <> "" Then
         strCon = strCon & " and ir11>=" & DBDATE(txtDate(2)) & " and ir11<=" & DBDATE(txtDate(3))
      End If
      '刪除日期
      'If Frame1.Visible = True Then
      If cboII06 <> "" Then
         If txtDate(6) <> "" And txtDate(7) <> "" Then
            strCon = strCon & " and ir04='" & strUserId & "'"
            strCon = strCon & " and ir08>=" & DBDATE(txtDate(6)) & " and ir08<=" & DBDATE(txtDate(7))
         End If
      End If
      'End If
      '同時有輸入收受者和轉寄者時,才要考慮or或and問題
      If cboII06 <> "" And cboIR13 <> "" Then
         If Combo1.Enabled = True And Combo1.ListIndex = 0 Then '或
            If UCase(cboIR13) = UCase("ipdept") Then
               strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or ir15='Y')"
            Else
               If m_WorkType = "1" Then
                  strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null))"
               Else
                  strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')))"
               End If
            End If
         Else '及
            If UCase(cboIR13) = UCase("ipdept") Then
               strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and ir15='Y'"
            Else
               If m_WorkType = "1" Then
                  strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null)"
               Else
                  strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "'))"
               End If
            End If
         End If
      Else
         '收受者
         If cboII06 <> "" Then
            strCon = strCon & " and upper(ir04)=upper('" & strUserId & "')"
         End If
         '轉寄者
         If UCase(cboIR13) = UCase("ipdept") Then
            strCon = strCon & " and ir15='Y'"
         ElseIf cboIR13 <> "" Then
            If m_WorkType = "1" Then
               strCon = strCon & " and ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null)"
            Else
               strCon = strCon & " and (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "'))"
            End If
         End If
      End If
      '有輸入收受者並且無轉寄者時,才可以顯示本人刪除日期 --- Modify By Sindy 105/6/26 取消
'      strSql = "select distinct '' V,sqldatet(Ir01)||' '||sqltime6(Ir02) 轉入日期時間,II17 主旨" & _
'               ",decode(ii05," & Show國外部信件分類 & ",ii05) 分類,'' 收受者,decode(ir15,'Y','IPDept',st02||decode(ir14,null,'','代')) 轉寄者" & _
'               ",sqldatet(ir11)||' '||sqltime6(ir12) 轉寄日期時間" & _
'               "," & IIf(cboII06.Text <> "" And cboIR13.Text = "", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期" & _
'               ",sqldatet(ii12)||' '||sqltime6(ii13) 信件日期時間,II01,II02,II19,II06,II14,II03 檔名,ii18||decode(ii18," & Show國外部信件分類 & ",ii18) 系統記錄" & _
'               ",'' 本所案號,ii08,ii09,ir11,ir12" & _
'               " From ipdeptinput,inputrecord,staff" & _
'               " where ii01(+)=ir01 and ii02(+)=ir02 and ii03(+)=ir03 and ir13=st01(+)" & strCon & _
'               " order by ir11 desc,ir12 desc"
      '有輸入收受者,才可以顯示本人刪除日期 --- Modify By Sindy 105/6/26
      'Modify By Sindy 2016/11/16 IIf(cboII06.Text <> "", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期
      '==> decode(sign(instr(ii06,';')),0,sqldatet(ir08)||' '||sqltime6(ir09),'') 本人刪除日期
      '==> IIf(cboII06.Text <> "" and txtDate(6)<>"", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期
      'ii18||decode(ii18," & Show國外部信件分類 & ",ii18) 系統記錄
      Label11.Caption = "代表：＊已處理完"
      strSql = "select distinct '' V,decode(ii16,0,'    ','＊')||sqldatet(Ir01)||' '||sqltime6(Ir02) 轉入日期時間,II17 主旨" & _
               ",decode(ii05," & Show國外部信件分類 & ",ii05)||GetMailBox(ii01,ii03) 分類,'' 收受者,decode(ir15,'Y','IPDept',st02||decode(ir14,null,'','代')) 轉寄者" & _
               ",sqldatet(ir11)||' '||sqltime6(ir12) 轉寄日期時間" & _
               "," & IIf(cboII06.Text <> "" And txtDate(6) <> "", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期" & _
               ",sqldatet(ii12)||' '||sqltime6(ii13) 信件日期時間,II01,II02,IR21 總收文號,II06,II14,II03 檔名,ii18 轉信條件和系統記錄" & _
               ",decode(ii23,null,'',ii23||'-'||ii24||'-'||ii25||'-'||ii26) 本所案號,ii08,ii09,ir11,ir12,sqldatet(Ir01)||' '||sqltime6(Ir02) as Ir01Sort" & _
               " From ipdeptinput,inputrecord,staff" & _
               " where ii01=ir01 and ii02=ir02 and ii03=ir03 and ir13=st01(+)" & strCon & _
               " order by ir11 desc,ir12 desc" '" and ii19=cp09(+)"
   '信件匯入區:
   '發現有轉入日期(20160525)不等於轉寄日期(20160526)
   Else
      '轉寄日期
      If txtDate(2) <> "" And txtDate(3) <> "" Then
         'strCon = strCon & " and ii07 is null and ii08>=" & DBDATE(txtDate(2)) & " and ii08<=" & DBDATE(txtDate(3))
         strCon = strCon & " and ii08>=" & DBDATE(txtDate(2)) & " and ii08<=" & DBDATE(txtDate(3))
      End If
      '刪除日期
      If txtDate(6) <> "" And txtDate(7) <> "" Then
         Check1.Value = 1 'Add By Sindy 2016/9/26
         strCon = strCon & " and ii07='Y' and ii08>=" & DBDATE(txtDate(6)) & " and ii08<=" & DBDATE(txtDate(7))
      End If
      '含刪除未轉寄資料
      If Check1.Value = 1 Then '含
      Else '不含
         strCon = strCon & " and ii07 is null"
      End If
      '收受者
      If cboII06 <> "" Then
         strCon = strCon & " and instr(upper(ii06),upper('" & strUserId & "'))>0"
      End If
      '轉寄者
      If cboIR13 <> "" Then
         strCon = strCon & " and upper(ii10)=upper('" & Trim(Left(cboIR13, 5)) & "')"
      End If
      'ii18||decode(ii18," & Show國外部信件分類 & ",ii18) 系統記錄
      Label11.Caption = "代表：＊已處理完"
      strSql = "select '' V,decode(ii16,0,'    ','＊')||sqldatet(II01)||' '||sqltime6(II02) 轉入日期時間,II17 主旨" & _
               ",decode(ii05," & Show國外部信件分類 & ",ii05)||GetMailBox(ii01,ii03) 分類,'' 收受者,st02 轉寄者" & _
               ",decode(ii07,null,sqldatet(ii08)||' '||sqltime6(ii09),'') 轉寄日期時間" & _
               ",decode(ii07,null,'',sqldatet(ii08)||' '||sqltime6(ii09)) 刪除日期時間" & _
               ",sqldatet(ii12)||' '||sqltime6(ii13) 信件日期時間,II01,II02,II19 總收文號,II06,II14,II03 檔名,ii18 轉信條件和系統記錄" & _
               ",decode(ii23,null,'',ii23||'-'||ii24||'-'||ii25||'-'||ii26) 本所案號,ii08,ii09,'','',sqldatet(ii01)||' '||sqltime6(ii02) as Ir01Sort" & _
               " From ipdeptinput,staff" & _
               " where ii10=st01(+)" & strCon & _
               " order by ii05 asc,ii08 desc,ii09 desc" '" and ii19=cp09(+)"
   End If
   ReadIpDeptData = strSql
End Function

'專利處信箱
Private Function ReadPatentData(ByVal strUserId As String) As String
Dim strCon As String
   
   strCon = ""
   '轉入日期
   If txtDate(0) <> "" And txtDate(1) <> "" Then
      strCon = strCon & " and pi01>=" & DBDATE(txtDate(0)) & " and pi01<=" & DBDATE(txtDate(1))
      
      'Add By Sindy 2017/12/22
      If txtII03 <> "" And Frame2.Visible = True Then
         strCon = strCon & " and pi03='" & txtII03 & "'"
      End If
      '2017/12/22 END
   End If
   
   'Add By Sindy 2020/7/28
   '轉入時間
   If txtTime_S <> "" And txtTime_E <> "" And Frame2.Visible = True Then
      strCon = strCon & " and pi02>=" & txtTime_S & " and pi02<=" & txtTime_E
   End If
   '2020/7/28 END
   
   '信件日期
   If txtDate(4) <> "" And txtDate(5) <> "" Then
      strCon = strCon & " and ((pi12>=" & DBDATE(txtDate(4)) & " and pi12<=" & DBDATE(txtDate(5)) & ")"
      'Add by Sindy 2020/11/16 未傳遞信件:無收信日期
      strCon = strCon & " or (pi01>=" & DBDATE(txtDate(4)) & " and pi01<=" & DBDATE(txtDate(5)) & " and pi12=0)"
      '2020/11/16 END
      strCon = strCon & ")"
   End If
   
   'Add By Sindy 2020/7/28
   '信件時間
   If txtTime_S2 <> "" And txtTime_E2 <> "" Then
      strCon = strCon & " and pi13>=" & txtTime_S2 & " and pi13<=" & txtTime_E2
   End If
   '2020/7/28 END
   
   'Add By Sindy 2018/6/28 本所案號
   If txtPI18 <> "" And txtPI19 <> "" Then
      strCon = strCon & " and pi18='" & txtPI18 & "' and pi19='" & txtPI19 & "' and pi20='" & txtPI20 & "' and pi21='" & txtPI21 & "'"
   'Modify By Sindy 2023/5/11
   ElseIf txtPI18 <> "" And txtPI19 = "" Then
      strCon = strCon & " and pi18='" & txtPI18 & "'"
      '2023/5/11 END
   End If
   '2018/6/28 END
   '分類
   If cboII05 <> "" Then
      strCon = strCon & " and pi05='" & Trim(Left(cboII05, 2)) & "'"
   End If
   '主旨
   If Trim(txtII17) <> "" Then
      strCon = strCon & " and instr(upper(pi17),upper('" & ChgSQL(Trim(txtII17)) & "'))>0"
   End If
   '寄件者
   If txtII11 <> "" Then
      strCon = strCon & " and instr(upper(pi11),upper('" & ChgSQL(txtII11) & "'))>0"
   End If
   
   'Add By Sindy 2020/11/9
   '系統記錄
   If TxtPi15 <> "" And Frame2.Visible = True Then
      If OptPi15(0).Value = True Then '排除
         strCon = strCon & " and instr(upper(pi15),upper('" & ChgSQL(TxtPi15) & "'))=0"
      Else '包含
         strCon = strCon & " and instr(upper(pi15),upper('" & ChgSQL(TxtPi15) & "'))>0"
      End If
   End If
   '2020/11/9 END
   
   '個人收件區:刪除日期放個人的刪除日期
   If Check1.Visible = False Or Trim(cboII06) <> "" Or Trim(cboIR13) <> "" Then
      '轉寄日期
      If txtDate(2) <> "" And txtDate(3) <> "" Then
         strCon = strCon & " and ir11>=" & DBDATE(txtDate(2)) & " and ir11<=" & DBDATE(txtDate(3))
      End If
      '刪除日期
      If cboII06 <> "" Then
         If txtDate(6) <> "" And txtDate(7) <> "" Then
            strCon = strCon & " and upper(ir04)=upper('" & strUserId & "')"
            strCon = strCon & " and ir08>=" & DBDATE(txtDate(6)) & " and ir08<=" & DBDATE(txtDate(7))
         End If
      End If
      '同時有輸入收受者和轉寄者時,才要考慮or或and問題
      If cboII06 <> "" And cboIR13 <> "" Then
         If Combo1.Enabled = True And Combo1.ListIndex = 0 Then '或
            If UCase(cboIR13) = UCase("patent") Then
               strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or ir15='Y')"
            Else
               If m_WorkType = "1" Then
                  strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null))"
               Else
                  strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')))"
               End If
            End If
         Else '及
            If UCase(cboIR13) = UCase("patent") Then
               strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and ir15='Y'"
            Else
               If m_WorkType = "1" Then
                  strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null)"
               Else
                  strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "'))"
               End If
            End If
         End If
      Else
         '收受者
         If cboII06 <> "" Then
            strCon = strCon & " and ir04='" & strUserId & "'"
         End If
         '轉寄者
         If UCase(cboIR13) = UCase("patent") Then
            strCon = strCon & " and ir15='Y'"
         ElseIf cboIR13 <> "" Then
            If m_WorkType = "1" Then
               strCon = strCon & " and ((ir13='" & Trim(Left(cboIR13, 5)) & "' or ir14='" & Trim(Left(cboIR13, 5)) & "') and ir15 is null)"
            Else
               strCon = strCon & " and (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "'))"
            End If
         End If
      End If
      '有輸入收受者,才可以顯示本人刪除日期
      'Modify By Sindy 2016/11/16 IIf(cboII06.Text <> "", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期
      '==> decode(sign(instr(PI06,';')),0,sqldatet(ir08)||' '||sqltime6(ir09),'') 本人刪除日期
      '==> IIf(cboII06.Text <> "" and txtDate(6)<>"", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期
      Label11.Caption = "代表：＊已處理●不處理△歸卷"
      strSql = "select distinct '' V,decode(ir16,'1','＊','5','＊','2','●','4','△','    ')||sqldatet(Ir01)||' '||sqltime6(Ir02) 轉入日期時間,PI17 主旨" & _
               ",decode(Pi05," & Show專利處信件分類 & ",PI05)||GetMailBox(pi01,pi03) 分類,'' 收受者,decode(ir15,'Y','Patent',st02||decode(ir14,null,'','代')) 轉寄者" & _
               ",sqldatet(ir11)||' '||sqltime6(ir12) 轉寄日期時間" & _
               "," & IIf(cboII06.Text <> "" And txtDate(6) <> "", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期" & _
               ",sqldatet(pi12)||' '||sqltime6(pi13) 信件日期時間,PI01,PI02,IR21 總收文號,PI06,PI14,PI03 檔名,Pi15 轉信條件和系統記錄" & _
               ",PI18||'-'||PI19||'-'||PI20||'-'||PI21 本所案號,Pi08,Pi09,ir11,ir12,sqldatet(Ir01)||' '||sqltime6(Ir02) as Ir01Sort" & _
               " From patentinput,inputrecord,staff" & _
               " where Pi01=ir01 and Pi02=ir02 and Pi03=ir03 and ir13=st01(+)" & strCon & _
               " order by ir11 desc,ir12 desc"
   '信件匯入區:
   '轉入日期不等於轉寄日期
   Else
      '轉寄日期
      If txtDate(2) <> "" And txtDate(3) <> "" Then
         strCon = strCon & " and Pi08>=" & DBDATE(txtDate(2)) & " and Pi08<=" & DBDATE(txtDate(3))
      End If
      '刪除日期
      If txtDate(6) <> "" And txtDate(7) <> "" Then
         Check1.Value = 1
         strCon = strCon & " and Pi07='Y' and Pi08>=" & DBDATE(txtDate(6)) & " and Pi08<=" & DBDATE(txtDate(7))
      End If
      '含刪除未轉寄資料
      If Check1.Value = 1 Then '含
      Else '不含
         strCon = strCon & " and Pi07 is null"
      End If
      '收受者
      If cboII06 <> "" Then
         strCon = strCon & " and instr(upper(Pi06),upper('" & strUserId & "'))>0"
      End If
      '轉寄者
      If cboIR13 <> "" Then
         strCon = strCon & " and upper(Pi10)=upper('" & Trim(Left(cboIR13, 5)) & "')"
      End If
      Label11.Caption = "代表：＊已處理完"
      strSql = "select '' V,decode(pi16,0,'    ','＊')||sqldatet(PI01)||' '||sqltime6(PI02) 轉入日期時間,PI17 主旨" & _
               ",decode(Pi05," & Show專利處信件分類 & ",PI05)||GetMailBox(pi01,pi03) 分類,'' 收受者,st02 轉寄者" & _
               ",decode(Pi07,null,sqldatet(Pi08)||' '||sqltime6(Pi09),'') 轉寄日期時間" & _
               ",decode(Pi07,null,'',sqldatet(Pi08)||' '||sqltime6(Pi09)) 刪除日期時間" & _
               ",sqldatet(Pi12)||' '||sqltime6(Pi13) 信件日期時間,PI01,PI02,'' 總收文號,PI06,PI14,PI03 檔名,Pi15 轉信條件和系統記錄" & _
               ",PI18||'-'||PI19||'-'||PI20||'-'||PI21 本所案號,Pi08,Pi09,'','',sqldatet(pi01)||' '||sqltime6(pi02) as Ir01Sort" & _
               " From patentinput,staff" & _
               " where Pi10=st01(+)" & strCon & _
               " order by Pi05 asc,Pi08 desc,Pi09 desc"
   End If
   ReadPatentData = strSql
End Function

'Add By Sindy 2019/4/15
'商標處信箱
Private Function ReadTMData(ByVal strUserId As String) As String
Dim strCon As String
   
   strCon = ""
   '轉入日期
   If txtDate(0) <> "" And txtDate(1) <> "" Then
      strCon = strCon & " and Ti01>=" & DBDATE(txtDate(0)) & " and Ti01<=" & DBDATE(txtDate(1))
      
      If txtII03 <> "" And Frame2.Visible = True Then
         strCon = strCon & " and Ti03='" & txtII03 & "'"
      End If
   End If
   
   'Add By Sindy 2020/7/28
   '轉入時間
   If txtTime_S <> "" And txtTime_E <> "" And Frame2.Visible = True Then
      strCon = strCon & " and Ti02>=" & txtTime_S & " and Ti02<=" & txtTime_E
   End If
   '2020/7/28 END
   
   '信件日期
   If txtDate(4) <> "" And txtDate(5) <> "" Then
      strCon = strCon & " and ((Ti12>=" & DBDATE(txtDate(4)) & " and Ti12<=" & DBDATE(txtDate(5)) & ")"
      'Add by Sindy 2020/11/16 未傳遞信件:無收信日期
      strCon = strCon & " or (Ti01>=" & DBDATE(txtDate(4)) & " and Ti01<=" & DBDATE(txtDate(5)) & " and Ti12=0)"
      '2020/11/16 END
      strCon = strCon & ")"
   End If
   
   'Add By Sindy 2020/7/28
   '信件時間
   If txtTime_S2 <> "" And txtTime_E2 <> "" Then
      strCon = strCon & " and Ti13>=" & txtTime_S2 & " and Ti13<=" & txtTime_E2
   End If
   '2020/7/28 END
   
   '本所案號
   If txtPI18 <> "" And txtPI19 <> "" Then
      strCon = strCon & " and Ti18='" & txtPI18 & "' and Ti19='" & txtPI19 & "' and Ti20='" & txtPI20 & "' and Ti21='" & txtPI21 & "'"
   'Modify By Sindy 2023/5/11
   ElseIf txtPI18 <> "" And txtPI19 = "" Then
      strCon = strCon & " and Ti18='" & txtPI18 & "'"
      '2023/5/11 END
   End If
   '分類
   If cboII05 <> "" Then
      strCon = strCon & " and Ti05='" & Trim(Left(cboII05, 2)) & "'"
   End If
   '主旨
   If Trim(txtII17) <> "" Then
      strCon = strCon & " and instr(upper(Ti17),upper('" & ChgSQL(Trim(txtII17)) & "'))>0"
   End If
   '寄件者
   If txtII11 <> "" Then
      strCon = strCon & " and instr(upper(Ti11),upper('" & ChgSQL(txtII11) & "'))>0"
   End If
   
   'Add By Sindy 2020/11/9
   '系統記錄
   If TxtPi15 <> "" And Frame2.Visible = True Then
      If OptPi15(0).Value = True Then '排除
         strCon = strCon & " and instr(upper(Ti15),upper('" & ChgSQL(TxtPi15) & "'))=0"
      Else '包含
         strCon = strCon & " and instr(upper(Ti15),upper('" & ChgSQL(TxtPi15) & "'))>0"
      End If
   End If
   '2020/11/9 END
   
   '個人收件區:刪除日期放個人的刪除日期
   If Check1.Visible = False Or Trim(cboII06) <> "" Or Trim(cboIR13) <> "" Then
      '轉寄日期
      If txtDate(2) <> "" And txtDate(3) <> "" Then
         strCon = strCon & " and ir11>=" & DBDATE(txtDate(2)) & " and ir11<=" & DBDATE(txtDate(3))
      End If
      '刪除日期
      If cboII06 <> "" Then
         If txtDate(6) <> "" And txtDate(7) <> "" Then
            strCon = strCon & " and ir04='" & strUserId & "'"
            strCon = strCon & " and ir08>=" & DBDATE(txtDate(6)) & " and ir08<=" & DBDATE(txtDate(7))
         End If
      End If
      '同時有輸入收受者和轉寄者時,才要考慮or或and問題
      If cboII06 <> "" And cboIR13 <> "" Then
         If Combo1.Enabled = True And Combo1.ListIndex = 0 Then '或
            If UCase(cboIR13) = UCase("tm") Then
               strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or ir15='Y')"
            Else
               If m_WorkType = "1" Then
                  strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null))"
               Else
                  strCon = strCon & " and (upper(ir04)=upper('" & strUserId & "') or (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')))"
               End If
            End If
         Else '及
            If UCase(cboIR13) = UCase("tm") Then
               strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and ir15='Y'"
            Else
               If m_WorkType = "1" Then
                  strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null)"
               Else
                  strCon = strCon & " and upper(ir04)=upper('" & strUserId & "') and (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "'))"
               End If
            End If
         End If
      Else
         '收受者
         If cboII06 <> "" Then
            strCon = strCon & " and upper(ir04)=upper('" & strUserId & "')"
         End If
         '轉寄者
         If UCase(cboIR13) = UCase("tm") Then
            strCon = strCon & " and ir15='Y'"
         ElseIf cboIR13 <> "" Then
            If m_WorkType = "1" Then
               strCon = strCon & " and ((upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "')) and ir15 is null)"
            Else
               strCon = strCon & " and (upper(ir13)=upper('" & Trim(Left(cboIR13, 5)) & "') or upper(ir14)=upper('" & Trim(Left(cboIR13, 5)) & "'))"
            End If
         End If
      End If
      '有輸入收受者,才可以顯示本人刪除日期
      'Modify By Sindy 2016/11/16 IIf(cboII06.Text <> "", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期
      '==> decode(sign(instr(PI06,';')),0,sqldatet(ir08)||' '||sqltime6(ir09),'') 本人刪除日期
      '==> IIf(cboII06.Text <> "" and txtDate(6)<>"", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期
      Label11.Caption = "代表：＊已處理●不處理△歸卷"
      strSql = "select distinct '' V,decode(ir16,'1','＊','5','＊','2','●','4','△','    ')||sqldatet(Ir01)||' '||sqltime6(Ir02) 轉入日期時間,Ti17 主旨" & _
               ",decode(Ti05," & Show商標處信件分類 & ",Ti05)||GetMailBox(ti01,ti03) 分類,'' 收受者,decode(ir15,'Y','TM',st02||decode(ir14,null,'','代')) 轉寄者" & _
               ",sqldatet(ir11)||' '||sqltime6(ir12) 轉寄日期時間" & _
               "," & IIf(cboII06.Text <> "" And txtDate(6) <> "", "sqldatet(ir08)||' '||sqltime6(ir09)", "''") & " 本人刪除日期" & _
               ",sqldatet(Ti12)||' '||sqltime6(Ti13) 信件日期時間,Ti01,Ti02,IR21 總收文號,Ti06,Ti14,Ti03 檔名,Ti15 轉信條件和系統記錄" & _
               ",Ti18||'-'||Ti19||'-'||Ti20||'-'||Ti21 本所案號,Ti08,Ti09,ir11,ir12,sqldatet(Ir01)||' '||sqltime6(Ir02) as Ir01Sort" & _
               " From TMinput,inputrecord,staff" & _
               " where Ti01=ir01 and Ti02=ir02 and Ti03=ir03 and ir13=st01(+)" & strCon & _
               " order by ir11 desc,ir12 desc"
   '信件匯入區:
   '轉入日期不等於轉寄日期
   Else
      '轉寄日期
      If txtDate(2) <> "" And txtDate(3) <> "" Then
         strCon = strCon & " and Ti08>=" & DBDATE(txtDate(2)) & " and Ti08<=" & DBDATE(txtDate(3))
      End If
      '刪除日期
      If txtDate(6) <> "" And txtDate(7) <> "" Then
         Check1.Value = 1
         strCon = strCon & " and Ti07='Y' and Ti08>=" & DBDATE(txtDate(6)) & " and Ti08<=" & DBDATE(txtDate(7))
      End If
      '含刪除未轉寄資料
      If Check1.Value = 1 Then '含
      Else '不含
         strCon = strCon & " and Ti07 is null"
      End If
      '收受者
      If cboII06 <> "" Then
         strCon = strCon & " and instr(upper(Ti06),upper('" & strUserId & "'))>0"
      End If
      '轉寄者
      If cboIR13 <> "" Then
         strCon = strCon & " and upper(Ti10)=upper('" & Trim(Left(cboIR13, 5)) & "')"
      End If
      Label11.Caption = "代表：＊已處理完"
      strSql = "select '' V,decode(Ti16,0,'    ','＊')||sqldatet(Ti01)||' '||sqltime6(Ti02) 轉入日期時間,Ti17 主旨" & _
               ",decode(Ti05," & Show商標處信件分類 & ",Ti05)||GetMailBox(ti01,ti03) 分類,'' 收受者,st02 轉寄者" & _
               ",decode(Ti07,null,sqldatet(Ti08)||' '||sqltime6(Ti09),'') 轉寄日期時間" & _
               ",decode(Ti07,null,'',sqldatet(Ti08)||' '||sqltime6(Ti09)) 刪除日期時間" & _
               ",sqldatet(Ti12)||' '||sqltime6(Ti13) 信件日期時間,Ti01,Ti02,'' 總收文號,Ti06,Ti14,Ti03 檔名,Ti15 轉信條件和系統記錄" & _
               ",Ti18||'-'||Ti19||'-'||Ti20||'-'||Ti21 本所案號,Ti08,Ti09,'','',sqldatet(ti01)||' '||sqltime6(ti02) as Ir01Sort" & _
               " From TMinput,staff" & _
               " where Ti10=st01(+)" & strCon & _
               " order by Ti05 asc,Ti08 desc,Ti09 desc"
   End If
   ReadTMData = strSql
End Function

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strUser As String
Dim strUserId As String
   
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass
   
   GRD1.Clear
   Call SetGrd
   '檢查收受者是否為員工編號
   If GetPrjSalesNM(Trim(Left(cboII06, 5))) <> "" Then '依員工編號抓取員工姓名
      strUserId = Trim(Left(cboII06, 5))
   Else
      strUserId = LCase(cboII06)
   End If
   
   Call txtPI20_LostFocus 'Add By Sindy 2018/6/28
   Call txtPI21_LostFocus 'Add By Sindy 2018/6/28
   
   '各信箱查詢SQL:
   'Modify By Sindy 2016/9/26
'   If UCase(m_PrevForm.Name) = UCase("frm04010518") Or _
'      UCase(m_PrevForm.Name) = UCase("frm04010519") Then
   If UCase(Combo2.Text) = UCase("patent") Then
      Label11.Visible = True
      strSql = ReadPatentData(strUserId)
   'Add By Sindy 2019/4/15
   ElseIf UCase(Combo2.Text) = UCase("tm") Then
      Label11.Visible = True
      strSql = ReadTMData(strUserId)
   '2019/4/15 END
   Else
      Label11.Visible = True
      strSql = ReadIpDeptData(strUserId)
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   LblTotCnt.Caption = "總筆數: "
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      LblTotCnt.Caption = "總筆數: " & rsTmp.RecordCount
      If Check1.Visible = False Then
         '解析收受者
         For i = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(i, 17) = GRD1.TextMatrix(i, 19) And GRD1.TextMatrix(i, 18) = GRD1.TextMatrix(i, 20) Then
               GRD1.TextMatrix(i, 4) = PUB_ReadUserData(GRD1.TextMatrix(i, 12))
            Else
               strSql = "SELECT ir04 FROM inputrecord" & _
                        " WHERE ir01=" & GRD1.TextMatrix(i, 9) & _
                          " and ir02=" & GRD1.TextMatrix(i, 10) & _
                          " and ir03='" & GRD1.TextMatrix(i, 14) & "'" & _
                          " and ir11=" & GRD1.TextMatrix(i, 19) & _
                          " and ir12=" & GRD1.TextMatrix(i, 20)
               intI = 1
               strUser = ""
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  With RsTemp
                     RsTemp.MoveFirst
                     Do While RsTemp.EOF = False
                        strUser = strUser & ";" & RsTemp.Fields("ir04")
                        RsTemp.MoveNext
                     Loop
                  End With
               End If
               GRD1.TextMatrix(i, 4) = PUB_ReadUserData(Mid(strUser, 2))
            End If
'            'Add By Sindy 2016/9/26 多個收受者時,本人刪除日期不顯示
'            If InStr(GRD1.TextMatrix(i, 4), ",") > 0 Then
'               GRD1.TextMatrix(i, 7) = ""
'            End If
'            '2016/9/26 END
         Next i
      Else
         '解析收受者
         For i = 1 To GRD1.Rows - 1
            GRD1.TextMatrix(i, 4) = PUB_ReadUserData(GRD1.TextMatrix(i, 12))
         Next i
      End If
      SetColor
   Else
      ShowNoData
   End If
   rsTmp.Close
   
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   GRD1.Visible = True
   dblPrevRow = 0
   
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub cmdok_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   Call PubShowNextData
End Sub

Public Function PubShowNextData() As Boolean
   Select Case cmdState
      Case 4 '基本資料
         Dim Str01 As String
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               GRD1.Text = ""
               For j = 0 To GRD1.Cols - 1
                  GRD1.col = j
                  GRD1.CellBackColor = QBColor(15)
               Next j
               Str01 = SystemNumber(Trim(GRD1.TextMatrix(i, 16)), 1)
               If Mid(UCase(Str01), 1, 1) = "N" Then
                   Str01 = Mid(Str01, 2, 3)
               End If
               GRD1.col = 16
               If GRD1.Text <> "" Then
                  fnCloseAllFrm100
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Function
                  End If
                  Select Case Pub_RplStr(Str01)
                      Case "CFP", "FCP", "P"   '專利
                            Screen.MousePointer = vbHourglass
                            frm100101_3.Show
                            frm100101_3.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                            frm100101_3.StrMenu
                            Screen.MousePointer = vbDefault
                      Case "CFT", "FCT", "T", "TF"   '商標
                            Screen.MousePointer = vbHourglass
                            frm100101_4.Show
                            frm100101_4.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                            frm100101_4.StrMenu
                            Screen.MousePointer = vbDefault
                      'Modify By Sindy 2009/07/24 增加LIN系統類別
                      'modify by sonia 2019/7/29 +ACS系統類別
                      Case "CFL", "FCL", "L", "LIN", "ACS"   '法務
                            Screen.MousePointer = vbHourglass
                            frm100101_5.Show
                            frm100101_5.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                            frm100101_5.StrMenu
                            Screen.MousePointer = vbDefault
                      Case "LA"            '顧問
                            Screen.MousePointer = vbHourglass
                            frm100101_6.Show
                            frm100101_6.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                            frm100101_6.StrMenu
                            Screen.MousePointer = vbDefault
                      Case Else                  '服務
                           Select Case Pub_RplStr(Trim(GRD1.TextMatrix(i, 6)))
                               Case "TB"    '條碼
                                  Screen.MousePointer = vbHourglass
                                  frm100101_7.Show
                                  frm100101_7.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                                  frm100101_7.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TM"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_8.Show
                                  frm100101_8.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                                  frm100101_8.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TD"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_9.Show
                                  frm100101_9.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                                  frm100101_9.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TC", "CFC"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_A.Show
                                  frm100101_A.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                                  frm100101_A.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case Else
                                  Screen.MousePointer = vbHourglass
                                  frm100101_B.Show
                                  frm100101_B.Tag = Pub_RplStr(Trim(GRD1.TextMatrix(i, 16)))
                                  frm100101_B.StrMenu
                                  Screen.MousePointer = vbDefault
                            End Select
                  End Select
               Else
                  MsgBox "無本所案號！", vbInformation
               End If
               Me.Enabled = True
               Exit Function
            End If
         Next i
         Me.Enabled = True
         
      Case 5 '進度
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               GRD1.Text = ""
               For j = 0 To GRD1.Cols - 1
                   GRD1.col = j
                   GRD1.CellBackColor = QBColor(15)
               Next j
               GRD1.col = 16
               If GRD1.Text <> "" Then
                  fnCloseAllFrm100
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Function
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100101_2.Show
                  frm100101_2.Tag = Trim(GRD1.TextMatrix(i, 16))
                  'frm100101_2.cmdOK(6).Visible = False
                  frm100101_2.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Function
               Else
                  MsgBox "無本所案號！", vbInformation
               End If
            End If
         Next i
         Me.Enabled = True
   End Select
End Function

Private Sub cmdQuery_Click()
'   '個人收件區，收受者不可空白
'   If Check1.Visible = False Then
'      If Pub_StrUserSt03 <> "M51" Then
'         If cboII06.Text = "" Then
'            MsgBox "收受者不可空白！", vbExclamation
'            Call cboII06.SetFocus
'            Exit Sub
'         End If
'      End If
'   End If
   
'   '收受者和轉寄者，至少輸入一項
'   If cboII06.Text = "" And cboIR13.Text = "" Then
'      MsgBox "收受者和轉寄者，至少輸入一項！", vbExclamation
'      Exit Sub
'   End If
   
   If cboII06.Text <> "" Then Call cboII06_Validate(False)
   If cboIR13.Text <> "" Then Call cboIr13_Validate(False)
   
   If txtDate(0) = "" And txtDate(1) = "" And _
      txtDate(2) = "" And txtDate(3) = "" And _
      txtDate(4) = "" And txtDate(5) = "" And _
      txtDate(6) = "" And txtDate(7) = "" And _
      cboII05 = "" And Trim(txtII17) = "" And _
      txtII11 = "" And cboII06.Text = "" And cboIR13.Text = "" And _
      (txtPI18 = "") Then 'And txtPI19 = ""
      MsgBox "請至少輸入一項查詢條件！", vbExclamation
      Exit Sub
   End If
   
   Call cboIr13_Validate(False)
   
   Call QueryData
End Sub

'Add By Sindy 2020/7/28
Private Sub cmdShow_Click()
   Frame2.Visible = True
   cmdShow.Visible = False
End Sub
Private Sub cmdClose_Click()
   Frame2.Visible = False
   cmdShow.Visible = True
End Sub
'2020/7/28 END

'Add By Sindy 2017/12/12
Private Sub Combo2_Click()
   Call SetOldValue
End Sub

'沖銷信件,上確認註記
Private Sub cmdM51_Click()
Dim rsA As ADODB.Recordset
Dim strIR10 As String
   
On Error GoTo ErrHand
   
   'Modify By Sindy 2024/5/30 + IPDEPT
   If UCase(Combo2.Text) <> UCase("patent") And _
      UCase(Combo2.Text) <> UCase("TM") And _
      UCase(Combo2.Text) <> UCase("IPDEPT") Then
      MsgBox "patent/TM/IPDEPT 信箱才可使用此功能!!!"
      Exit Sub
   End If
   If MsgBox("確定要沖銷信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
      Exit Sub
   End If
   
   cmdM51.Tag = ""
   strIR10 = UCase(InputBox("請輸入指示沖銷人員的代號？(處理勾選資料...)" & vbCrLf & vbCrLf & "＊空白代表取消", , "QPGMR"))
   If strIR10 = "" Then
      Exit Sub
   End If
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" And GRD1.TextMatrix(i, 14) <> "" Then
         'Add By Sindy 2024/5/30
         If MsgBox(GRD1.TextMatrix(i, 9) & "-" & GRD1.TextMatrix(i, 14) & vbCrLf & _
            "確定要沖銷( " & GetPrjSalesNM(strIR10) & " )的信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            Exit Sub
         End If
         '2024/5/30 END
         '只有一筆資料時才能直接沖銷
         strSql = "SELECT ir01 FROM InputRecord" & _
                  " WHERE ir01=" & GRD1.TextMatrix(i, 9) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                       " and ir08=0"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If rsA.RecordCount = 1 Then
               cmdM51.Tag = "U"
               cnnConnection.BeginTrans
               '同意-沖銷
               strExc(0) = "update InputRecord set " & _
                           " ir08=" & strSrvDate(1) & ",ir09=" & Right("000000" & ServerTime, 6) & ",ir10='" & strIR10 & "'" & _
                           " where ir01=" & GRD1.TextMatrix(i, 9) & _
                             " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                             " and ir08=0"
               Pub_SeekTbLog strExc(0)
               cnnConnection.Execute strExc(0)
               
               '若信件收受者全部已處理或已刪除,主檔就可以掛上msg檔刪除日期,等待AutoBatchDay一個月後刪除實體檔
               strExc(0) = "select ir01 from InputRecord" & _
                           " where ir01=" & GRD1.TextMatrix(i, 9) & _
                             " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                             " and ir08=0" 'and ir05=0 and ir08=0 : 若信件收受者全部已讀取或已刪除
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  '更新"無"Msg檔刪除日期
                  If UCase(Combo2.Text) = UCase("patent") Then
                     strExc(0) = "update PatentInput set" & _
                                 " pi16=" & strSrvDate(1) & _
                                 " where pi01=" & GRD1.TextMatrix(i, 9) & _
                                   " and pi03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                                   " and pi16=0"
                  'Add By Sindy 2020/8/10
                  ElseIf UCase(Combo2.Text) = UCase("TM") Then
                     strExc(0) = "update TMInput set" & _
                                 " ti16=" & strSrvDate(1) & _
                                 " where ti01=" & GRD1.TextMatrix(i, 9) & _
                                   " and ti03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                                   " and ti16=0"
                  'Modify By Sindy 2024/5/30 + IPDEPT
                  ElseIf UCase(Combo2.Text) = UCase("IPDEPT") Then
                     strExc(0) = "update IPDEPTInput set" & _
                                 " ii16=" & strSrvDate(1) & _
                                 " where ii01=" & GRD1.TextMatrix(i, 9) & _
                                   " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                                   " and ii16=0"
                  End If
                  '2020/8/10 END
                  Pub_SeekTbLog strExc(0)
                  cnnConnection.Execute strExc(0)
               End If
               cnnConnection.CommitTrans
            Else
               MsgBox "尚有多筆收受者待確認！"
               Exit Sub
            End If
         End If
         'Exit Sub
      End If
   Next i
   If cmdM51.Tag = "U" Then
      MsgBox "更新成功！"
   End If
   
   Set rsA = Nothing
   Call cmdQuery_Click
   Exit Sub
   
ErrHand:
   Set rsA = Nothing
   cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox " 沖銷失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2018/8/9
'恢復信件
Private Sub cmdM51c_Click()
Dim rsA As ADODB.Recordset
'Dim strIR10 As String
Dim strIR04 As String, strConSql As String
Dim strIR16 As String, strIR16nm As String, strII27 As String, strIR10 As String, strII29 As String 'Add By Sindy 2023/7/11
   
On Error GoTo ErrHand

   If UCase(Combo2.Text) <> UCase("patent") And _
      UCase(Combo2.Text) <> UCase("TM") And _
      UCase(Combo2.Text) <> UCase("ipdept") Then
      MsgBox "patent/TM/ipdept 信箱才可使用此功能!!!"
      Exit Sub
   End If
'   If MsgBox("確定要恢復信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'      Exit Sub
'   End If
   
   strIR16 = "": strIR16nm = ""
   cmdM51c.Tag = ""
'   strIR10 = UCase(InputBox("請輸入欲沖銷人員的代號？(處理勾選資料...)" & vbCrLf & "* 空白代表取消", , "QPGMR"))
'   If strIR10 = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      '還有郵件電子檔及已處理完的,才能恢復信件
      'Modify By Sindy 2020/11/19 And Left(GRD1.TextMatrix(i, 1), 1) = "＊"：往下move判斷
      If GRD1.TextMatrix(i, 0) = "V" And GRD1.TextMatrix(i, 13) <> "" Then
         '收受者記錄檔只有一筆資料時,才能直接恢復
         strExc(0) = "select ir01,ir17,ir16" & _
                     ",decode(substr(ir03,1,1),'F',decode(ir16," & 信件處理狀態 & ",ir16),ir16) as ir16Nm" & _
                     ",ir10,ir04 from InputRecord" & _
                     " where ir01=" & GRD1.TextMatrix(i, 9) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modify By Sindy 2020/9/9
            'Modify By Sindy 2020/11/19
            If rsA.RecordCount = 1 Then
               'Modify By Sindy 2023/5/10 + And Val("" & rsA.Fields("ir17")) = 0: 有處理日期者
               If Left(GRD1.TextMatrix(i, 1), 1) <> "＊" And Val("" & rsA.Fields("ir17")) = 0 Then
                  MsgBox "尚未處理，無需執行恢復！"
                  Exit Sub
               End If
            End If
            '2020/11/19 END
            If rsA.RecordCount > 1 Then
               strIR04 = InputBox(GRD1.TextMatrix(i, 9) & "-" & GRD1.TextMatrix(i, 14) & vbCrLf & "收受人員不只一位，請輸入欲恢復的收受者員編？" & vbCrLf & vbCrLf & "＊若放棄恢復，不用輸入（空白）")
               If Trim(strIR04) = "" Then
                  Exit Sub
               Else
                  strIR04 = UCase(strIR04)
               End If
            Else
               strIR04 = UCase("" & rsA.Fields("ir04"))
               strIR16 = "" & rsA.Fields("ir16") 'Add By Sindy 2023/7/11 原處理狀態
               strIR16nm = "" & rsA.Fields("ir16Nm")
               strIR10 = "" & rsA.Fields("ir10")
            End If
            If MsgBox(GRD1.TextMatrix(i, 9) & "-" & GRD1.TextMatrix(i, 14) & vbCrLf & "確定要恢復( " & GetPrjSalesNM(strIR04) & " )的信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Exit Sub
            End If
            strConSql = ""
            If Trim(strIR04) <> "" Then
               strConSql = " and ir04='" & UCase(strIR04) & "'"
               'Add By Sindy 2023/7/11 讀取原處理狀態
               strExc(0) = "select ir01,ir17,ir16" & _
                           ",decode(substr(ir03,1,1),'F',decode(ir16," & 信件處理狀態 & ",ir16),ir16) as ir16Nm" & _
                           ",ir10 from InputRecord" & _
                           " where ir01=" & GRD1.TextMatrix(i, 9) & _
                             " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & strConSql
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strIR16 = "" & rsA.Fields("ir16") '原處理狀態
                  strIR16nm = "" & rsA.Fields("ir16Nm")
                  strIR10 = "" & rsA.Fields("ir10")
                  'Add By Sindy 2023/7/11
                  If UCase(strIR10) <> UCase(strIR04) And strIR10 <> "" Then '回信鎖此訊息,代表真的有信件出去了才會掛QPGMR
                     If MsgBox("此信件是" & GetPrjSalesNM(strIR10) & "沖銷的，非" & GetPrjSalesNM(strIR04) & "，確定還要恢復嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                        Exit Sub
                     End If
                  End If
                  '2023/7/11 END
               End If
               '2023/7/11 END
            End If
            '2020/9/9 END
            
            cmdM51c.Tag = "U"
            cnnConnection.BeginTrans '*****
            '恢復為無處理
            strExc(0) = "UPDATE InputRecord SET" & _
                        " ir08=0,ir09=null,ir10=null,ir16=null,ir17=0,ir18=null,ir19=null,ir22=null" & _
                        " WHERE ir01=" & GRD1.TextMatrix(i, 9) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & strConSql & _
                          " " 'and ir08>0
            Pub_SeekTbLog strExc(0)
            cnnConnection.Execute strExc(0), strExc(10)
            'Modify By Sindy 2020/9/9
            If Val(strExc(10)) = 0 Then
               MsgBox "無此收受者資料，請重新確認！"
               cnnConnection.RollbackTrans '*****
               Exit Sub
            End If
            '2020/9/9 END
            
            '恢復"無"Msg檔刪除日期=0
            If UCase(Combo2.Text) = UCase("patent") Then
               strExc(0) = "UPDATE PatentInput SET" & _
                           " pi16=0" & _
                           " WHERE pi01=" & GRD1.TextMatrix(i, 9) & _
                             " and pi03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                             " and pi16>0"
            'Add By Sindy 2020/8/10
            ElseIf UCase(Combo2.Text) = UCase("TM") Then
               strExc(0) = "UPDATE TMInput SET" & _
                           " ti16=0" & _
                           " WHERE ti01=" & GRD1.TextMatrix(i, 9) & _
                             " and ti03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                             " and ti16>0"
            'Add By Sindy 2022/8/2
            ElseIf UCase(Combo2.Text) = UCase("ipdept") Then
               'Add By Sindy 2023/7/11 檢查是否要把主檔的處理結果清掉
               If strIR16 <> "" Then
                  strExc(0) = "select ii01,ii03,ii16,ii27,ii29 from ipdeptInput" & _
                              " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                              " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                              " and ii16>0"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Or strIR16 = "9" Then '9.回信
                     If intI = 0 Then
                        '重抓ii27,ii29資料 : 狀況因 intI = 0 and strIR16 = "9" 回信尚未確認沖銷
                        strExc(0) = "select ii01,ii03,ii16,ii27,ii29 from ipdeptInput" & _
                                    " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                                    " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 0 Then
                           MsgBox "查無此信件( " & GRD1.TextMatrix(i, 9) & "-" & ChgSQL(GRD1.TextMatrix(i, 14)) & " )資料！"
                           GoTo ErrHand
                        End If
                     End If
IR16ForUpdate:
                     If Left(PUB_GetST03(strIR04), 2) = "F1" Then
                        strII29 = "" & rsA.Fields("ii29")
                        If strII29 <> "" Then
                           If strII29 = "1" Or strII29 = "2" Or strII29 = "3" Or strII29 = "4" Or strII29 = "5" Or strII29 = "6" Or strII29 = "9" Then
                              strII29 = "1" '輸入
                           ElseIf strII29 = "7" Or strII29 = "10" Then
                              strII29 = "9" '回信
                           ElseIf strII29 = "8" Then
                              strII29 = "2" '不處理
                           ElseIf strII29 = "12" Then
                              strII29 = "5" '已處理
                           Else
                              strII29 = ""
                           End If
                        End If
                        '主檔是記錄一樣的處理狀態才清掉
                        If strIR16 = strII29 Then
                           strExc(0) = "UPDATE ipdeptInput SET" & _
                                       " II29=null" & IIf(strIR16 = "9", ",ii28=null", "") & _
                                       " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                                         " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                                         " and ii29 is not null and ii29<>'11'" 'and ii16>0
                           Pub_SeekTbLog strExc(0)
                           cnnConnection.Execute strExc(0), intI
                        Else
                           If MsgBox("此信件已有處理結果（" & strIR16 & "=" & strIR16nm & "），是否要一併清除？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                              strExc(0) = "UPDATE ipdeptInput SET" & _
                                          " II29=null,ii28=null" & _
                                          " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                                            " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                                            " and ii29 is not null and ii29<>'11'" 'and ii16>0
                              Pub_SeekTbLog strExc(0)
                              cnnConnection.Execute strExc(0), intI
                           End If
                        End If
                     
                     '外專
                     ElseIf Left(PUB_GetST03(strIR04), 2) = "F2" Then
                        strII27 = "" & rsA.Fields("ii27")
                        If strII27 <> "" Then
                           If strII27 = "1" Or strII27 = "2" Or strII27 = "3" Or strII27 = "4" Or strII27 = "5" Or strII27 = "6" Or strII27 = "9" Then
                              strII27 = "1" '輸入
                           ElseIf strII27 = "7" Or strII27 = "10" Then
                              strII27 = "9" '回信
                           ElseIf strII27 = "8" Then
                              strII27 = "2" '不處理
                           ElseIf strII27 = "12" Then
                              strII27 = "5" '已處理
                           Else
                              strII27 = ""
                           End If
                        End If
                        '主檔是記錄一樣的處理狀態才清掉
                        If strIR16 = strII27 Then
                           strExc(0) = "UPDATE ipdeptInput SET" & _
                                       " ii27=null" & IIf(strIR16 = "9", ",ii28=null", "") & _
                                       " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                                         " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                                         " and ii27 is not null and ii27<>'11'" 'and ii16>0
                           Pub_SeekTbLog strExc(0)
                           cnnConnection.Execute strExc(0)
                        Else
                           If MsgBox("此信件已有處理結果（" & strIR16 & "=" & strIR16nm & "），是否要一併清除？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                              strExc(0) = "UPDATE ipdeptInput SET" & _
                                          " ii27=null,ii28=null" & _
                                          " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                                            " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                                            " and ii27 is not null and ii27<>'11'" 'and ii16>0
                              Pub_SeekTbLog strExc(0)
                              cnnConnection.Execute strExc(0)
                           End If
                        End If
                     End If
                  Else
                     If MsgBox("此信件已有處理結果（" & strIR16 & "=" & strIR16nm & "），是否要一併清除？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                        strExc(0) = "select ii01,ii03,ii16,ii27,ii29 from ipdeptInput" & _
                                    " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                                    " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                        GoTo IR16ForUpdate
                     End If
                  End If
               End If
               '2023/7/11 END
               strExc(0) = "UPDATE ipdeptInput SET" & _
                           " ii16=0" & _
                           " WHERE ii01=" & GRD1.TextMatrix(i, 9) & _
                             " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'" & _
                             " and ii16>0"
            End If
            Pub_SeekTbLog strExc(0)
            cnnConnection.Execute strExc(0)
            
            cnnConnection.CommitTrans '*****
         Else
            '檢查是否 "刪除未轉寄" 要恢復
            If Left(Trim(ChgSQL(GRD1.TextMatrix(i, 14))), 1) = "T" Then
               strExc(0) = "select * from TMinput" & _
                           " where ti01=" & GRD1.TextMatrix(i, 9) & _
                             " and ti03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If rsA.Fields("ti07") = "Y" Then 'TI07:刪除未轉寄(Y)
                     If MsgBox(GRD1.TextMatrix(i, 9) & "-" & GRD1.TextMatrix(i, 14) & vbCrLf & "信件已刪除，確定要恢復嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                        Exit Sub
                     Else
                        strExc(0) = "UPDATE TMInput SET" & _
                                    " Ti07=null,Ti08=0,Ti09=null,Ti10=null,Ti16=0" & _
                                    " WHERE ti01=" & GRD1.TextMatrix(i, 9) & _
                                      " and ti03='" & ChgSQL(GRD1.TextMatrix(i, 14)) & "'"
                        Pub_SeekTbLog strExc(0)
                        cnnConnection.Execute strExc(0)
                     End If
                  Else
                     MsgBox "查無資料！"
                  End If
               Else
                  MsgBox "查無資料！"
               End If
            Else
               MsgBox "查無資料！"
            End If
         End If
      End If
   Next i
   If cmdM51c.Tag = "U" Then
      MsgBox "更新成功！"
   End If
   
   Set rsA = Nothing
   Call cmdQuery_Click
   Exit Sub
   
ErrHand:
   Set rsA = Nothing
   cnnConnection.RollbackTrans '*****
   Screen.MousePointer = vbDefault
   MsgBox " 沖銷失敗！" & vbCrLf & Err.Description
End Sub

'Private Sub Form_Activate()
'   If Pub_StrUserSt03 = "M51" Then
'      Check1.Visible = True
'   End If
'End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   Call SetCombo2Limit
   Call SetOldValue
   
   'Add By Sindy 2017/12/22
   cmdShow.Visible = False
   If Pub_StrUserSt03 = "M51" Then
      Frame2.Visible = True
   Else
      Frame2.Visible = False
   End If
   '2017/12/22 END
End Sub

Private Sub SetOldValue()
   If Pub_StrUserSt03 <> "M51" Then
      Me.Height = 7020 '6720 '6500
   End If
   
   '組合下拉選單
   'Modify By Sindy 2016/9/26
'   If UCase(m_PrevForm.Name) = UCase("frm04010518") Or _
'      UCase(m_PrevForm.Name) = UCase("frm04010519") Then
   If Val(Combo2.Tag) <> Combo2.ListIndex Or Combo2.Tag = "" Then
      If UCase(Combo2.Text) = UCase("patent") Then
         'If PUB_GetST03(strUserNum) = "P12" And UCase(m_PrevForm.Name) = UCase("frm04010519") Then '專利處程序
         '分類
         cboII05.Tag = cboII05.ListIndex
         cboII05.Clear
         cboII05.AddItem ""
         'Add By Sindy 2025/1/16
         If strSrvDate(1) >= P業務區劃分啟用日 Then
            Call PUB_AddItemCFPHandler(cboII05, Combo5, , "P")
         Else
         '2025/1/16 ENED
            cboII05.AddItem "1 P程序1"
            cboII05.AddItem "2 P程序2"
         End If
         'Modify By Sindy 2018/6/21
'         cboII05.AddItem "3 亞洲"
'         cboII05.AddItem "4 歐洲"
'         cboII05.AddItem "5 美洋非(單)"
'         cboII05.AddItem "6 美洋非(雙)"
         'Modify by Sindy 2020/3/18
         '109/4/1以後改業務區劃分
         If strSrvDate(1) >= CFP業務區劃分啟用日 Then
            Call PUB_AddItemCFPHandler(cboII05, Combo5)
         Else
         '2020/3/18 END
            cboII05.AddItem "3 美日(單)"
            cboII05.AddItem "4 美日(雙)"
            cboII05.AddItem "5 美日外(單)"
            cboII05.AddItem "6 美日外(雙)"
            '2018/6/21 END
         End If
         cboII05.AddItem "7 其他"
         cboII05.AddItem "8 垃圾信箱"
         'If UCase(m_PrevForm.Name) = UCase("frm04010518") Then
'         If m_WorkType = "0" Then '信箱主檔
'            cboII05.AddItem "9 其他信箱" '國外部匯入
'         End If
'**********  舊分類 **********
         'Add By Sindy 2025/1/16
         If strSrvDate(1) >= P業務區劃分啟用日 Then
            cboII05.AddItem "1 P程序1"
            cboII05.AddItem "2 P程序2"
         End If
         '2025/1/16 ENED
         'Modify by Sindy 2020/3/18
         If strSrvDate(1) >= CFP業務區劃分啟用日 Then
            cboII05.AddItem "3 美日(單)"
            cboII05.AddItem "4 美日(雙)"
            cboII05.AddItem "5 美日外(單)"
            cboII05.AddItem "6 美日外(雙)"
         End If
         '2020/3/18 END
         'Add By Sindy 2018/6/21
         cboII05.AddItem "A 亞洲"
         cboII05.AddItem "B 歐洲"
         cboII05.AddItem "C 美洋非(單)"
         cboII05.AddItem "D 美洋非(雙)"
         '2018/6/21 END
'*****************************
         If cboII05.Tag <> "" And Val(cboII05.Tag) >= 0 And cboII05.ListCount - 1 >= Val(cboII05.Tag) Then
            cboII05.ListIndex = cboII05.Tag
         End If
         '收受者
         cboII06.Tag = cboII06.Text
         cboII06.Clear
         cboII06.AddItem UCase("ipdept")
         cboII06.AddItem UCase("TM")
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  cboII06.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         If cboII06.Tag <> "" Then cboII06.Text = cboII06.Tag
         '轉寄者
         cboIR13.Tag = cboIR13.Text
         cboIR13.Clear
         cboIR13.AddItem UCase("patent")
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  cboIR13.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         If cboIR13.Tag <> "" Then cboIR13.Text = cboIR13.Tag
      ElseIf UCase(Combo2.Text) = UCase("ipdept") Then
      '2016/9/26 END
         '分類
         cboII05.Tag = cboII05.ListIndex
         cboII05.Clear
         cboII05.AddItem ""
         cboII05.AddItem "1 個案"
         cboII05.AddItem "2 外商"
         cboII05.AddItem "3 外專"
         cboII05.AddItem "4 專利處"
         cboII05.AddItem "5 外法"
         cboII05.AddItem "6 新知"
         cboII05.AddItem "7 財務"
         cboII05.AddItem "8 開拓" 'Add By Sindy 2016/6/15
         cboII05.AddItem "Z 其他"
         If cboII05.Tag <> "" And Val(cboII05.Tag) >= 0 And cboII05.ListCount - 1 >= Val(cboII05.Tag) Then
            cboII05.ListIndex = cboII05.Tag
         End If
         '收受者
         cboII06.Tag = cboII06.Text
         cboII06.Clear
         cboII06.AddItem UCase("patent")
         'cboII06.AddItem UCase("account")
         cboII06.AddItem UCase("TM")
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  cboII06.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         If cboII06.Tag <> "" Then cboII06.Text = cboII06.Tag
         '轉寄者
         cboIR13.Tag = cboIR13.Text
         cboIR13.Clear
         cboIR13.AddItem UCase("ipdept")
         cboIR13.AddItem UCase("QPGMR 系統自動收信")
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  cboIR13.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         If cboIR13.Tag <> "" Then cboIR13.Text = cboIR13.Tag
      Else 'TM
         '分類
         cboII05.Tag = cboII05.ListIndex
         cboII05.Clear
         cboII05.AddItem ""
         cboII05.AddItem "1 MCTF"
         cboII05.AddItem "2 大陸案"
         cboII05.AddItem "3 個人"
         cboII05.AddItem "4 非大陸案"
         cboII05.AddItem "5 其他"
'         If m_WorkType = "0" Then '信箱主檔
'            cboII05.AddItem "6 其他信箱"
'         End If
         If cboII05.Tag <> "" And Val(cboII05.Tag) >= 0 And cboII05.ListCount - 1 >= Val(cboII05.Tag) Then
            cboII05.ListIndex = cboII05.Tag
         End If
         '收受者
         cboII06.Tag = cboII06.Text
         cboII06.Clear
         cboII06.AddItem UCase("ipdept")
         cboII06.AddItem UCase("patent")
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  cboII06.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         If cboII06.Tag <> "" Then cboII06.Text = cboII06.Tag
         '轉寄者
         cboIR13.Tag = cboIR13.Text
         cboIR13.Clear
         cboIR13.AddItem UCase("TM")
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  cboIR13.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         If cboIR13.Tag <> "" Then cboIR13.Text = cboIR13.Tag
      End If
      
      GRD1.Clear
      Call SetGrd
      Combo1.ListIndex = 0
      Combo2.Tag = Combo2.ListIndex
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
   End If
   
   DestroyToolTip '清除物件
   Set m_PrevForm = Nothing
   Set frm06010613 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   '                        0    1               2       3       4         5         6               7               8               9       10      11          12      13      14      15                    16          17      18      19      20      21
   arrGridHeadText = Array("V", "轉入日期時間", "主旨", "分類", "收受者", "轉寄者", "轉寄日期時間", "刪除日期時間", "信件日期時間", "II01", "II02", "總收文號", "II06", "II14", "檔名", "轉信條件和系統記錄", "本所案號", "ii08", "ii09", "ir11", "ir12", "Ir01Sort")
   If Check1.Visible = True Then
      '會顯示”轉信條件和系統記錄”欄位
      If UCase(Combo2.Text) = UCase("patent") Or _
         UCase(Combo2.Text) = UCase("tm") Then
         arrGridHeadWidth = Array(200, 1300, 2800, 400, 1300, 0, 1300, 1300, 1300, 0, 0, 0, 0, 0, 0, 800, 1200, 0, 0, 0, 0, 0)
      Else
         '會顯示”總收文號”欄位
         arrGridHeadWidth = Array(200, 1300, 2800, 400, 1300, 0, 1300, 1300, 1300, 0, 0, 900, 0, 0, 0, 800, 1200, 0, 0, 0, 0, 0)
      End If
   'ElseIf cboII06.Text <> "" And cboIR13.Text = "" Then '有輸入收受者並且無轉寄者時,才可以顯示本人刪除日期
   'ElseIf cboII06.Text = "" Then '無輸入收受者,才不顯示本人刪除日期
   ElseIf Not (cboII06.Text <> "" And txtDate(6) <> "") Then '無輸入收受者及刪除日期時,不顯示本人刪除日期
      arrGridHeadWidth = Array(200, 1300, 2800, 400, 1300, 600, 1300, 0, 1300, 0, 0, 0, 0, 0, 0, 800, 1200, 0, 0, 0, 0, 0)
   Else
      arrGridHeadWidth = Array(200, 1300, 2800, 400, 1300, 600, 1300, 1300, 1300, 0, 0, 0, 0, 0, 0, 800, 1200, 0, 0, 0, 0, 0)
   End If
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next iRow
   GRD1.Visible = True
End Sub

Private Sub Grd1_Click()
GRD1.Visible = False
GRD1.row = GRD1.MouseRow
GRD1.col = GRD1.MouseCol
nRow = GRD1.row
nCol = GRD1.col
'Add By Sindy 2024/2/26
If nCol = 1 Then
   nCol = 21
   GRD1.col = 21 '轉入日期時間
End If
'2024/2/26 END
If nRow = 0 Then
   If GRD1.Text <> "V" Then
      If GRD1.Text = "無" Then
         If m_blnColOrderAsc = True Then
            GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
Else
'   '上一筆資料列清除反白
'   If dblPrevRow > 0 Then
'      GRD1.col = 0
'      GRD1.row = dblPrevRow
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = QBColor(15)
'      Next i
'      Call SetColor(dblPrevRow)
'   End If
'   '目前資料列反白
'   GRD1.row = nRow
'   dblPrevRow = GRD1.row
'
'   If GRD1.TextMatrix(GRD1.row, 14) <> "" Then
'      GRD1.col = 0
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   
   GRD1.row = nRow 'GRD1.MouseRow
   dblPrevRow = GRD1.row '記錄目前筆數
   GRD1.col = 0
   If GRD1.TextMatrix(GRD1.row, 14) <> "" Then
      '清除反白
      'If GRD1.TextMatrix(GRD1.row, 0) = "V" Then
      If GRD1.CellBackColor = &HFFC0C0 Then
         Call CancelRowColor(GRD1.row) '清除反白
      Else
         '將點選資料列反白
         GRD1.TextMatrix(GRD1.row, 0) = "V"
         GRD1.col = 0
         GRD1.row = nRow
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub

'Add By Sindy 2016/5/19
Private Sub CancelRowColor(intRow As Integer)
   '清除反白
   GRD1.TextMatrix(intRow, 0) = ""
   GRD1.col = 0
   GRD1.row = intRow
   For j = 0 To GRD1.Cols - 1
      GRD1.col = j
      GRD1.CellBackColor = QBColor(15)
   Next j
   Call SetColor(CDbl(intRow))
End Sub

'開啟附件
Private Sub GRD1_DblClick()
Dim strFileName As String
   
On Error GoTo ErrHand
   
   GRD1.row = GRD1.MouseRow
   GRD1.col = GRD1.MouseCol
   nRow = GRD1.row
   nCol = GRD1.col
   If GRD1.col = 2 Then
      'Modify By Sindy 2018/4/24
'      If GRD1.TextMatrix(dblPrevRow, 14) <> "" And _
'         PUB_ChkOpenLetterLimit(GRD1.TextMatrix(dblPrevRow, 12)) = True Then
      If GRD1.TextMatrix(dblPrevRow, 14) <> "" Then
      '2018/4/24 END
         '讀取檔案
         Screen.MousePointer = vbHourglass
         'Modify By Sindy 2016/10/4
         'strFileName = GRD1.TextMatrix(dblPrevRow, 14)
         strFileName = Mid(GRD1.TextMatrix(dblPrevRow, 13), InStrRev(GRD1.TextMatrix(dblPrevRow, 13), "/") + 1)
         '2016/10/4
         Call PUB_ChkFileTypeOpenExE(strFileName) 'Add By Sindy 2017/9/13
         If GetAttachFile(GRD1.TextMatrix(dblPrevRow, 9), GRD1.TextMatrix(dblPrevRow, 10), GRD1.TextMatrix(dblPrevRow, 14), strFileName, GRD1.TextMatrix(dblPrevRow, 11), m_AttachPath & "\" & strFileName) = True Then
            ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
'         Else
'            MsgBox "無此郵件！", vbInformation
         'Add By Sindy 2024/11/8
         Else
            MsgBox "已逾保留期限！" & vbCrLf & "信件不留存，請至「郵件備份系統」查詢！", vbInformation
            '2024/11/8 END
         End If
         Screen.MousePointer = vbDefault
      End If
   Else
      Call cmdDetail_Click
   End If
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox " 讀取失敗！" & vbCrLf & Err.Description
End Sub

Private Sub SetColor(Optional intSetRow As Double = 0)
   Dim ii As Integer, jj As Integer
   
   With GRD1
   If .Rows > 1 Then
      .Visible = False
      For ii = IIf(intSetRow = 0, 1, intSetRow) To IIf(intSetRow = 0, .Rows - 1, intSetRow)
         '若無實體檔案路徑時,則檔名欄位變灰色
         .row = ii
         If Trim(.TextMatrix(ii, 14)) = "" Then
            .col = 1
            '淺黃色 '灰
            .CellBackColor = &HC0FFFF   '&HE0E0E0
         End If
      Next ii
      If intSetRow = 0 Then .TopRow = 1
      .Visible = True
   End If
   End With
End Sub

Private Function GetAttachFile(ByVal strPkey1 As String, ByVal strPkey2 As String, ByVal strPkey3 As String, _
                               ByRef pFileName As String, ByVal strCP09 As String, _
                               Optional pSavePath As String) As Boolean
Dim stAttPath As String

On Error GoTo ErrHnd
   
   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
   Else
      '改傳完整的檔案路徑:路徑+檔名
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      stAttPath = pSavePath
   End If
   
   'Modify By Sindy 2019/5/2
   GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'   'Modify By Sindy 2016/9/30
''   If UCase(m_PrevForm.Name) = UCase("frm04010518") Or _
''      UCase(m_PrevForm.Name) = UCase("frm04010519") Then
'   If UCase(Combo2.Text) = UCase("patent") Then
'      GetAttachFile = PUB_GetAttachFile_PImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'   'Add By Sindy 2019/4/15
'   ElseIf UCase(Combo2.Text) = UCase("tm") Then
'      GetAttachFile = PUB_GetAttachFile_TImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'   Else
'   '2016/9/30 END
'      If strCP09 <> "" Then '個案
'         GetAttachFile = PUB_GetAttachFile_CPP(strCP09, pFileName, stAttPath, True)
'         'ADD BY SONIA 2016/4/8 因之前放入個案,故個案讀不到加入下面語法
'         If GetAttachFile = False Then
'            GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'         End If
'         'END 2016/4/8
'      Else
'         GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'      End If
'   End If
   
   Exit Function
   
ErrHnd:
   If Err.NUMBER = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 And _
      (GRD1.MouseCol = 1 Or GRD1.MouseCol = 2 Or GRD1.MouseCol = 4 Or _
      GRD1.MouseCol = 15 Or GRD1.MouseCol = 3) Then
      If iRow <> GRD1.MouseRow Or iCol <> GRD1.MouseCol Then
         'Modify By Sindy 2017/12/22
         If GRD1.MouseCol = 1 Then
            'GRD1.ToolTipText = "信件編號:" & GRD1.TextMatrix(GRD1.MouseRow, 9) & "-" & GRD1.TextMatrix(GRD1.MouseRow, 14)
            CreateToolTip GetHWndForToolTip(GRD1), "信件編號:" & GRD1.TextMatrix(GRD1.MouseRow, 9) & "-" & GRD1.TextMatrix(GRD1.MouseRow, 14)
         '2017/12/22 END
         ElseIf GRD1.MouseCol = 2 Then
            CreateToolTip GetHWndForToolTip(GRD1), PUB_GetInputData(GRD1.TextMatrix(GRD1.MouseRow, 9), GRD1.TextMatrix(GRD1.MouseRow, 14), "主旨")
         ElseIf GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
            'GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
            CreateToolTip GetHWndForToolTip(GRD1), GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
         End If
         iRow = GRD1.MouseRow
         iCol = GRD1.MouseCol
      End If
   End If
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index).Text <> "" Then
      If ChkDate(txtDate(Index)) = False Then
         Call txtDate_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
   If Index = 0 Or Index = 2 Or Index = 4 Then
      If txtDate(Index) <> "" And txtDate(Index + 1) = "" Then
         txtDate(Index + 1) = txtDate(Index)
      End If
      If Val(txtDate(Index)) > Val(txtDate(Index + 1)) Then
         txtDate(Index + 1) = txtDate(Index)
      End If
   ElseIf Index = 1 Or Index = 3 Or Index = 5 Then
      If txtDate(Index) <> "" And txtDate(Index - 1) = "" Then
         txtDate(Index - 1) = txtDate(Index)
      End If
      If txtDate(Index - 1) <> "" And txtDate(Index) <> "" Then
         If RunNick2(txtDate(Index - 1), txtDate(Index)) Then
            Call txtDate_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

'Private Sub txtUsernum_Change()
'   If Len(txtUsernum) >= 5 Then
'      lblUserName = GetStaffName(txtUsernum, True)
'   Else
'      lblUserName = ""
'   End If
'End Sub
'
'Private Sub txtUsernum_GotFocus()
'   TextInverse txtUsernum
'End Sub
'
'Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

Private Sub txtII11_GotFocus()
   TextInverse txtII11
End Sub

Private Sub txtII11_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtII11
End Sub

'Add By Sindy 2022/7/27
Private Sub txtII17_Change()
   PUB_RefreshText txtII17
End Sub

Private Sub txtII17_GotFocus()
   TextInverse txtII17
End Sub

Private Sub txtII17_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   txtII17.ToolTipText = "文字貼上：請使用Ctrl+V 文字複製：請使用Ctrl+C"
End Sub

Private Sub txtII17_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtII17
End Sub

'Add By Sindy 2017/12/22
Private Sub txtII03_GotFocus()
   TextInverse txtII03
End Sub
Private Sub txtII03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'2017/12/22 END

Private Sub txtPI18_GotFocus()
   TextInverse txtPI18
   CloseIme
End Sub

Private Sub txtPI18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Private Sub txtPI18_Validate(Cancel As Boolean)
'   If txtPI18 <> "" Then
'      txtPI18 = UCase(txtPI18)
''      If ChkSysName(txtPI18) = True Then
'         If txtPI18 <> "P" And txtPI18 <> "PS" And _
'            txtPI18 <> "CFP" And txtPI18 <> "CPS" Then
'            MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
'            Cancel = True
'         End If
''      Else
''         Cancel = True
''      End If
'   End If
'   If Cancel Then TextInverse txtPI18
'End Sub

Private Sub txtPI19_GotFocus()
   TextInverse txtPI19
End Sub

Private Sub txtPI20_GotFocus()
   TextInverse txtPI20
End Sub

Private Sub txtPI20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPI20_LostFocus()
   If txtPI18 <> "" And txtPI19 <> "" And txtPI20 = "" Then txtPI20 = "0"
End Sub

Private Sub txtPI21_GotFocus()
   TextInverse txtPI21
End Sub

Private Sub txtPI21_LostFocus()
   If txtPI18 <> "" And txtPI19 <> "" And txtPI21 = "" Then txtPI21 = "00"
End Sub

Private Sub txtPI21_Validate(Cancel As Boolean)
Dim strPI05 As String
Dim strPI06 As String
   
   If txtPI18 <> "" And txtPI19 <> "" Then
      If txtPI20 = "" Then txtPI20 = "0"
      If txtPI21 = "" Then txtPI21 = "00"
   End If
End Sub
