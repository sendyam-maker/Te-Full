VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010616 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專系統收件區"
   ClientHeight    =   6760
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6760
   ScaleWidth      =   9040
   Begin VB.CommandButton cmdCanclReMail 
      BackColor       =   &H00FFFF80&
      Caption         =   "取消回信"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7950
      Style           =   1  '圖片外觀
      TabIndex        =   69
      Top             =   6420
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdMainQuy 
      Caption         =   "未處理查詢"
      Height          =   330
      Left            =   2280
      Style           =   1  '圖片外觀
      TabIndex        =   67
      Top             =   2460
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   345
      Left            =   3210
      TabIndex        =   36
      Top             =   2460
      Width           =   5775
      Begin VB.CommandButton cmdProDel 
         Caption         =   "已處理"
         Height          =   330
         Left            =   1770
         TabIndex        =   64
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdReMail 
         Caption         =   "回信作業"
         Height          =   330
         Left            =   840
         TabIndex        =   9
         Top             =   0
         Width           =   885
      End
      Begin VB.CommandButton cmdInput 
         Caption         =   "輸入"
         Height          =   330
         Left            =   225
         TabIndex        =   8
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmdMgOK 
         BackColor       =   &H00C0C0FF&
         Caption         =   "主管核准"
         Height          =   330
         Left            =   4860
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   0
         Width           =   885
      End
      Begin VB.CommandButton cmdRecall 
         Caption         =   "回覆確收"
         Height          =   330
         Left            =   3915
         TabIndex        =   12
         Top             =   0
         Width           =   885
      End
      Begin VB.CommandButton cmdPDF 
         Caption         =   "歸卷"
         Height          =   330
         Left            =   3300
         TabIndex        =   11
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmdNotProDel 
         Caption         =   "不處理"
         Height          =   330
         Left            =   2535
         TabIndex        =   10
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "轉寄副本:"
      Height          =   315
      Left            =   30
      Style           =   1  '圖片外觀
      TabIndex        =   63
      Top             =   990
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "轉寄收受者:"
      Height          =   315
      Left            =   30
      Style           =   1  '圖片外觀
      TabIndex        =   62
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdWait 
      Caption         =   "待歸檔"
      Height          =   330
      Left            =   7320
      TabIndex        =   61
      Top             =   6840
      Width           =   705
   End
   Begin VB.Frame FrameCont 
      Caption         =   "FrameCont"
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   2670
      TabIndex        =   52
      Top             =   1080
      Visible         =   0   'False
      Width           =   3525
      Begin VB.TextBox txtCOR01 
         Height          =   270
         Left            =   2220
         MaxLength       =   9
         TabIndex        =   57
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelCont 
         Caption         =   "選擇往來記錄"
         Height          =   300
         Left            =   2160
         TabIndex        =   56
         Top             =   30
         Width           =   1320
      End
      Begin VB.TextBox txtCOR03 
         Height          =   270
         Left            =   930
         MaxLength       =   9
         TabIndex        =   53
         Top             =   60
         Width           =   1215
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Left            =   90
         TabIndex        =   55
         Top             =   330
         Width           =   3345
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         Caption         =   "X,Y,R編號:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   54
         Top             =   90
         Width           =   975
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "待核准信件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   750
      TabIndex        =   46
      Top             =   90
      Width           =   2025
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm06010616.frx":0000
      Left            =   630
      List            =   "frm06010616.frx":0019
      Style           =   2  '單純下拉式
      TabIndex        =   43
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   330
      Index           =   1
      Left            =   1500
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   2460
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      Height          =   345
      Left            =   2940
      TabIndex        =   37
      Top             =   7080
      Visible         =   0   'False
      Width           =   1635
      Begin VB.CommandButton cmdback 
         Caption         =   "退回"
         Height          =   330
         Left            =   810
         TabIndex        =   27
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAgree 
         Caption         =   "同意"
         Height          =   330
         Left            =   50
         TabIndex        =   26
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "刪除"
      Height          =   330
      Left            =   4590
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   4980
      TabIndex        =   33
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSendMail 
         Caption         =   "立即寄發通知信"
         Height          =   300
         Left            =   2490
         TabIndex        =   29
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         Caption         =   "程式結束後”寄發通知信”或"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   60
         TabIndex        =   34
         Top             =   30
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdUpdRow 
      Caption         =   "轉寄"
      Height          =   330
      Left            =   5364
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "信件狀況"
      Height          =   330
      Left            =   7065
      TabIndex        =   5
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   330
      Left            =   3390
      TabIndex        =   1
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "記錄查詢"
      Height          =   330
      Left            =   6156
      TabIndex        =   4
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8010
      TabIndex        =   6
      Top             =   0
      Width           =   765
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm06010616.frx":0042
      Height          =   3560
      Left            =   60
      TabIndex        =   28
      Top             =   2820
      Width           =   8870
      _ExtentX        =   15646
      _ExtentY        =   6279
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|符號|確|收信日期時間|本所案號|主旨|收受者|轉寄者|轉寄日期時間|讀取日期時間|檔名|總收文號|處理原因"
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
      _Band(0).Cols   =   13
   End
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   3060
      TabIndex        =   58
      Top             =   630
      Width           =   1035
      Begin VB.ComboBox cboReason 
         Height          =   300
         ItemData        =   "frm06010616.frx":0057
         Left            =   600
         List            =   "frm06010616.frx":0059
         Style           =   2  '單純下拉式
         TabIndex        =   59
         Top             =   0
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "處理原因："
         Height          =   180
         Left            =   90
         TabIndex        =   60
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "歸卷方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   3930
      TabIndex        =   49
      Top             =   1680
      Width           =   1485
      Begin VB.OptionButton Option1 
         Caption         =   "往來記錄"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   51
         Top             =   390
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "卷宗區"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   50
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "轉寄後不顯示"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   6630
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Frame FrameRecv 
      Caption         =   "FrameRecv"
      ForeColor       =   &H8000000F&
      Height          =   795
      Left            =   5430
      TabIndex        =   39
      Top             =   1710
      Width           =   3525
      Begin VB.CommandButton cmdSelCp09 
         Caption         =   "選擇總收文號"
         Height          =   300
         Left            =   2220
         TabIndex        =   24
         Top             =   300
         Width           =   1230
      End
      Begin VB.TextBox txtPI21 
         Height          =   270
         Left            =   2460
         MaxLength       =   2
         TabIndex        =   22
         Top             =   30
         Width           =   375
      End
      Begin VB.TextBox txtPI19 
         Height          =   270
         Left            =   1365
         MaxLength       =   6
         TabIndex        =   20
         Top             =   30
         Width           =   855
      End
      Begin VB.TextBox txtPI18 
         Height          =   270
         Left            =   870
         MaxLength       =   3
         TabIndex        =   19
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox txtPI20 
         Height          =   270
         Left            =   2220
         MaxLength       =   1
         TabIndex        =   21
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox txtRecvNo 
         Height          =   270
         Left            =   870
         MaxLength       =   9
         TabIndex        =   23
         Top             =   285
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "案件性質名稱"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   870
         TabIndex        =   66
         Top             =   570
         Width           =   1170
      End
      Begin VB.Label Label7 
         Caption         =   "本所案號:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   41
         Top             =   60
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "總收文號:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   40
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "符號／"
      Height          =   180
      Left            =   90
      TabIndex        =   68
      Top             =   2430
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "備註：雙擊”主旨”選取狀況下開啟信件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   30
      Top             =   6510
      Width           =   3825
   End
   Begin MSForms.TextBox txtIR20_show 
      Height          =   945
      Left            =   6900
      TabIndex        =   65
      Top             =   720
      Width           =   2055
      VariousPropertyBits=   -1466941409
      BackColor       =   -2147483644
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "3625;1667"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCC 
      Height          =   300
      Left            =   870
      TabIndex        =   15
      Top             =   7110
      Width           =   1590
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2805;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtIR20 
      Height          =   945
      Left            =   4110
      TabIndex        =   25
      Top             =   720
      Width           =   2775
      VariousPropertyBits=   -1466941413
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "4895;1667"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox List1 
      Height          =   315
      Left            =   1140
      TabIndex        =   16
      Top             =   690
      Width           =   1875
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "3307;556"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextContext 
      Height          =   765
      Left            =   60
      TabIndex        =   17
      Top             =   1530
      Width           =   3825
      VariousPropertyBits=   -1466941415
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "6747;1349"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      TabIndex        =   0
      Top             =   360
      Width           =   1605
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "2831;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboII06 
      Height          =   300
      Left            =   870
      TabIndex        =   14
      Top             =   6810
      Width           =   1590
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2805;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "* 代表有轉寄給他人"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   35
      Top             =   6510
      Width           =   1635
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Left            =   30
      TabIndex        =   48
      Top             =   30
      Visible         =   0   'False
      Width           =   1725
      VariousPropertyBits=   746604571
      Size            =   "3043;503"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblTM 
      Caption         =   "輸入時一併將附件匯入卷宗區, 檔案位置："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   2850
      TabIndex        =   45
      Top             =   6750
      Width           =   4425
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "顏色："
      Height          =   180
      Left            =   90
      TabIndex        =   44
      Top             =   2610
      Width           =   540
   End
   Begin VB.Label Label8 
      Caption         =   "轉寄內容："
      Height          =   255
      Left            =   90
      TabIndex        =   42
      Top             =   1320
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "註:收受者點二下即可移除"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   2760
      TabIndex        =   38
      Top             =   420
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   405
      Width           =   900
   End
   Begin VB.Label LblTotCnt 
      Caption         =   "總筆數:"
      ForeColor       =   &H00C00000&
      Height          =   200
      Left            =   6330
      TabIndex        =   31
      Top             =   6510
      Width           =   1610
   End
   Begin VB.Label LblSec2Query 
      BackColor       =   &H0080FFFF&
      Height          =   330
      Left            =   660
      TabIndex        =   47
      Top             =   30
      Visible         =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "frm06010616"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/22 Form2.0已修改
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Dim m_AttachPath As String
Dim nCol As Long, nRow As Long
Dim m_iRow As Integer
Public m_strIR01 As String, m_strIR02 As String, m_strIR03 As String, m_strIR04 As String
Public m_strPi12 As String
Public cmdState As Integer, bolQuery As Boolean '紀錄作用按鍵
Dim m_TxtIR20 As String
Public m_AppNo As String
Public m_RegNo As String
Public m_strFullFileName As String, m_strFullFileName_order As String
Public m_LstEmp As String
'Dim Tmpfrm060504 As Form '案件命名
Dim m_WorkEmpList As String '工作代理 Add By Sindy 2022/9/5


''呼叫案件命名
'Public Sub SetTmpfrm060504(ByRef fm As Form)
'   Set Tmpfrm060504 = fm
'End Sub

Private Sub cboReason_Click()
   If cboReason.List(cboReason.ListIndex) = "已處理：" Or cboReason.List(cboReason.ListIndex) = "不處理：" Then
   Else
      If txtIR20.Text = "" Then
         If Trim(cboReason.List(cboReason.ListIndex)) = "其他" Then
            txtIR20.Text = "其他，"
         Else
            txtIR20.Text = Trim(cboReason.List(cboReason.ListIndex))
         End If
      End If
   End If
   txtPI20_GotFocus
   txtIR20.SetFocus
   cboReason.ListIndex = -1
   cboReason.Width = 315
End Sub
Private Sub cboReason_DropDown()
   cboReason.Width = 3645
End Sub
Private Sub cboReason_LostFocus()
   cboReason.Width = 315
End Sub

'Add By Sindy 2022/6/2
Private Sub Check2_Click()
   Call SetCombo1 'Add By Sindy 2025/8/6
   'Call cmdQuery_Click 'Modify By Sindy 2025/9/2 會重覆查第2次, mark
End Sub

'Add By Sindy 2024/4/24 主管才能取消回信,僅一筆資料做勾選時才會顯示此按鍵
Private Sub cmdCanclReMail_Click()
Dim bolConn As Boolean

   If MsgBox("確定要取消回信嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
      Exit Sub
   End If
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         cnnConnection.BeginTrans: bolConn = True
                  
         '恢復為無處理
         strExc(0) = "UPDATE ipdeptInput SET" & _
                     " ii27=null,ii28=null" & _
                     " WHERE ii01=" & GRD1.TextMatrix(i, 10) & _
                       " and ii03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and ii16=0" '必須尚無"可刪除日期"
         Pub_SeekTbLog strExc(0)
         cnnConnection.Execute strExc(0), intI
         If intI = 0 Then
            GoTo ErrHand
         End If
         
         strExc(0) = "UPDATE InputRecord SET" & _
                     " ir08=0,ir09=null,ir10=null,ir16=null,ir17=0,ir18=null,ir19=null,ir22=null" & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                     " and ir02=" & GRD1.TextMatrix(i, 11) & _
                     " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                     " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                     " and ir08=0"
         Pub_SeekTbLog strExc(0)
         cnnConnection.Execute strExc(0), intI
         
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
         Exit For
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "信件取消回信失敗！" & vbCrLf & Err.Description
End Sub

'刪除鍵
Private Sub cmdDelRow_Click()
Dim bolSelectRow As Boolean
Dim strUpdTime As String
Dim bolConn As String
   
   bolSelectRow = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
            Exit Sub
         End If
         bolSelectRow = True
         Exit For
      End If
   Next i
   If bolSelectRow = False Then
      MsgBox "請至少勾選一筆要刪除的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If MsgBox("確定要刪除信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         'Add By Sindy 2022/2/22 針對分類有*號時，要增加提醒訊息。
         Dim strBox As String
         If InStr(GRD1.TextMatrix(i, 28), "*") > 0 Then
            strBox = Mid(GRD1.TextMatrix(i, 28), InStr(GRD1.TextMatrix(i, 28), "*") - 1, 1)
            If MsgBox(GRD1.TextMatrix(i, 5) & vbCrLf & vbCrLf & _
                   "提醒：信件狀態「" & IIf(strBox = "F", "國外部", IIf(strBox = "P", "專利處", "商標處")) & "」是直接刪除的，若人員處理信件時發現非該單位信件，請轉寄回該信箱。" & vbCrLf & vbCrLf & _
                   "確定要刪除信件嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
               'Call CancelRowColor(i) '清除反白
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
         '2022/2/22 END
         
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         
         '刪除
         strExc(0) = "update InputRecord set " & _
                     " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'"
         'Add By Sindy 2021/1/22
         If txtIR20 <> m_TxtIR20 Then
            strExc(0) = strExc(0) & _
                        ",IR20=decode(IR20,null,'','" & strSrvDate(2) & "-" & strUserName & ":')||'" & ChgSQL(txtIR20) & "'||decode(IR20,null,'',';'||chr(13)||chr(10)||IR20)"
         End If
         '2021/1/22 END
         strExc(0) = strExc(0) & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                       " and ir08=0"
         cnnConnection.Execute strExc(0)
         Call SaveInputRecord(i, False)
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Sub

Private Function CheckDataValid(Optional strII06 As String = "") As Boolean
Dim intTotList As Integer
Dim strChkEmp As String, strChkName As String
Dim ArrStr As Variant
   
   CheckDataValid = False
   TextContext.Enabled = False
   '檢查收受者是否重覆
   If strII06 <> "" Or List1.ListCount > 0 Then
      '欲檢查幾個收受者
      If strII06 <> "" Then
         intTotList = 0
      Else
         intTotList = List1.ListCount - 1
      End If
      Screen.MousePointer = vbHourglass
      For i = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(i, 0) = "V" Then
            For j = 0 To intTotList
               If strII06 <> "" Then
                  ArrStr = Split(strII06, " ")
                  strChkEmp = ArrStr(0) 'Left(strII06, 5)
                  If UBound(ArrStr) > 0 Then
                     strChkName = ArrStr(1)
                  Else
                     strChkName = ArrStr(0) 'Trim(Mid(strII06, 6))
                  End If
               Else
                  'strChkEmp = Left(List1.List(j), 5)
                  'strChkName = Trim(Mid(List1.List(j), 6))
                  ArrStr = Split(List1.List(j), " ")
                  strChkEmp = ArrStr(0)
                  If UBound(ArrStr) > 0 Then
                     strChkName = ArrStr(1)
                  Else
                     strChkName = ArrStr(0)
                  End If
               End If
               
               '非外專承辦組，程序組人員
               If PUB_GetST03(strChkEmp) <> "F22" And PUB_GetST03(strChkEmp) <> "F23" Then
                  TextContext.Enabled = True
               End If
               
               strExc(0) = "select ir04 from inputrecord" & _
                           " where ir01=" & GRD1.TextMatrix(i, 10) & _
                             " and ir02=" & GRD1.TextMatrix(i, 11) & _
                             " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                             " and ir04='" & strChkEmp & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Screen.MousePointer = vbDefault
                  MsgBox "收受者（" & strChkName & "）已收過此郵件！" & vbCrLf & _
                         GRD1.TextMatrix(i, 3) & vbCrLf & _
                         GRD1.TextMatrix(i, 5), vbExclamation
                  Me.List1.SetFocus
                  Exit Function
               End If
            Next j
         End If
      Next i
      Screen.MousePointer = vbDefault
   End If
   
   CheckDataValid = True
End Function

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
   Call ClearText
End Sub

'信件狀況
Private Sub cmdDetail_Click()
   cmdState = 99
   Call PubShowNextData
End Sub

Public Function PubShowNextData() As Boolean
Dim rsA As New ADODB.Recordset
Dim stFileName As String
Dim hLocalFile As Long
Dim strCaseNo As String

Select Case cmdState
Case 1 '進度
   If bolQuery = True Then
      Me.Enabled = False
      For i = 1 To GRD1.Rows - 1
         GRD1.col = 0
         GRD1.row = i
         If Trim(GRD1.Text) = "V" Then
            'Modify By Sindy 2022/10/14
            If Check2.Value = 1 Then '待核准信件資料區
               GRD1.col = 0
               GRD1.Text = ""
               Call CancelRowColor(i)
            End If
            '2022/10/14 END
'            For j = 0 To GRD1.Cols - 1
'                GRD1.col = j
'                GRD1.CellBackColor = QBColor(15)
'            Next j
            GRD1.col = 4: strCaseNo = ""
            If GRD1.TextMatrix(i, 4) = "" Then
               If txtPI18 <> "" And txtPI19 <> "" Then
                  If txtPI20 = "" Then txtPI20 = "0"
                  If txtPI21 = "" Then txtPI21 = "00"
                  strCaseNo = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
               End If
            Else
               strCaseNo = GRD1.TextMatrix(i, 4)
            End If
            'If GRD1.Text <> "" Then
            If strCaseNo <> "" Then
'                'Modified by Morgan 2016/3/24 排除母層是共同查詢
'                If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
'                   fnCloseAllFrm100 'Added by Morgan 2016/2/22
'                End If
'                'end 2016/3/24

               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Function
               End If
               Screen.MousePointer = vbHourglass
               bolQuery = False
               frm100101_2.Show
               frm100101_2.Tag = strCaseNo 'Pub_RplStr(GRD1.TextMatrix(i, 4))
               frm100101_2.cmdok(5).Visible = False '下一筆不顯示
               frm100101_2.StrMenu
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Function
            Else
               MsgBox "無本所案號！", vbExclamation, "警告！"
            End If
         End If
      Next i
      Me.Enabled = True
   End If
'Case 4 '完整卷宗
'   Screen.MousePointer = vbHourglass
'   frm100101_L.m_strKey = lblCaseNo.Caption
'   frm100101_L.SetParent Me
'   If frm100101_L.QueryData = True Then
'      frm100101_L.Show
'      Me.Hide
'   Else
'      Unload frm100101_L
'   End If
'   Screen.MousePointer = vbDefault
Case Else
   PubShowNextData = False
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" And GRD1.TextMatrix(i, 16) <> "" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
            Exit Function
         End If
         PubShowNextData = True
         '明細資料
         frm06010613_1.m_II01 = GRD1.TextMatrix(i, 10)
         frm06010613_1.m_II02 = GRD1.TextMatrix(i, 11)
         frm06010613_1.m_II03 = GRD1.TextMatrix(i, 16)
         frm06010613_1.m_II19 = GRD1.TextMatrix(i, 12)
         'Modify By Sindy 2017/12/25 Mark
'         Call CancelRowColor(i)
         frm06010613_1.CmdNext.Enabled = False
'         For j = i To grd1.Rows - 1
'            If grd1.TextMatrix(j, 0) = "V" And grd1.TextMatrix(j, 16) <> "" Then
'               frm06010613_1.cmdNext.Enabled = True
'               Exit For
'            End If
'         Next j
         Call frm06010613_1.SetParent(Me)
         frm06010613_1.Show
         frm06010613_1.QueryData
         Me.Hide
         Exit Function
      End If
   Next i
End Select
End Function

Private Sub cmdExit_Click()
Dim strMsg As String

   'Add By Sindy 2022/9/5 為他人職代時,檢查是否有待處理信件,若有要彈訊息
   If GetWorkEmpListData(strMsg, False) = True Then
      If MsgBox(strMsg & "，要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         Exit Sub
      End If
   End If
   '2022/9/5 END
   
   Unload Me
End Sub

Public Function QueryData(Optional bolShowMsg As Boolean = True) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strUser As String
   
On Error GoTo ErrHand
   
   m_blnColOrderAsc = True
   QueryData = False
   Screen.MousePointer = vbHourglass
   
   Call ClearText
   Frame2.Visible = False
   Label9.Visible = False
   Combo2.Visible = False
   Label6.Visible = False
   Frame3.Visible = False '同意/退回
   GRD1.Clear
   List1.Clear
   Call SetGrd
   cmdCanclReMail.Visible = False 'Add By Sindy 2024/4/24
   
   LblSec2Query.Visible = False
   cmdWait.Visible = True '待歸檔按鍵
   cmdMgOK.Enabled = True '主管核准
   
   '待核准信件資料區:1.輸入 2.不處理 9.回信 5.已處理
   'Modified by Morgan 2023/4/17 +ir21
   strSql = "select '' V,IR23 符號,GetInputRecordReply(ir01,ir02,ir03) 確,sqldatet(Ii12)||' '||sqltime6(Ii13) 收信日期時間,decode(Ii23,null,'',Ii23||'-'||Ii24||'-'||Ii25||'-'||Ii26) 本所案號,Ii17 主旨" & _
            ",'' 收受者,s2.st02||'-'||decode(ii27," & 外專信件處理結果 & ",ii27) 處理人員" & _
            ",sqldatet(ir17)||' '||sqltime6(ir18) 處理日期時間" & _
            ",sqldatet(IR05)||' '||sqltime6(IR06) 讀取日期時間" & _
            ",IR01,IR02,Ii18,IR04,Ii06,Ii14,Ii03 檔名,Ii08,Ii09,ir11,ir12,Ii12,Ii05,IR16,decode(IR21,null,ii19,IR21) 總收文號,decode(FO02,null,IR20,FO02) 處理原因,ir24,ir19,getmailbox(Ii01,Ii03) 信箱來源,ir21" & _
            " From inputrecord,IPDeptInput,staff s1,staff s2,form" & _
            " where IR08=0 and IR16 in('1','2','9','5')" & _
            " and IR01=Ii01(+) and IR02=Ii02(+) and IR03=Ii03(+) and ii08>0" & _
            " and ir13=s1.st01(+)" & _
            " and ir19=s2.st01(+)" & _
            " and ir22='" & Trim(Left(Combo1, 6)) & "' and ir20=FO01(+)" & _
            " order by decode(IR16,'9',0,1) desc,IR23 asc,ir11 asc,ir12 asc"
   If Check2.Value = 1 Then '待核准信件資料區
      Check2.BackColor = &H80FFFF
      LblSec2Query.Visible = True
      Frame2.Visible = True
      cmdWait.Visible = False
      
      cmdInput.Enabled = False
      cmdReMail.Enabled = False
      cmdNotProDel.Enabled = False
      cmdProDel.Enabled = False
      cmdPDF.Enabled = False
      cmdRecall.Enabled = False '回覆確收
   Else
      cmdMgOK.Enabled = False '主管核准
      cmdInput.Enabled = True
      cmdReMail.Enabled = True
      cmdNotProDel.Enabled = True
      cmdPDF.Enabled = True
      cmdRecall.Enabled = True '回覆確收
      SetCmdStatus
      
      '顯示[待核准信件]的筆數
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If InStr(Check2.Caption, "(") > 0 Then Check2.Caption = Left(Check2.Caption, InStr(Check2.Caption, "(") - 1)
      If intI = 1 Then
         Check2.Caption = Check2.Caption & "(" & RsTemp.RecordCount & ")"
      Else
         Check2.Caption = Check2.Caption & "(0)"
      End If
      Check2.BackColor = &H8000000F
      
      Label9.Visible = True
      Combo2.Visible = True
      Label6.Visible = True
      
      Frame2.Visible = True
      
      '未處理和 4.歸卷 9.回信 3.退回 8.退回2(2次確認退回)
      'Modified by Morgan 2023/4/12 +ir21,IR16條件加 8(退回2), "decode(ir16,'3',"-->"decode(sign(instr('3,8',ir16)),1,"
      strSql = "select '' V,IR23 符號,GetInputRecordReply(ir01,ir02,ir03) 確,sqldatet(Ii12)||' '||sqltime6(Ii13) 收信日期時間,decode(Ii23,null,'',Ii23||'-'||Ii24||'-'||Ii25||'-'||Ii26) 本所案號,Ii17 主旨,'' 收受者" & _
               ",decode(ir16,'3',s3.st02,decode(ii27||ir16,null,decode(ir15,'Y',decode(length(ir03),5,decode(substr(ir03,1,1),'T','TM','P','Patent','IPDept'),'IPDept'),s1.st02),s2.st02))||decode(sign(instr('3,8',ir16)),1,decode(ir16," & 信件處理狀態 & ",ir16),decode(ii27," & 外專信件處理結果 & ",ii27)) 處理或轉寄者" & _
               ",decode(ii27,null,sqldatet(Ii08)||' '||sqltime6(Ii09),sqldatet(ir17)||' '||sqltime6(ir18)) 處理或轉寄日期時間" & _
               ",sqldatet(IR05)||' '||sqltime6(IR06) 讀取日期時間" & _
               ",IR01,IR02,Ii18,IR04,Ii06,Ii14,Ii03 檔名,Ii08,Ii09,ir11,ir12,Ii12,Ii05,IR16,decode(IR21,null,ii19,IR21) 總收文號,decode(FO02,null,IR20,FO02) 處理原因,ir24,ir19,getmailbox(Ii01,Ii03) 信箱來源,ir21" & _
               " From inputrecord,IPDeptInput,staff s1,staff s2,staff s3,form" & _
               " where IR08=0 and (IR16 is null or IR16 in('4','9','3','8'))" & _
               " and IR01=Ii01 and IR02=Ii02 and IR03=Ii03 and ii08>0" & _
               " and ir13=s1.st01(+)" & _
               " and ir19=s2.st01(+)" & _
               " and ir22=s3.st01(+)" & _
               " and IR04='" & Trim(Left(Combo1, 6)) & "' and ir20=FO01(+)" & _
               " order by ir11 desc,ir12 desc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   LblTotCnt.Caption = "總筆數: "
   If Check2.Value = 1 Then
      If InStr(Check2.Caption, "(") > 0 Then Check2.Caption = Left(Check2.Caption, InStr(Check2.Caption, "(") - 1)
   End If
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      LblTotCnt.Caption = "總筆數: " & rsTmp.RecordCount
      QueryData = True
      For i = 1 To GRD1.Rows - 1
         '有無轉寄給他人:*.有
         strSql = "SELECT ir04 FROM inputrecord" & _
                  " WHERE ir01=" & GRD1.TextMatrix(i, 10) & _
                    " and ir02=" & GRD1.TextMatrix(i, 11) & _
                    " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                    " and (ir11<>" & GRD1.TextMatrix(i, 19) & " or ir12<>" & GRD1.TextMatrix(i, 20) & ")" & _
                    " and (ir13='" & Trim(Left(Combo1, 6)) & "' or ir14='" & Trim(Left(Combo1, 6)) & "')" & _
                    " and instr(ir04,'確收')=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            GRD1.TextMatrix(i, 3) = "*" & GRD1.TextMatrix(i, 3)
         End If
         '解析收受者:
         '轉寄日期和轉寄時間相同
         If GRD1.TextMatrix(i, 17) = GRD1.TextMatrix(i, 19) And GRD1.TextMatrix(i, 18) = GRD1.TextMatrix(i, 20) Then
            GRD1.TextMatrix(i, 6) = IIf(Trim(GRD1.TextMatrix(i, 26)) = "Y", "[副]", "") & PUB_ReadUserData(GRD1.TextMatrix(i, 14))
         Else
            strSql = "SELECT ir04,ir24 FROM inputrecord" & _
                     " WHERE ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and ir11=" & GRD1.TextMatrix(i, 19) & _
                       " and ir12=" & GRD1.TextMatrix(i, 20)
            intI = 1
            strUser = ""
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               With RsTemp
                  RsTemp.MoveFirst
                  Do While RsTemp.EOF = False
                     strUser = strUser & ";" & IIf(Trim("" & RsTemp.Fields("ir24")) = "Y", "[副]", "") & PUB_ReadUserData(RsTemp.Fields("ir04"))
                     RsTemp.MoveNext
                  Loop
               End With
            End If
            GRD1.TextMatrix(i, 6) = Mid(strUser, 2)
         End If
         'Add By Sindy 2025/1/8 檢查是否要清空處理狀態
         If GRD1.TextMatrix(i, 7) <> "" Then
            If PUB_CheckIRStatus(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), _
                             ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), "", 1) = True Then
               GRD1.TextMatrix(i, 7) = ""
            End If
         End If
         '2025/1/8 END
      Next i
      Call SetColor
   Else
      If bolShowMsg = True Then
         Screen.MousePointer = vbDefault
         ShowNoData
      End If
   End If
   rsTmp.Close
   
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   GRD1.Visible = True
   dblPrevRow = 0
   
   Screen.MousePointer = vbDefault
   
   'Add By Sindy 2022/9/5 為他人職代時,檢查是否有待處理信件,若有要彈訊息
   If bolShowMsg = True Then '一般信件資料區才檢查
      Call GetWorkEmpListData
   End If
   
   Set rsTmp = Nothing
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   MsgBox "查詢失敗！" & vbCrLf & Err.Description & vbCrLf & vbCrLf & strSql
End Function

'記錄查詢
Private Sub cmdHistory_Click()
Dim nFrm As Form
   
   '檢查表單是否已開啟，若是，則關閉
   For Each nFrm In Forms
      If StrComp(nFrm.Name, "frm06010613", vbTextCompare) = 0 Then
         Unload frm06010613
      End If
   Next
   
   Call frm06010613.SetParent(Me)
   frm06010613.m_WorkType = 1 '轉寄 Add By Sindy 2017/12/12
   frm06010613.m_MailUsernum = Me.Combo1.Text
'   frm06010613.cboII06 = IIf(Me.Combo1.Text = "不處理信件", "", Me.Combo1.Text)
'   frm06010613.cboII06.Tag = frm06010613.cboII06.Text
'   frm06010613.cboIR13 = IIf(Me.Combo1.Text = "不處理信件", "", Me.Combo1.Text)
'   frm06010613.cboIR13.Tag = frm06010613.cboIR13.Text
'   frm06010613.txtDate(2) = strSrvDate(2)
'   frm06010613.txtDate(3) = strSrvDate(2)
'   frm06010613.Check1.Visible = False '含刪除未轉寄資料
'   frm06010613.Caption = "信件記錄查詢 - 信件收受者"
'   Call frm06010613.SetFrame1
   frm06010613.Show
   Me.Hide
End Sub

Public Sub GoNext()
Dim strIR16 As String
Dim strIR22 As String

   With GRD1
      If Val(m_iRow) = Val(txtPI18.Tag) Then
'         '上刪除標記,高度設零
'         .row = m_iRow
         If PUB_CheckIRStatus(m_strIR01, m_strIR02, m_strIR03, m_strIR04) = True Then
            If m_strIR01 = .TextMatrix(m_iRow, 10) And _
               m_strIR03 = .TextMatrix(m_iRow, 16) And _
               m_strIR04 = .TextMatrix(m_iRow, 13) Then
               LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
               Call SaveInputRecord(m_iRow)
               Call CancelRowColor(m_iRow) '清除反白
               GRD1.RowHeight(m_iRow) = 0
               
'               '檢查是否有輸入動作且需主管核准的通知信要發
'               strExc(0) = "select * from inputrecord" & _
'                           " where ir01=" & .TextMatrix(m_iRow, 10) & _
'                             " and ir03='" & .TextMatrix(m_iRow, 16) & "'" & _
'                             " and ir04='" & .TextMatrix(m_iRow, 13) & "'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strIR16 = "" & RsTemp.Fields("IR16")
'                  strIR22 = "" & RsTemp.Fields("IR22")
'               End If
'               'Modify By Sindy 2022/8/19
'               'If strIR16 = "1" And strIR22 <> "" Then
'               If strIR22 <> "" Then
'               '2022/8/19 END
'                  strExc(0) = "select cum02 from CaseUseMemo" & _
'                              " where cum05='02'" & _
'                                " and cum06=" & CNULL(strUserNum) & _
'                                " and cum02='" & strIR22 & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 0 Then
'                     strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
'                                 " values('0','" & strIR22 & "','0','0','02')"
'                     cnnConnection.Execute strExc(0)
'                     Frame1.Visible = True '*****
'                  End If
'               End If
               
               If Val(Replace(LblTotCnt.Caption, "總筆數:", "")) = 0 Then
                  Call QueryData(False)
               End If
            Else
               Call QueryData(False)
            End If
'            If .TextMatrix(.row, 0) = "V" Then
'            End If
'            .TextMatrix(.row, 0) = "X"
'            .RowHeight(.row) = 0
'            LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
'            Call ReadFirstGrd1Text '查詢勾選的第一筆資料
         End If
      Else
         Call QueryData(False)
      End If
   End With
End Sub

'Add By Sindy 2022/8/19
Private Sub cmdMainQuy_Click()
   '信件未處理查詢
   If CheckUse("frm100106_9", strExec) = True Then
      frm100106_9.m_WorkType = 2 '外專
      frm100106_9.Show
   End If
End Sub

'主管核准
Private Sub cmdMgOK_Click()
Dim bolHavdData As Boolean
Dim strUpdTime As String
Dim bolConn As Boolean
Dim intQ As Integer ', strIR20 As String
Dim bolOK As Boolean
Dim strIR16nm As String 'Added by Morgan 2023/5/23
   
   bolHavdData = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         'Modify By Sindy 2022/9/8 開放可以多筆點選資料列
'         If dblPrevRow <> i Then
'            MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
'            Exit Sub
'         End If
         bolHavdData = True
         Exit For
      End If
   Next i
   If bolHavdData = False Then
      MsgBox "請至少勾選一筆欲主管核准的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      '檢查是否有處理狀態,有,才能操作主管核准
      'Modified by Morgan 2023/5/23 +strIR16nm
      If PUB_CheckIRStatus(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), strIR16nm) = False Then
         MsgBox "此封信件未處理，不可操作核准！", vbExclamation, "警告！"
         Exit Sub
      End If
      
      'Added by Morgan 2023/4/12
      '有期限來函進2次確認畫面
      'Modified by Morgan 2023/5/23 +strIR16nm
      If IsReKeyInCase(GRD1.TextMatrix(i, 29), strIR16nm) Then
         With GRD1
         .row = i
         m_strIR01 = .TextMatrix(.row, 10)
         m_strIR02 = .TextMatrix(.row, 11)
         m_strIR03 = .TextMatrix(.row, 16)
         m_strIR04 = .TextMatrix(.row, 13)
         End With
         Call Forms(0).SetTmpfrm04010519(Me)
         frm02010605.m_CP09 = GRD1.TextMatrix(i, 24)
         frm02010605.m_strIR01 = m_strIR01
         frm02010605.m_strIR02 = m_strIR02
         frm02010605.m_strIR03 = m_strIR03
         frm02010605.m_strIR04 = m_strIR04
         frm02010605.Show
         Exit Sub
      Else
      'end 2023/4/12
      
         intQ = MsgBox("確定是否核准？" & vbCrLf & vbCrLf & _
                   "核准：請按【是】" & vbCrLf & _
                   "不核准,需加註原因：請按【否】" & vbCrLf & _
                   "取消此動作請按【取消】", vbExclamation + vbYesNoCancel + vbDefaultButton1, "重要訊息！")
         If intQ = vbCancel Then
            Exit Sub
         End If
         If intQ = vbNo Then
            bolOK = False '3.退回
   '         strIR20 = InputBox("請輸入退回原因")
   '         If strIR20 = "" Then Exit Sub
            If Trim(txtIR20) = "" Or Trim(txtIR20) = "其他," Or Trim(txtIR20) = "其他，" Then
               MsgBox "請輸入退回原因！", vbExclamation, "警告！"
               txtIR20.SetFocus
               Exit Sub
            End If
         Else
            bolOK = True '主管核准
         End If
         
      End If 'Added by Morgan 2023/4/12
   End If
   
On Error GoTo ErrHand
   
   '檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      'Modify By Sindy 2023/3/7 要排除回信的郵件,主管不可核准,直到Backup沖銷
      If GRD1.TextMatrix(i, 0) = "V" And GRD1.TextMatrix(i, 23) <> "9" Then '9.回信
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         
         If bolOK = False Then '3.退回
            strExc(0) = "update InputRecord set " & _
                        " ir16='3',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                        ",IR20=decode(IR20,null,'','" & strSrvDate(2) & "-" & strUserName & IIf(bolOK = False, "退回", "") & ":')||'" & ChgSQL(txtIR20) & "'||decode(IR20,null,'',';'||chr(13)||chr(10)||IR20)" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0)
            
            'Add By Sindy 2023/7/14 清除處理結果
            strExc(0) = "update IPDeptInput set ii27=null" & _
                        " where Ii01=" & GRD1.TextMatrix(i, 10) & _
                        " and Ii02=" & GRD1.TextMatrix(i, 11) & _
                        " and Ii03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "' and Ii16=0"
            cnnConnection.Execute strExc(0)
            '2023/7/14 END
            
            strExc(0) = "select cum02 from CaseUseMemo" & _
                        " where cum05='02'" & _
                          " and cum06=" & CNULL(strUserNum) & _
                          " and cum02=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                           " values('0',upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "'),'0','0','02')"
               cnnConnection.Execute strExc(0)
               Frame1.Visible = True '*****
            End If
            
         Else
            '7.已確認(主管核准)
            strExc(0) = "update InputRecord set " & _
                        "ir16='7',ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0), intI
            
            '未處理的都自沖(上確認日期時間人員)
            'Modify By Sindy 2023/4/28 + 增加同部門判斷 and exists(select st01 from staff where st01=ir04 and st03='" & PUB_GetST03(Trim(Left(Combo1, 6))) & "')
            strExc(0) = "update InputRecord set " & _
                        "ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and exists(select st01 from staff where st01=ir04 and st03='" & PUB_GetST03(Trim(Left(Combo1, 6))) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0), intI
            
            Call SaveInputRecord(i, False)
         End If
         
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "信件主管核准註記失敗！" & vbCrLf & Err.Description
End Sub

'不處理
Private Sub cmdNotProDel_Click()
Dim intHavdData As Integer
Dim strUpdTime As String
Dim bolConn As Boolean
Dim bolUptCaseNo As Boolean
Dim strIR16 As String
   
   intHavdData = 0: bolUptCaseNo = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
            Exit Sub
         End If
         
         '檢查是否有處理狀態
         If PUB_CheckIRStatus(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), _
                          ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), strIR16) = True Then
            MsgBox "此封信件已有人操作【" & strIR16 & "】，請畫面更新！", vbExclamation, "警告！"
            Exit Sub
         End If
         
         intHavdData = i
         Exit For
      End If
   Next i
   If intHavdData = 0 Then
      MsgBox "請至少勾選一筆不處理的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If Trim(txtIR20) = "" Or Trim(txtIR20) = "其他," Or Trim(txtIR20) = "其他，" Then
         MsgBox "不處理原因不可空白！", vbExclamation, "警告！"
         txtIR20.SetFocus
         Exit Sub
      End If
      
'      If Pub_StrUserSt03 = "F23" Then '承辦組
         If txtPI18 <> "" Then '有輸入案號代表要歸卷，所以要檢查案號和文號資料
            'Modify By Sindy 2024/7/15 + 因會歸卷詢問系統類別的權限
            If ChkCaseNo(, True) = False Then
            '2024/7/15 END
               Exit Sub
            Else
               'Modify By Sindy 2024/7/15
               If txtPI18.Enabled = True Then
               '2024/7/15 END
                  bolUptCaseNo = True
               End If
            End If
            If PUB_ChkFileOpening2(m_strFullFileName, "後續才能一併歸卷！") = True Then
               Exit Sub
            End If
         End If
'      Else
'         If txtPI18 <> "" Then
'            '檢查卷宗區是否已有此信件,若有,無需重覆歸卷
'            strExc(10) = "." & GRD1.TextMatrix(intHavdData, 10) & Format(GRD1.TextMatrix(intHavdData, 11), "0#####") & "." & GRD1.TextMatrix(intHavdData, 16) & "."
'            strExc(0) = "select count(*) from casepaperpdf,caseprogress" & _
'                        " where instr(cpp02,'" & strExc(10) & "')>0" & _
'                          " and cpp01=cp09" & _
'                          " and cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "'"
'            intI = 1
'            strExc(10) = ""
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If RsTemp.Fields(0) = 0 Then
'                  If MsgBox("此信件未歸卷，要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
'                     Exit Sub
'                  End If
'               End If
'            End If
'         End If
'      End If
      
      If MsgBox("確定信件不處理嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   '檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         '不處理
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
                  
         strExc(0) = "update InputRecord set " & _
                     " ir16='2',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                     ",ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & IIf(bolUptCaseNo = True, ",ir21='" & txtRecvNo.Text & "'", "")
         If txtIR20 <> m_TxtIR20 Then
            strExc(0) = strExc(0) & _
                        ",IR20=decode(IR20,null,'','" & strSrvDate(2) & "-" & strUserName & ":')||'" & ChgSQL(txtIR20) & "'||decode(IR20,null,'',';'||chr(13)||chr(10)||IR20)"
         End If
         strExc(0) = strExc(0) & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                       " and ir08=0"
         cnnConnection.Execute strExc(0)
         
         '更新案號
         If bolUptCaseNo = True Then
            strExc(0) = "update IPDeptInput set " & _
                        "Ii23='" & txtPI18 & "',Ii24='" & txtPI19 & "'," & _
                        "Ii25='" & txtPI20 & "',Ii26='" & txtPI21 & "'" & _
                        " where Ii01=" & GRD1.TextMatrix(i, 10) & _
                        " and Ii02=" & GRD1.TextMatrix(i, 11) & _
                        " and Ii03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'"
            cnnConnection.Execute strExc(0)
         End If
         
         If PUB_IPDeptEMailF2UptRec(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), _
            ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), "", "2", _
            strUpdTime, IIf(bolUptCaseNo = True, txtRecvNo.Text, "")) = False Then
            
            GoTo ErrHand 'Add By Sindy 2023/2/24
         End If
         
         Call SaveInputRecord(i)
         
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
         Exit For 'Modify By Sindy 2022/8/26
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "信件不處理註記失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   bolQuery = True
   PubShowNextData
   Exit Sub
End Sub

'檢查本所案號和總收文號
'bolChkRecv: 檢查收文號
'Modify By Sindy 2024/7/15 + , Optional bolPro11Ask As Boolean = False: 因會歸卷詢問系統類別的權限
Public Function ChkCaseNo(Optional bolChkRecv As Boolean = True, Optional bolPro11Ask As Boolean = False) As Boolean
Dim strCP12 As String

   ChkCaseNo = False
   Screen.MousePointer = vbHourglass 'Add By Sindy 2025/6/19
   
   If txtPI18 <> "FCP" And txtPI18 <> "FG" And txtPI18 <> "P" Then
      'Add By Sindy 2024/7/15
      If bolPro11Ask = True Then
         If MsgBox("系統類別非屬本部門，" & vbCrLf & "確定「不須歸卷至本部門的類別(FCP、FG、P)」嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            txtPI18.Enabled = False
            txtPI19.Enabled = False
            txtPI20.Enabled = False
            txtPI21.Enabled = False
            txtRecvNo.Enabled = False
            Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
            ChkCaseNo = True
            Exit Function
         End If
      End If
      '2024/7/15 END
      Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
      MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
      Me.txtPI18.SetFocus
      Exit Function
   End If
   If txtPI20 = "" Then txtPI20 = "0"
   If txtPI21 = "" Then txtPI21 = "00"
   
   If txtRecvNo = "" And bolChkRecv = True Then
      Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
      MsgBox "請選擇要歸卷的總收文號！", vbExclamation, "警告！"
      Me.txtRecvNo.SetFocus
      Exit Function
'   ElseIf txtRecvNo <> "" Then
'      bolChkRecv = True
   End If
   
   strExc(0) = "select cp09,cp01,cp12 from caseprogress" & _
               " where cp01='" & txtPI18 & "'" & _
                 " and cp02='" & txtPI19 & "'" & _
                 " and cp03='" & txtPI20 & "'" & _
                 " and cp04='" & txtPI21 & "'"
   'Modify By Sindy 2025/6/9 改在下面做檢查
'   If bolChkRecv = True Then
'      strExc(0) = strExc(0) & " and cp09='" & txtRecvNo & "'"
'   End If
   '2025/6/9 END
   'Modify By Sindy 2022/10/28
   strExc(0) = strExc(0) & " order by cp66 desc,cp67 desc"
   '2022/10/28 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      'Modify By Sindy 2025/6/9 改在下面做檢查
'      If bolChkRecv = True Then
'         MsgBox "查無進度資料，請確認！", vbExclamation, "警告！"
'         Me.txtRecvNo.SetFocus
'         Exit Function
'      Else
      '2025/6/9 END
         Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
         MsgBox "無此案號，請重新輸入！", vbExclamation, "警告！"
         Me.txtPI19.SetFocus
         Exit Function
'      End If
   Else
      strCP12 = "" & RsTemp.Fields("cp12")
      '檢查權限
      If CheckSR09(strUserNum, txtPI18.Text, "Y", , txtPI18, txtPI19, txtPI20, txtPI21) = False Then
         Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
         txtPI19.SetFocus
         Exit Function
      End If
      
      'Modify By Sindy 2022/10/28 Mark; ex:FCP-065283
'      If Left(strCP12, 2) <> "F2" Then
'         MsgBox "無此案號權限！", vbExclamation, "警告！"
'         Me.txtPI19.SetFocus
'         Exit Function
'      End If
   End If
   
   'Modify By Sindy 2025/6/9
   If bolChkRecv = True Then
      strExc(0) = "select cp09,cp01,cp12 from caseprogress" & _
                  " where cp01='" & txtPI18 & "'" & _
                    " and cp02='" & txtPI19 & "'" & _
                    " and cp03='" & txtPI20 & "'" & _
                    " and cp04='" & txtPI21 & "'"
      strExc(0) = strExc(0) & " and cp09='" & txtRecvNo & "'"
      strExc(0) = strExc(0) & " order by cp66 desc,cp67 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
         MsgBox "此總收文號有誤(非屬此案號)，請確認！", vbExclamation, "警告！"
         Me.txtRecvNo.SetFocus
         Exit Function
      End If
   End If
   '2025/6/9 END
   
   Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
   ChkCaseNo = True
End Function

'歸卷
Private Sub cmdPDF_Click()
Dim strUpdTime As String
Dim bolConn As Boolean
Dim strIR21 As String
Dim stFileName As String, strCR02 As String, StrCR04 As String
Dim fs, f
Dim strMailDate As String, strMailTime As String
Dim bolChkExists As Boolean 'Add By Sindy 2023/8/24
   
On Error GoTo ErrHand
   
   If Frame5.Visible = True Then
      If Option1(0).Value = False And Option1(1).Value = False Then
         MsgBox "請選擇一種歸卷方式！", vbExclamation, "警告！"
         Exit Sub
      End If
   Else
      Option1(0).Value = True '只歸卷宗區
   End If
   
   With GRD1
      For m_iRow = 1 To .Rows - 1
         .row = m_iRow
         If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
            If dblPrevRow <> .row Then
               MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
               Exit Sub
            End If
            
'*******************************
            '檢查資料:
'*******************************
            '歸卷宗區
            If Option1(0).Value = True Then
               If txtPI18 = "" Or txtPI19 = "" Then
                  MsgBox "請輸入本所案號！", vbExclamation, "警告！"
                  If txtPI18 = "" Then
                     Me.txtPI18.SetFocus
                  ElseIf txtPI19 = "" Then
                     Me.txtPI19.SetFocus
                  End If
                  Exit Sub
               Else
                  If ChkCaseNo = False Then
                     Exit Sub
                  End If
               End If
               
            '歸往來記錄
            Else
               If Trim(txtCOR03) = "" Then
                  MsgBox "請輸入X,Y,R編號 並且選擇 往來記錄編號！", vbExclamation, "警告！"
                  Me.txtCOR03.SetFocus
                  Exit Sub
               ElseIf Trim(txtCOR01) = "" Then
                  MsgBox "請選擇 往來記錄編號！", vbExclamation, "警告！"
                  Me.txtCOR01.SetFocus
                  Exit Sub
               End If
            End If
            
            If MsgBox("確定信件要歸卷嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Exit Sub
            End If
            
'*******************************
            '儲存資料:
'*******************************
            Screen.MousePointer = vbHourglass
            cnnConnection.BeginTrans: bolConn = True
            strUpdTime = Right("000000" & ServerTime, 6)
            
            strMailDate = DBDATE(Left(GRD1.TextMatrix(.row, 3), 9))
            strMailTime = Format(Replace(Right(GRD1.TextMatrix(.row, 3), 8), ":", ""), "000000")
            '歸卷宗區
            If Option1(0).Value = True Then
               strExc(0) = "update IPDeptInput set " & _
                           "Ii23='" & txtPI18 & "',Ii24='" & txtPI19 & "'," & _
                           "Ii25='" & txtPI20 & "',Ii26='" & txtPI21 & "'" & _
                           " where Ii01=" & GRD1.TextMatrix(.row, 10) & _
                           " and Ii02=" & GRD1.TextMatrix(.row, 11) & _
                           " and Ii03='" & ChgSQL(GRD1.TextMatrix(.row, 16)) & "'"
               cnnConnection.Execute strExc(0)
               .TextMatrix(.row, 4) = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
               .TextMatrix(.row, 24) = txtRecvNo
               
               '下載信件檔,上傳卷宗區
               '敏莉說,程序組且系統別P,PS的歸卷,副檔名為代理人來函
               'Modify By Sindy 2023/8/24 + bolChkExists
               If PUB_UploadPatentLetterFile(GRD1.TextMatrix(.row, 10), GRD1.TextMatrix(.row, 16), GRD1.TextMatrix(.row, 24), _
                  IIf(Pub_StrUserSt03 = "F22" And (txtPI18 = "P" Or txtPI18 = "PS"), "ALTR", ""), , , , bolChkExists) = False Then
                  Screen.MousePointer = vbDefault
                  If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
                  Exit Sub
               End If
               'Add By Sindy 2023/8/24
               If bolChkExists = True Then
                  MsgBox "此信件已存在卷宗區！", , "已歸卷"
                  Screen.MousePointer = vbDefault
                  If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
                  Exit Sub
               End If
               '2023/8/24 END
               
               strIR21 = txtRecvNo '總收文號
               
            '歸往來記錄
            Else
               '純下載信件檔
               If PUB_UploadPatentLetterFile(GRD1.TextMatrix(.row, 10), GRD1.TextMatrix(.row, 16), _
                     GRD1.TextMatrix(.row, 24), , stFileName, True) = False Then
                  Screen.MousePointer = vbDefault
                  If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
                  Exit Sub
               End If
               
               strExc(0) = "select * from Contactfile where CF01='" & txtCOR01 & "'" & _
                           " and (instr(CF02,'" & txtCOR01 & "_" & GRD1.TextMatrix(.row, 10) & Format(GRD1.TextMatrix(.row, 11), "000000") & "')>0" & _
                                " or instr(CF02,'" & GRD1.TextMatrix(.row, 10) & Format(GRD1.TextMatrix(.row, 11), "000000") & "." & GRD1.TextMatrix(.row, 16) & "')>0" & _
                                " or instr(CF02,'" & txtCOR01 & "_" & strMailDate & strMailTime & "')>0" & _
                                " or instr(CF02,'" & strMailDate & strMailTime & "." & GRD1.TextMatrix(.row, 16) & "')>0)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Screen.MousePointer = vbDefault
                  MsgBox "附件已存在該往來記錄中！", vbExclamation, "警告！"
                  Me.txtCOR01.SetFocus
                  Exit Sub
               End If
               
               '直接存入電子檔
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  Screen.MousePointer = vbDefault
                  ShowMsg stFileName & MsgText(9221)
                  If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
                  Exit Sub
               End If
               
               strIR21 = txtCOR01 '往來記錄編號
               
               '儲存往來記錄附件檔
               If PUB_UpdateCFData(txtCOR01, stFileName, f.Size, "Rx") = False Then 'Rx:外來郵件
                  Screen.MousePointer = vbDefault
                  If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
                  Exit Sub
               Else
                  Call PUB_DelPCOrgFile(stFileName, , False) '一併將PC上的實體檔案刪除
               End If
            End If
            
            '不沖銷信件
            strExc(0) = "update InputRecord set ir21='" & strIR21 & "'" & _
                        ",ir16='4',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'"
            If txtIR20 <> m_TxtIR20 Then
               strExc(0) = strExc(0) & _
                           ",IR20=decode(IR20,null,'','" & strSrvDate(2) & "-" & strUserName & ":')||'" & ChgSQL(txtIR20) & "'||decode(IR20,null,'',';'||chr(13)||chr(10)||IR20)"
            End If
            strExc(0) = strExc(0) & _
                        " where ir01=" & GRD1.TextMatrix(.row, 10) & _
                          " and ir02=" & GRD1.TextMatrix(.row, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(.row, 16)) & "'" & _
                          " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(.row, 13)) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0)
            
            '11.歸卷
            strExc(0) = "update IPDeptInput set " & _
                        "Ii27='11'" & _
                        " where Ii01=" & GRD1.TextMatrix(.row, 10) & _
                        " and Ii02=" & GRD1.TextMatrix(.row, 11) & _
                        " and Ii03='" & ChgSQL(GRD1.TextMatrix(.row, 16)) & "'"
            cnnConnection.Execute strExc(0)
            
'            Call SaveInputRecord(.row, False)
            
            cnnConnection.CommitTrans: bolConn = False
            
'            LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
'            Call CancelRowColor(.row) '清除反白
'            GRD1.RowHeight(.row) = 0
            'Call ReadFirstGrd1Text '查詢勾選的第一筆資料
            Exit For
         End If
      Next m_iRow
   End With
   Call QueryData(False)
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
   MsgBox "信件歸卷失敗！" & vbCrLf & Err.Description
End Sub

'已處理
Private Sub cmdProDel_Click()
Dim bolHavdData As Boolean
Dim strUpdTime As String
Dim bolConn As Boolean
Dim bolUptCaseNo As Boolean
   
   bolHavdData = False: bolUptCaseNo = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
            Exit Sub
         End If
         bolHavdData = True
         Exit For
      End If
   Next i
   If bolHavdData = False Then
      MsgBox "請至少勾選一筆已處理的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      'Add By Sindy 2019/6/11
      If Trim(txtIR20) = "" Or Trim(txtIR20) = "其他," Or Trim(txtIR20) = "其他，" Then
         MsgBox "原因不可空白！", vbExclamation, "警告！"
         txtIR20.SetFocus
         Exit Sub
      End If
      '2019/6/11 END
      
      If txtPI18 <> "" Then '有輸入案號代表要歸卷，所以要檢查案號和文號資料
         'Modify By Sindy 2024/7/15 + 因會歸卷詢問系統類別的權限
         If ChkCaseNo(, True) = False Then
         '2024/7/15 END
            Exit Sub
         Else
            'Modify By Sindy 2024/7/15
            If txtPI18.Enabled = True Then
            '2024/7/15 END
               bolUptCaseNo = True
            End If
         End If
         If PUB_ChkFileOpening2(m_strFullFileName, "後續才能一併歸卷！") = True Then
            Exit Sub
         End If
      End If
      
      If MsgBox("確定信件已處理？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   'Add by Sindy 2021/11/19 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         '5.已處理
         strExc(0) = "update InputRecord set " & _
                     " ir16='5',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                     ",ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & IIf(bolUptCaseNo = True, ",ir21='" & txtRecvNo.Text & "'", "")
         If txtIR20 <> m_TxtIR20 Then
            strExc(0) = strExc(0) & _
                        ",IR20=decode(IR20,null,'','" & strSrvDate(2) & "-" & strUserName & ":')||'" & ChgSQL(txtIR20) & "'||decode(IR20,null,'',';'||chr(13)||chr(10)||IR20)"
         End If
         strExc(0) = strExc(0) & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                       " and ir08=0"
         cnnConnection.Execute strExc(0)
         
         '更新案號
         If bolUptCaseNo = True Then
            strExc(0) = "update IPDeptInput set " & _
                        "Ii23='" & txtPI18 & "',Ii24='" & txtPI19 & "'," & _
                        "Ii25='" & txtPI20 & "',Ii26='" & txtPI21 & "'" & _
                        " where Ii01=" & GRD1.TextMatrix(i, 10) & _
                        " and Ii02=" & GRD1.TextMatrix(i, 11) & _
                        " and Ii03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'"
            cnnConnection.Execute strExc(0)
         End If
         
         If PUB_IPDeptEMailF2UptRec(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), _
            ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), "", "5", _
            strUpdTime, IIf(bolUptCaseNo = True, txtRecvNo.Text, "")) = False Then
            
            GoTo ErrHand 'Add By Sindy 2023/2/24
         End If
         
         Call SaveInputRecord(i)
         
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
         Exit For 'Modify By Sindy 2022/8/26
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "信件已處理註記失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdQuery_Click()
   If Me.Combo1.Text = "" Then
      MsgBox "員工編號不可空白！", vbExclamation, "警告！"
      Combo1.SetFocus
      Exit Sub
   End If
   
   Call QueryData
End Sub

'回覆確收
Private Sub cmdRecall_Click()
Dim strFileName As String, strFullFileName As String
Dim strUpdTime As String
Dim bolConn As Boolean
Dim strCnt As String
Dim strIR20 As String
   
On Error GoTo ErrHand

   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If GRD1.TextMatrix(i, 2) = "Y" Then
            If MsgBox("要【重覆確收】嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Exit Sub
            End If
         End If
         '讀取檔案
         Screen.MousePointer = vbHourglass
         strFileName = "$$" & Mid(GRD1.TextMatrix(i, 15), InStrRev(GRD1.TextMatrix(i, 15), "/") + 1)
         If GetAttachFile(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), GRD1.TextMatrix(i, 16), strFullFileName, m_AttachPath & "\" & strFileName) = True Then
            ShellExecute 0, "open", strFullFileName, vbNullString, vbNullString, 1 '開檔
         
            cnnConnection.BeginTrans: bolConn = True
            strUpdTime = Right("000000" & ServerTime, 6)
            
            If txtIR20 <> m_TxtIR20 Then
               strIR20 = txtIR20
            End If
            
            '檢查目前此封郵件已確收幾次
            strExc(0) = "select count(ir01) from inputrecord" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and instr(ir04,'確收')>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            strCnt = ""
            If intI = 1 Then
               strCnt = CStr(RsTemp.Fields(0) + 1)
            End If
            '確收
            strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR14" & _
                        ",IR08,IR09,IR10,IR20)" & _
                        " values(" & GRD1.TextMatrix(i, 10) & _
                                 "," & GRD1.TextMatrix(i, 11) & _
                                 ",'" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                                 ",'確收" & strCnt & "'," & strSrvDate(1) & "," & _
                                 strUpdTime & ",'" & strUserNum & "'," & CNULL(IIf(Trim(Left(Combo1, 6)) <> strUserNum, Trim(Left(Combo1, 6)), "")) & _
                                 "," & strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "','" & ChgSQL(strIR20) & "')"
            cnnConnection.Execute strExc(0)
            
            cnnConnection.CommitTrans: bolConn = False
            GRD1.TextMatrix(i, 2) = "Y" '確收註記
            'Call CancelRowColor(i) '清除反白
         End If
         Exit For 'Modify By Sindy 2022/8/26
      End If
   Next i
   Screen.MousePointer = vbDefault
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox "確收失敗！" & vbCrLf & Err.Description
End Sub

'9.回信
Private Sub cmdReMail_Click()
Dim strFileName As String, strFullFileName As String
Dim strUpdTime As String
Dim bolConn As Boolean
Dim strCnt As String
Dim stPwd As String
Dim bolUptCaseNo As Boolean
Dim strIR16 As String

'Dim objOutLook As Object
'Dim objMail As Object
'Dim myForward As Object
'Dim jj As Integer
   
On Error GoTo ErrHand
   
   bolUptCaseNo = False
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         '檢查是否有處理狀態
         If PUB_CheckIRStatus(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), _
                          ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), strIR16) = True Then
            MsgBox "此封信件已有人操作【" & strIR16 & "】，請畫面更新！", vbExclamation, "警告！"
            Exit Sub
         End If
         
         '讀取檔案
         Screen.MousePointer = vbHourglass
         strFileName = "$$" & Mid(GRD1.TextMatrix(i, 15), InStrRev(GRD1.TextMatrix(i, 15), "/") + 1)
         If GetAttachFile(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), GRD1.TextMatrix(i, 16), strFullFileName, m_AttachPath & "\" & strFileName) = True Then
'            If Pub_StrUserSt03 = "F23" Then '承辦組
               If txtPI18 <> "" Then '有輸入案號代表要歸卷，所以要檢查案號和文號資料
                  'Modify By Sindy 2024/7/15 + 因會歸卷詢問系統類別的權限
                  If ChkCaseNo(, True) = False Then
                  '2024/7/15 END
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  Else
                     'Modify By Sindy 2024/7/15
                     If txtPI18.Enabled = True Then
                     '2024/7/15 END
                        bolUptCaseNo = True
                     End If
'                     strFullFileName = App.path & "\" & strUserNum & "\" & Mid(GRD1.TextMatrix(i, 15), InStrRev(GRD1.TextMatrix(i, 15), "/") + 1)
'                     If PUB_ChkFileOpening2(strFullFileName) = True Then
'                        Exit Sub
'                     End If
                  End If
               End If
'            Else
'               If txtPI18 <> "" Then
'                  '檢查卷宗區是否已有此信件
'                  strExc(10) = "." & GRD1.TextMatrix(i, 10) & Format(GRD1.TextMatrix(i, 11), "0#####") & "." & GRD1.TextMatrix(i, 16) & "."
'                  strExc(0) = "select count(*) from casepaperpdf,caseprogress" & _
'                              " where instr(cpp02,'" & strExc(10) & "')>0" & _
'                                " and cpp01=cp09" & _
'                                " and cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "'"
'                  intI = 1
'                  strExc(10) = ""
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If RsTemp.Fields(0) > 0 Then
'                        If MsgBox("此信件已歸卷，要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
'                           Screen.MousePointer = vbDefault
'                           Exit Sub
'                        End If
'                     End If
'                  End If
'               End If
'            End If
                       
            If bolUptCaseNo = True Then
               '下載信件檔,上傳卷宗區
               If PUB_UploadPatentLetterFile(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 16), txtRecvNo.Text) = False Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            
            ShellExecute 0, "open", strFullFileName, vbNullString, vbNullString, 1 '開檔
            DoEvents
            
            '亂數:[%亂數5碼]
RunPwd:
            Clipboard.Clear
            Randomize
            stPwd = Fix(99999 * Rnd)
            stPwd = IIf(Len(stPwd) > 5, Left(stPwd, 5), IIf(Len(stPwd) < 5, Left(stPwd & "#####", 5), stPwd))
            stPwd = "[%" & stPwd & "]"
            '檢查是否有待沖銷回信的亂碼重覆了
            strExc(0) = "select count(*) from IPDeptinput,InputRecord" & _
                        " where Ii28='" & stPwd & "'" & _
                          " and Ii01=Ir01 and Ii03=Ir03 and Ir08=0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) > 0 Then
                  GoTo RunPwd
               End If
            End If
            
            Screen.MousePointer = vbDefault
            
'            '啟動轉寄功能
'            If strFullFileName <> "" Then
'               Set objOutLook = CreateObject("Outlook.Application")
'               Set objMail = objOutLook.CreateItemFromTemplate(strFullFileName) 'oForm.txtPathIPDept.Text & "\" & oFile.Name
'
'               '*** 轉寄 *** 會用inbound名義寄出
'         '目前問題是內文加文字怕會有亂碼問題
'         '        寄件者程式無法自動帶
'               Set myForward = objMail.Forward '轉寄
'               'Set myForward = objMail.Reply '回覆
''               '移除原信的收件人及副本;密件副本不會留在msg中
''               For jj = myForward.Recipients.Count To 1 Step -1
''                  myForward.Recipients.Remove jj
''               Next jj
''               'myForward.Recipients.add "test"
''               '副本
''               myForward.cc = ""
'               '主旨增加,當個案且有案號時,顯示歸入那一個案號
'               myForward.Subject = "RE: " & myForward.Subject & " " & stPwd
'               'myForward.senderemailaddress = "ipdept@taie.com.tw"
'               'myForward.sentonbehalfofname = "ipdept"
'               myForward.Display
'               'myForward.Send
'               DoEvents
'
'               Set myForward = Nothing
'               Set objMail = Nothing
'               Set objOutLook = Nothing
'               '*** END
'            End If
            
            Clipboard.SetText stPwd '複製編號至剪貼簿
            'If MsgBox("亂數已複製【" & stPwd & "】，確定回信了嗎？" & vbCrLf & vbCrLf & "注意：先回信後，再按下「是」因為信件關閉，後續才能一併歸卷！", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
            If MsgBox("沖銷碼已複製【" & stPwd & "】，請回貼於主旨後方。" & vbCrLf & vbCrLf & _
                      "是：已使用此沖銷碼進行沖銷" & vbCrLf & vbCrLf & _
                      "否：取消此動作", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
               Exit Sub
'            Else
'               If PUB_ChkFileOpening2(m_strFullFileName, "後續才能一併歸卷！") = True Then
'                  Exit Sub
'               End If
            End If
            
            Screen.MousePointer = vbHourglass
            cnnConnection.BeginTrans: bolConn = True
            
            strUpdTime = Right("000000" & ServerTime, 6)
            
            strExc(0) = "update InputRecord set " & _
                        " ir16='9',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & IIf(bolUptCaseNo = True, ",ir21='" & txtRecvNo.Text & "'", "")
            If txtIR20 <> m_TxtIR20 Then
               strExc(0) = strExc(0) & _
                           ",IR20=decode(IR20,null,'','" & strSrvDate(2) & "-" & strUserName & ":')||'" & ChgSQL(txtIR20) & "'||decode(IR20,null,'',';'||chr(13)||chr(10)||IR20)"
            End If
            strExc(0) = strExc(0) & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0)
            
            '更新回信沖銷序號
            strExc(0) = "update IPDeptInput set Ii28='" & stPwd & "'" & _
                        " where Ii01=" & GRD1.TextMatrix(i, 10) & _
                        " and Ii02=" & GRD1.TextMatrix(i, 11) & _
                        " and Ii03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'"
            cnnConnection.Execute strExc(0)
            
            '更新案號
            If bolUptCaseNo = True Then
               strExc(0) = "update IPDeptInput set " & _
                           "Ii23='" & txtPI18 & "',Ii24='" & txtPI19 & "'," & _
                           "Ii25='" & txtPI20 & "',Ii26='" & txtPI21 & "'" & _
                           " where Ii01=" & GRD1.TextMatrix(i, 10) & _
                           " and Ii02=" & GRD1.TextMatrix(i, 11) & _
                           " and Ii03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'"
               cnnConnection.Execute strExc(0)
            End If
            
            If PUB_IPDeptEMailF2UptRec(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), _
               ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), "", "9", _
               strUpdTime) = False Then
               
               GoTo ErrHand 'Add By Sindy 2023/2/24
            End If
         
            cnnConnection.CommitTrans: bolConn = False
            GRD1.TextMatrix(i, 23) = "9" '回信
            GRD1.TextMatrix(i, 7) = strUserName & "回信"
'            LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
            Call CancelRowColor(i) '清除反白
'            GRD1.RowHeight(i) = 0
         End If
         Exit For 'Modify By Sindy 2022/8/26
      End If
   Next i
   Screen.MousePointer = vbDefault
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox "回信失敗！" & vbCrLf & Err.Description
End Sub

'選擇往來記錄
Private Sub cmdSelCont_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
Dim ii As Integer, jj As Integer
   
   If Trim(txtCOR03) <> "" Then
      Me.Tag = ""
      sqlB = "select '' AS V,CR01 AS 往來記錄編號," & SQLDate("CR02") & " 往來日期,CR06 主旨,CR08 內容,ST02 建檔人員," & SQLDate("CR13") & " 建檔日期" & _
             " from contactrecord,staff" & _
             " where SUBSTR(cr03,1,8)='" & Left(txtCOR03, 8) & "' and CR12=ST01(+)" & _
             " order by cr01 desc"
      intI = 0
      Set rsRead = ClsLawReadRstMsg(intI, sqlB)
      If intI = 1 Then
         Set frm880012.grdDataList.Recordset = rsRead
         Set frm880012.fmParent = Me
         frm880012.iTyp = "6"
         frm880012.Show vbModal
         If Me.Tag <> "" Then
            txtCOR01.Text = Me.Tag
            txtCOR01.SetFocus
         Else
            txtCOR01.Text = ""
         End If
      End If
   Else
      MsgBox "請先輸入X,Y,R編號！", vbExclamation, "警告！"
      If Me.txtCOR03.Enabled = True Then Me.txtCOR03.SetFocus
   End If
End Sub

'選擇總收文號
Private Sub cmdSelCp09_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
Dim ii As Integer, jj As Integer
    
   If Trim(txtPI18) <> "" And Trim(txtPI19) <> "" Then
      Me.Tag = ""
      txtPI20.Text = IIf(txtPI20 = "", "0", txtPI20)
      txtPI21.Text = IIf(txtPI21 = "", "00", txtPI21)
      
      sqlB = "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(pa09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
             "from caseprogress,casepropertymap,staff s1,staff s2,patent " & _
             "where cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "' " & _
             "and cp01=cpm01(+) and cp10=cpm02(+) " & _
             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
             "and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 "
'      'Add By Sindy 2019/4/30 + 商標
'      sqlB = sqlB & " union " & _
'             "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(tm10,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
'             "from caseprogress,casepropertymap,staff s1,staff s2,trademark " & _
'             "where cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "' " & _
'             "and cp01=cpm01(+) and cp10=cpm02(+) " & _
'             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
'             "and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 "
      'Add By Sindy 2019/4/30 + 服務
      sqlB = sqlB & " union " & _
             "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(sp09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
             "from caseprogress,casepropertymap,staff s1,staff s2,servicepractice " & _
             "where cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "' " & _
             "and cp01=cpm01(+) and cp10=cpm02(+) " & _
             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
             "and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 "
      sqlB = sqlB & " ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"
      intI = 0
      Set rsRead = ClsLawReadRstMsg(intI, sqlB)
      If intI = 1 Then
         '檢查權限
         If CheckSR09(strUserNum, txtPI18.Text, "Y", , txtPI18, txtPI19, txtPI20, txtPI21) = False Then
            txtPI19.SetFocus
            Exit Sub
         End If
         Set frm880012.grdDataList.Recordset = rsRead
         Set frm880012.fmParent = Me
         frm880012.iTyp = "1"
         If txtRecvNo.Text <> "" Then
            For ii = 1 To frm880012.grdDataList.Rows - 1
               If frm880012.grdDataList.TextMatrix(ii, 2) = txtRecvNo.Text Then
                  frm880012.grdDataList.col = 0
                  frm880012.grdDataList.row = ii
                  frm880012.grdDataList.Text = "V": frm880012.m_iSelRow = ii
                  For jj = 0 To frm880012.grdDataList.Cols - 1
                     frm880012.grdDataList.col = jj
                     frm880012.grdDataList.CellBackColor = &HFFC0C0
                  Next jj
                  Exit For
               End If
            Next ii
         End If
         frm880012.Show vbModal
         If Me.Tag <> "" Then
            txtRecvNo.Text = Me.Tag
            If txtRecvNo <> "" Then
               strExc(0) = "select * from caseprogress" & _
                           " where cp09='" & txtRecvNo & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Label5.Caption = GetPrjState4(txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21, "" & RsTemp.Fields("cp10"))
               End If
            End If
            txtRecvNo.SetFocus
         Else
            txtRecvNo.Text = ""
            Label5.Caption = ""
         End If
      End If
   Else
      MsgBox "請先輸入本所案號！", vbExclamation, "警告！"
      If Me.txtPI18.Enabled = True Then Me.txtPI18.SetFocus
   End If
End Sub

'立即寄發通知信
Private Sub cmdSendMail_Click()
Dim strContent As String
Dim bolCaseDutyAgentMsg As Boolean, strRestKind As String
Dim strTempCC As String
Dim strRestEmp As String, strNormalEmp As String
Dim ArrStr As Variant, jj As Integer
   
   strExc(0) = "select cum02 from CaseUseMemo" & _
               " where cum05='02'" & _
                 " and cum06=" & CNULL(strUserNum)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         '因為有休假問題,所以有休假人員各自發信,其他人則一封
         strTempCC = GetCaseDutyAgent(RsTemp.Fields("cum02"), "", bolCaseDutyAgentMsg, strRestKind)
         If strTempCC <> "" Then
            strRestEmp = strRestEmp & ";" & RsTemp.Fields("cum02")
         Else
            strNormalEmp = strNormalEmp & ";" & RsTemp.Fields("cum02")
         End If
         RsTemp.MoveNext
      Loop
      strContent = "請至案件管理系統的一般作業\系統收件區，進行看查。"
      '薛經理:一起通知,減少操作人員等發信的時間
      If strNormalEmp <> "" Then
         strNormalEmp = Mid(strNormalEmp, 2)
         PUB_SendMail strUserNum, strNormalEmp, "", "通知已有信件轉入系統收件區", strContent, , , , , , , , , , , False, , , , , , , , , , , , , "1"
      End If
      '有休假人員各自發信
      If strRestEmp <> "" Then
         strRestEmp = Mid(strRestEmp, 2)
         ArrStr = Split(strRestEmp, ";")
         For jj = 0 To UBound(ArrStr)
            PUB_SendMail strUserNum, ArrStr(jj), "", "通知已有信件轉入系統收件區", strContent, , , , , , , , , , , False, , , , , , , , , , , , , "1"
         Next jj
      End If
      '刪除記錄
      strExc(0) = "delete from CaseUseMemo" & _
                  " where cum05='02'" & _
                  " and cum06=" & CNULL(strUserNum)
      cnnConnection.Execute strExc(0)
   End If
   Frame1.Visible = False '*****
End Sub

'轉寄鍵
Private Sub cmdUpdRow_Click()
Dim bolHavdSel As Boolean
Dim strIR16 As String
Dim intRunRow As Integer
   
   '收受者不可空白
   If List1.ListCount <= 0 Then
      MsgBox "收受者不可空白！", vbExclamation, "警告！"
      cboII06.SetFocus
      Exit Sub
   End If
   
   bolHavdSel = False
   '先檢查是否有資料要刪除
   If GRD1.Rows - 1 < 1 Then Exit Sub
   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
            Exit Sub
         End If
         
         '檢查是否有處理狀態
         If PUB_CheckIRStatus(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), _
                          ChgSQL(GRD1.TextMatrix(i, 16)), ChgSQL(GRD1.TextMatrix(i, 13)), strIR16) = True Then
            MsgBox "此封信件已有人操作【" & strIR16 & "】，請畫面更新！", vbExclamation, "警告！"
            Exit Sub
         End If
         
         bolHavdSel = True
         Exit For
      End If
   Next i
   If bolHavdSel = False Then
      MsgBox "請至少勾選一筆要轉寄的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If MsgBox("確定要轉寄信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
   'Modify By Sindy 2023/7/14
   If PUB_IPDeptTransMail(Trim(Left(Combo1, 6)), GRD1, Check1, List1, Frame1, m_AttachPath, _
      IIf(TextContext.Enabled = True And Trim(TextContext) <> "", TextContext, ""), intRunRow) = True Then
      Call SaveInputRecord(intRunRow, False)
      
      If Left(GRD1.TextMatrix(intRunRow, 3), 1) <> "*" Then
         GRD1.TextMatrix(intRunRow, 3) = "*" & GRD1.TextMatrix(intRunRow, 3)
      End If
      
      If Check1.Value = 1 Then '上刪除日期
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(intRunRow) '清除反白
         GRD1.RowHeight(intRunRow) = 0
      End If
      '清除收受者
      cboII06.Text = ""
      cboCC.Text = "" '副本
   End If
   '2023/7/14 END
End Sub

'待歸檔
Private Sub cmdWait_Click()
Dim bolHavdData As Boolean
Dim strUpdTime As String
Dim bolConn As Boolean
   Exit Sub '注意,防往下執行
   bolHavdData = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
            Exit Sub
         End If
         bolHavdData = True
         Exit For
      End If
   Next i
   If bolHavdData = False Then
      MsgBox "請至少勾選一筆待歸檔的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If MsgBox("確定信件待歸檔嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   'Add by Sindy 2021/11/19 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         
         strExc(0) = "update InputRecord set " & _
                     " ir16='6',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                     ",ir22='" & Trim(Left(Combo1, 6)) & "'"
         If txtIR20 <> m_TxtIR20 Then
            strExc(0) = strExc(0) & _
                        ",IR20=decode(IR20,null,'','" & strSrvDate(2) & "-" & strUserName & ":')||'" & ChgSQL(txtIR20) & "'||decode(IR20,null,'',';'||chr(13)||chr(10)||IR20)"
         End If
         strExc(0) = strExc(0) & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                       " and ir08=0"
         cnnConnection.Execute strExc(0)
         
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "信件待歸檔註記失敗！" & vbCrLf & Err.Description
End Sub

Private Sub Combo1_Click()
   If Combo1.Text <> "" Then
      Call QueryData '(False)
   End If
End Sub

'加註符號
Private Sub Combo2_Click()
   If dblPrevRow > 0 Then
      If GRD1.TextMatrix(dblPrevRow, 0) = "V" Then
         strExc(0) = "update InputRecord set ir23=" & CNULL(Left(Combo2.Text, 1)) & _
                     " where ir01=" & GRD1.TextMatrix(dblPrevRow, 10) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 13)) & "')"
         cnnConnection.Execute strExc(0)
         GRD1.TextMatrix(dblPrevRow, 1) = Left(Combo2.Text, 1)
         Call SetColor(dblPrevRow) 'Add By Sindy 2022/8/25
      End If
   End If
End Sub

Private Sub Form_Activate()
   '外專承辦F23和程序人員F22請使用「國外部專利及承辦人系統」操作系統收件區
   If Pub_StrUserSt03 = "F23" Then
      If UCase(Left(App.EXEName, 7)) <> "PATPRO1" And UCase(Left(App.EXEName, 9)) <> "TEPATPRO1" Then
         MsgBox "外專承辦請使用「國外部專利及承辦人系統」" & vbCrLf & vbCrLf & "操作系統收件區！", vbExclamation
         Unload Me
         Exit Sub
      End If
   ElseIf Pub_StrUserSt03 = "F22" Then
      If UCase(Left(App.EXEName, 7)) <> "PATPRO1" And UCase(Left(App.EXEName, 9)) <> "TEPATPRO1" And _
         UCase(Left(App.EXEName, 6)) <> "PATPRO" And UCase(Left(App.EXEName, 8)) <> "TEPATPRO" Then
         MsgBox "外專程序請使用「國外部專利及承辦人系統」或「專利及承辦人系統」" & vbCrLf & vbCrLf & "操作系統收件區！", vbExclamation
         Unload Me
         Exit Sub
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ClearText
   
   'If Me.Combo1.Text = "" Then Me.Combo1.Text = strUserNum & " " & strUserName
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   'Add By Sindy 2022/8/19
   cmdMainQuy.Visible = False
   If PUB_GetST05(strUserNum) = "31" Then '外專程序主管才顯示出來
      cmdMainQuy.Visible = True
   End If
   '2022/8/19 END
   
   '檢查是否有未寄送通知信
   strExc(0) = "select cum02 from CaseUseMemo" & _
               " where cum05='02'" & _
                 " and cum06=" & CNULL(strUserNum) & _
                 " and rownum<=1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Frame1.Visible = True '*****
   Else
      Frame1.Visible = False
   End If
   
   '備註/原因
   cboReason.Clear
   If Pub_StrUserSt03 = "F22" Then '程序組
      cboReason.AddItem "不處理:"
      cboReason.AddItem "  ACK信函(含自動回覆)"
      cboReason.AddItem "  重覆信件(含同一案件/事項多封來函)"
      cboReason.AddItem "  寰華回覆問題"
      cboReason.AddItem "  寰華通知期限 , 但期限已管制"
      cboReason.AddItem "  寰華提醒答覆期限"
      cboReason.AddItem "  寰華簡單報告提交日"
      cboReason.AddItem "  其他"
      cboReason.AddItem "主管退回:"
      cboReason.AddItem "  期限錯誤"
      cboReason.AddItem "  資料錯誤"
      cboReason.AddItem "  其他"
      cboReason.AddItem "已處理:"
      cboReason.AddItem "  已回覆處理"
      cboReason.AddItem "  已轉工程師處理"
      cboReason.AddItem "  已轉承辦處理"
      cboReason.AddItem "  已轉財務處理"
      cboReason.AddItem "  已修改期限/資料" 'Add By Sindy 2022/8/12
      cboReason.AddItem "  其他"
   Else
      cboReason.AddItem "不處理:"
      cboReason.AddItem "  ACK信函(含自動回覆)"
      cboReason.AddItem "  重覆信件(含同一案件/事項多封來函)"
      cboReason.AddItem "  已處理"
      cboReason.AddItem "  處理中"
      cboReason.AddItem "  已收文"
      cboReason.AddItem "  已轉工程師處理(含會議及非個案)"
      cboReason.AddItem "  已轉財務處理"
      cboReason.AddItem "  已轉寄他部門"
      cboReason.AddItem "  例行通知"
      cboReason.AddItem "  廣告開拓信函"
      cboReason.AddItem "  其他"
   End If
   
   '收受者
   cboII06.Clear
   cboII06.AddItem ""
   cboII06.AddItem "外商群組 國外部轉信外商群組" '洪琬姿 ,葉易雲,沈佳穎,陳蒲璇
'國外部轉信外法日文組群組: 桂齊恆 , 江郁仁, 顏裕洋, 葉易雲, 陳毓芳
'國外部轉信外法群組: 桂齊恆 , 陳亮之, 顏裕洋, 洪琬姿, 葉易雲, 江郁仁, 陳毓芳, 楊映慈, 潘子微, 林美宏, 沈佳穎
'國外部轉信外法英文組群組: 桂齊恆 , 陳亮之, 顏裕洋, 洪琬姿, 潘子微, 楊映慈, 林美宏, 沈佳穎, 葉易雲
   cboII06.AddItem "新知群組 國外部轉信新知群組" '閻?泰,EXTERNAL_NEWS@taie.com.tw,顏裕洋,鄒宜珊  ? 妳們轉寄給這群組, David也是一員, 但你會因若人員上了不處理 且 主管核准了,信件就沖銷了哦~~ 以上也是如此…
   cboII06.AddItem "開拓群組 國外部轉信開拓群組" '閻?泰,楊雯芳,陳增廣,鄒宜珊
   cboII06.AddItem "Patent Patent@taie.com.tw"
   cboII06.AddItem "TM TM@taie.com.tw"
   'LblTM.Caption = LblTM.Caption & strTMCppFilePath
   'LblTM.Visible = False
   'F2外專人員,若有其他人輸員工編號按Tab鍵即可
   strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and substr(st03,1,2)='F2' order by st03,st01 asc"
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
   
   '副本
   cboCC.Clear
   cboCC.AddItem ""
   'F2外專人員,若有其他人輸員工編號按Tab鍵即可
   strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and substr(st03,1,2)='F2' order by st03,st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            cboCC.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   
   SetCmdStatus
'   '設定同意/退回的位置
'   Frame3.Left = 7140
'   Frame3.Top = 2070
   
   FrameCont.Left = 5430
   FrameCont.Top = 1710
   
   Call SetCombo1
   'Call QueryData(False)
   
   'Added by Sindy 2021/11/22 如果一開始將ListBox拉到需要的大小，字型會自動放大；
   '所以畫面預設為一列高度(315)，Form_Load才放大到需要的大小
   List1.Clear
   List1.Height = 750 '目前此高度是可以看到完整最後一筆的
   List1.Width = 1875
End Sub

Private Sub SetCmdStatus()
   Frame5.Visible = False
   cmdDelRow.Enabled = False '刪除
   cmdProDel.Enabled = False '已處理
   '外專承辦F23
   cmdReMail.Caption = "回信"
   If Pub_StrUserSt03 = "F23" Then
      Option1(0).Value = True
      cmdReMail.Caption = "承辦作業"
      cmdProDel.Enabled = False
      Frame5.Visible = True
   ElseIf Pub_StrUserSt03 = "F22" Then '程序組
      cmdInput.Enabled = True
      cmdProDel.Enabled = True '已處理
      'cmdMgOK.Enabled = True
   '工程師
   Else
      cmdInput.Visible = False
      cmdPDF.Enabled = False
      FrameRecv.Enabled = False: cmdSelCp09.Enabled = False
      'cmdMgOK.Visible = False
      cmdDelRow.Enabled = True '刪除
   End If
   cmdDelRow.Tag = cmdDelRow.Enabled 'Add By Sindy 2021/1/22 記錄原狀況
   
'   '檢查是否有主管核准的權限
'   cmdMgOK.Enabled = False
'   strSql = "SELECT count(*) FROM Staff WHERE (st52='" & strUserNum & "' or st53='" & strUserNum & "' or st54='" & strUserNum & "' or st55='" & strUserNum & "') AND ST04='1'" & _
'            " order by 1 asc"
'   intI = 0
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      If RsTemp.Fields(0) > 0 Then
'         cmdMgOK.Enabled = True
'      End If
'   End If
End Sub

'Add By Sindy 2022/9/5 為他人職代時,檢查是否有待處理信件,若有要彈訊息
Private Function GetWorkEmpListData(Optional ByRef strMsg As String, Optional ByVal bolShowMsg As Boolean = True) As Boolean
Dim strEmp As String
   
   GetWorkEmpListData = False
   If Trim(m_WorkEmpList) = "" Then Exit Function
   If Right(m_WorkEmpList, 1) = ";" Then
      m_WorkEmpList = Mid(m_WorkEmpList, 1, Len(m_WorkEmpList) - 1)
   End If
   strEmp = Replace(m_WorkEmpList, ";", "','")
   'Modify By Sindy 2023/6/15 + and (ir22 is null or ir16 in('3','8'))
   strSql = "SELECT ir04 FROM InputRecord WHERE IR04 in('" & strEmp & "')" & _
            " and ir08=0 and (ir22 is null or ir16 in('3','8'))" & _
            " group by ir04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strEmp = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If InStr(Combo1.Text, RsTemp.Fields("ir04")) = 0 Then
            GetWorkEmpListData = True
            strEmp = strEmp & "、" & GetPrjSalesNM(RsTemp.Fields("ir04"))
         End If
         RsTemp.MoveNext
      Loop
      If strEmp <> "" Then
         strEmp = Mid(strEmp, 2)
         strMsg = strEmp & "請假，請至其收件區處理郵件"
         If bolShowMsg = True Then
            MsgBox strMsg, vbInformation, "[代理通知] 尚有郵件未處理"
         End If
      End If
   End If
End Function

Private Sub SetCombo1()
Dim strText As String, ii As Integer
Dim Rs As New ADODB.Recordset
Dim strTemp As String
   
   strTemp = Combo1.Text 'Modify By Sindy 2025/9/2 記錄欄位值
   
   Combo1.Clear
   Combo1.AddItem strUserNum & " " & strUserName
   '檢查當時是否需要為他人職代
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False, m_WorkEmpList)
   
   '外專承辦
   If Pub_StrUserSt03 = "F23" Then
      '檢查是否為主管，是的話可以幫主管代為操作
      strSql = "SELECT st01 FROM Staff WHERE (st52='" & strUserNum & "' or st53='" & strUserNum & "' or st54='" & strUserNum & "' or st55='" & strUserNum & "') AND ST04='1'" & _
               " order by 1 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modify By Sindy 2025/8/6 Anny反應"待核准信件"不應該出現組員
         If Check2.Value = 0 Then
         '2025/8/6 END
            'Add By Sindy 2024/5/2 日專承辦主管(陳毓芳)要看到她底下全部人員
            If Pub_StrUserSt93 = "J21" Then '日專承辦
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If Not IsNull(RsTemp.Fields(0)) Then
                     strText = Trim(RsTemp.Fields(0)) & " " & GetPrjSalesNM(Trim(RsTemp.Fields(0)))
                     For ii = 0 To Combo1.ListCount - 1
                        If InStr(Combo1.List(ii), Trim(RsTemp.Fields(0))) = 1 Then
                           Exit For
                        End If
                     Next
                     If ii = Combo1.ListCount Then
                        Combo1.AddItem strText
                     End If
                  End If
                  RsTemp.MoveNext
               Loop
            End If
            '2024/5/2 END
         End If
         
         '主管可以幫2級主管操作
         'Modify By Sindy 2025/8/6 改抓 AND st93='" & PUB_GetST93(strUserNum) & "'
         strSql = "SELECT distinct st52 FROM Staff WHERE st52 is not null AND st93='" & PUB_GetST93(strUserNum) & "' AND ST04='1' order by 1 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               .MoveFirst
               Do While Not .EOF
                  If Not IsNull(RsTemp.Fields(0)) Then
                     strText = Trim(RsTemp.Fields(0)) & " " & GetPrjSalesNM(Trim(RsTemp.Fields(0)))
                     For ii = 0 To Combo1.ListCount - 1
                        If InStr(Combo1.List(ii), Trim(RsTemp.Fields(0))) = 1 Then
                           Exit For
                        End If
                     Next
                     If ii = Combo1.ListCount Then
                        Combo1.AddItem strText
                     End If
                  End If
                  .MoveNext
               Loop
            End With
         End If
      End If
      '2級主管可以幫組員操作
      'Modify By Sindy 2025/8/6 Anny反應"待核准信件"不應該出現組員
      If Check2.Value = 0 Then
      '2025/8/6 END
         strSql = "SELECT distinct st01 FROM Staff WHERE st52='" & strUserNum & "' AND ST04='1' order by 1 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               .MoveFirst
               Do While Not .EOF
                  If Not IsNull(RsTemp.Fields(0)) Then
                     strText = Trim(RsTemp.Fields(0)) & " " & GetPrjSalesNM(Trim(RsTemp.Fields(0)))
                     For ii = 0 To Combo1.ListCount - 1
                        If InStr(Combo1.List(ii), Trim(RsTemp.Fields(0))) = 1 Then
                           Exit For
                        End If
                     Next
                     If ii = Combo1.ListCount Then
                        Combo1.AddItem strText
                     End If
                  End If
                  .MoveNext
               Loop
            End With
         End If
      End If
      
   '外專程序
   ElseIf Pub_StrUserSt03 = "F22" Then
      'Modify By Sindy 2022/8/10 寰華案外專程序窗口可以看到其他程序資料
      If InStr(Pub_GetSpecMan("寰華案外專程序窗口"), strUserNum) > 0 Then
         strSql = "SELECT distinct st01 FROM Staff WHERE st03='F22' AND ST04='1' order by 1 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               .MoveFirst
               Do While Not .EOF
                  If Not IsNull(RsTemp.Fields(0)) Then
                     strText = Trim(RsTemp.Fields(0)) & " " & GetPrjSalesNM(Trim(RsTemp.Fields(0)))
                     For ii = 0 To Combo1.ListCount - 1
                        If InStr(Combo1.List(ii), Trim(RsTemp.Fields(0))) = 1 Then
                           Exit For
                        End If
                     Next
                     If ii = Combo1.ListCount Then
                        Combo1.AddItem strText
                     End If
                  End If
                  .MoveNext
               Loop
            End With
         End If
      End If
   End If
   
   'Add By Sindy 2023/5/31
   If Left(Pub_StrUserSt03, 2) <> "F2" Then
      MsgBox "注意：此沖銷信件是屬外專承辦及程序人員操作，一定要用該單位人員ID，因程式判斷上是用單位做相關的處理！", vbCritical
      Combo1.Clear
   Else
      'Modify By Sindy 2025/9/2 還原欄位值
      'Combo1.Text = Combo1.List(0)
      If strTemp <> "" Then
         Combo1.Text = strTemp
      Else
      '2025/9/2 END
         Combo1.Text = Combo1.List(0)
      End If
   End If
   '2023/5/31 END
   
   Set Rs = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Modify By Sindy 2018/6/20 Run執行檔.exe才發E-Mail; 或測試時要執行
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Or _
      UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
   '2018/6/20 END
      '立即寄送通知信
      If Frame1.Visible = True Then Call cmdSendMail_Click
   End If
   
   DestroyToolTip '清除物件
   Set frm06010616 = Nothing
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   '                        0    1     2     3               4           5       6         7               8                     9               10      11      12      13      14      15      16      17      18      19      20      21      22      23      24          25          26      27      28          29
   arrGridHeadText = Array("V", "符", "確", "收信日期時間", "本所案號", "主旨", "收受者", "處理或轉寄者", "處理或轉寄日期時間", "讀取日期時間", "IR01", "IR02", "Ii18", "IR04", "Ii06", "Ii14", "檔名", "Ii08", "Ii09", "ir11", "ir12", "Ii12", "Ii05", "IR16", "總收文號", "處理原因", "ir24", "ir19", "信箱來源", "ir21")
   arrGridHeadWidth = Array(200, 250, 250, 1300, 1200, 2800, 1500, 1100, 1200, 1200, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 2000, 0, 0, 0, 0)
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
   If GRD1.RowHeight(1) = 0 Then GRD1.RowHeight(1) = 255 '***
   GRD1.Visible = True
End Sub

'Add By Sindy 2020/8/26
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Frame5.ToolTipText = "歸卷方式，是給歸卷銨鈕使用。"
End Sub

'Add By Sindy 2020/8/26
Private Sub cmdSelCont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmdSelCont.ToolTipText = "要按【歸卷】時才需要輸入。"
End Sub

'Add By Sindy 2020/8/26
Private Sub cmdSelCp09_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmdSelCp09.ToolTipText = "要按【輸入】或【歸卷】時才需要輸入。"
End Sub

Private Sub Grd1_Click()
Dim strIR16 As String
Dim ii As Integer
Dim bolManyRow As Boolean

GRD1.Visible = False
GRD1.row = GRD1.MouseRow
GRD1.col = GRD1.MouseCol
nRow = GRD1.row
nCol = GRD1.col
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
   'Add By Sindy 2022/9/13 全部清除反白
   If Check2.Value = 0 Then
      For i = 1 To GRD1.Rows - 1
         GRD1.col = 0
         GRD1.row = i
         If GRD1.CellBackColor = &HFFC0C0 Then
            Call CancelRowColor(i) '清除反白
            dblPrevRow = 0
            Exit For
         End If
      Next i
   End If
   '2022/9/13 END
Else
   cmdCanclReMail.Visible = False 'Add By Sindy 2024/4/24 預設值
   
   'Modify By Sindy 2022/9/8 待核准信件,開放可以多筆點選資料列
   'Modified by Morgan 2023/4/12 程序取消多筆點選
   If Check2.Value = 1 And Pub_StrUserSt03 <> "F22" Then
      If GRD1.TextMatrix(nRow, 0) = "V" Then
         GRD1.TextMatrix(nRow, 0) = ""
         Call CancelRowColor(CInt(nRow)) '清除反白
         dblPrevRow = 0
         GRD1.Visible = True
         Exit Sub
      End If
   Else
   '2022/9/8 END
      'Modify By Sindy 2017/12/22
      If dblPrevRow <= 0 Then
         dblPrevRow = 0
      Else
         'Modify By Sindy 2022/8/26 Mark,導至多筆勾選
   '      'Modify By Sindy 2017/12/29 記錄的目前資料列是未選取狀況,尋找目前反白的資料列,清除反白
   '      GRD1.col = 3
   '      GRD1.row = dblPrevRow
   '      If GRD1.CellBackColor <> &HFFC0C0 Then
   '         For i = 1 To GRD1.Rows - 1
   '            GRD1.row = i
   '            If GRD1.CellBackColor = &HFFC0C0 Then
   '               Call CancelRowColor(GRD1.row) '清除反白
   '               dblPrevRow = 0
   '               Exit For
   '            End If
   '         Next i
   '      '2017/12/29 END
   '      Else
         '2022/8/26 END
         
         If dblPrevRow <> nRow Then
            GRD1.TextMatrix(dblPrevRow, 0) = ""
            Call CancelRowColor(CInt(dblPrevRow)) '清除反白
         End If
      End If
   End If
   '2017/12/22 END
   GRD1.row = nRow 'GRD1.MouseRow
   dblPrevRow = GRD1.row '記錄目前筆數
   GRD1.col = 0
   
   'Modify By Sindy 2022/8/26 Mark,外專不適用
'   'Add By Sindy 2021/1/22 收受者前面加[副],已無電子檔時,才要亮刪除按鍵
'   If GRD1.TextMatrix(GRD1.row, 6) <> "" And Trim(GRD1.TextMatrix(GRD1.row, 15)) = "" Then
'      If Left(GRD1.TextMatrix(GRD1.row, 6), 3) = "[副]" Then '副本目前均為主管,查看的信件,是從別的信箱轉入的信件
'         cmdDelRow.Enabled = True
'      Else
'         cmdDelRow.Enabled = cmdDelRow.Tag
'      End If
'   End If
   
   If GRD1.TextMatrix(GRD1.row, 16) <> "" Then
      '將點選資料列反白
      GRD1.TextMatrix(GRD1.row, 0) = "V"
      
      If GRD1.TextMatrix(GRD1.row, 23) = "9" Then '9.回信
         Frame2.Enabled = False
         'Add By Sindy 2024/4/24
         If Check2.Value = 1 And GRD1.TextMatrix(GRD1.row, 0) = "V" Then  '待核准信件資料區
            bolManyRow = False
            For ii = 1 To GRD1.Rows - 1
               If GRD1.TextMatrix(ii, 0) = "V" Then
                  If ii <> dblPrevRow Then
                     bolManyRow = True '有多筆勾選資料列
                     Exit For
                  End If
               End If
            Next ii
            If bolManyRow = False Then
               cmdCanclReMail.Visible = True
            End If
         End If
         '2024/4/24 END
      Else
         Frame2.Enabled = True
      End If
'      '清除反白
'      'If GRD1.TextMatrix(GRD1.row, 0) = "V" Then
'      If grd1.CellBackColor = &HFFC0C0 Then
'         Call CancelRowColor(grd1.row) '清除反白
'         If txtPI18.Tag <> "" And (Val(grd1.row) <= Val(txtPI18.Tag)) Then Call ReadFirstGrd1Text '查詢勾選的第一筆資料
'      Else
''         If Trim(GRD1.TextMatrix(GRD1.row, 9)) = "" And nCol = 0 Then '無讀取日期不可以操作其功能
''            GRD1.Visible = True
''            MsgBox "此郵件尚未讀取(開啟)不可操作其功能，因此不可勾選!!", vbExclamation, "警告！"
''         Else
            
            '勾選資料列的檔名(在此處組實體檔名路徑,只是為了方便不用在每支Form再組一次檔名,因要判斷是否檔案開著)
            m_strFullFileName = App.path & "\" & strUserNum & "\" & Mid(GRD1.TextMatrix(GRD1.row, 15), InStrRev(GRD1.TextMatrix(GRD1.row, 15), "/") + 1)
            GRD1.col = 0
            GRD1.row = nRow
            For i = 0 To GRD1.Cols - 1
               If i <> 1 And i <> 7 Then
                  GRD1.col = i
                  GRD1.CellBackColor = &HFFC0C0
               End If
            Next i
            GRD1.Visible = True
            If List1.ListCount > 0 And CheckDataValid() = False Then
               Call CancelRowColor(GRD1.row) '清除反白
            End If
            'If txtPI18.Tag = "" Or (Val(grd1.row) <> Val(txtPI18.Tag)) Then Call ReadFirstGrd1Text '查詢勾選的第一筆資料
            Call ReadFirstGrd1Text '查詢勾選的資料
            
            '檢查是否有處理狀態
            If Check2.Value = 0 Then
               If PUB_CheckIRStatus(GRD1.TextMatrix(GRD1.row, 10), GRD1.TextMatrix(GRD1.row, 11), _
                                ChgSQL(GRD1.TextMatrix(GRD1.row, 16)), ChgSQL(GRD1.TextMatrix(GRD1.row, 13)), strIR16) = True Then
                  MsgBox "此封信件已有人操作【" & strIR16 & "】，請畫面更新！", vbExclamation, "警告！"
                  Exit Sub
               End If
            End If
''         End If
'      End If
   End If
End If
GRD1.Visible = True
End Sub

'Modify By Sindy 2017/12/27
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 And _
      (GRD1.MouseCol = 3 Or GRD1.MouseCol = 5 Or GRD1.MouseCol = 6 Or GRD1.MouseCol = 7 Or GRD1.MouseCol = 8) Then
      
      If iRow <> GRD1.MouseRow Or iCol <> GRD1.MouseCol Then
         If GRD1.MouseCol = 5 Or GRD1.MouseCol = 6 Or GRD1.MouseCol = 7 Or GRD1.MouseCol = 8 Then
            If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
               'GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
               CreateToolTip GetHWndForToolTip(GRD1), GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
               iRow = GRD1.MouseRow
               iCol = GRD1.MouseCol
            End If
         ElseIf GRD1.MouseCol = 3 Then
            'GRD1.ToolTipText = "信件編號:" & GRD1.TextMatrix(GRD1.MouseRow, 10) & "-" & GRD1.TextMatrix(GRD1.MouseRow, 16)
            CreateToolTip GetHWndForToolTip(GRD1), "信件編號:" & GRD1.TextMatrix(GRD1.MouseRow, 10) & "-" & GRD1.TextMatrix(GRD1.MouseRow, 16)
            iRow = GRD1.MouseRow
            iCol = GRD1.MouseCol
         End If
      End If
   End If
End Sub

'開啟附件
Private Sub GRD1_DblClick()
Dim strFileName As String, strFullFileName As String
Dim strUpdTime As String
Dim bolConn As Boolean
Dim strOpenFileName As String
   
On Error GoTo ErrHand
   
   GRD1.row = GRD1.MouseRow
   GRD1.col = GRD1.MouseCol
   nRow = GRD1.row
   nCol = GRD1.col
   If GRD1.col = 5 And nRow > 0 Then
      If GRD1.TextMatrix(dblPrevRow, 16) <> "" And dblPrevRow > 0 Then
         '讀取檔案
         Screen.MousePointer = vbHourglass
         strFileName = Mid(GRD1.TextMatrix(dblPrevRow, 15), InStrRev(GRD1.TextMatrix(dblPrevRow, 15), "/") + 1)
         Call PUB_ChkFileTypeOpenExE(strFileName) 'Add By Sindy 2017/9/13
         If GetAttachFile(GRD1.TextMatrix(dblPrevRow, 10), GRD1.TextMatrix(dblPrevRow, 11), GRD1.TextMatrix(dblPrevRow, 16), strFullFileName, m_AttachPath & "\" & strFileName) = True Then
            'Add By Sindy 2022/8/22 另外複製一個電子檔給使用者查看內容,因為有時msg檔關閉但OutLook資源沒有完整釋放,還是會判斷檔案是開啟的,無法歸檔
            strOpenFileName = m_AttachPath & "\$$" & strFileName
            If Dir(strOpenFileName) <> "" Then
               SetAttr strOpenFileName, vbNormal 'Add By Sindy 2020/1/17 檔案設定為正常屬性
               Kill strOpenFileName
            End If
            FileCopy strFullFileName, strOpenFileName
            DoEvents
            '2022/8/22 END
            'ShellExecute 0, "open", strFullFileName, vbNullString, vbNullString, 1 '開檔
            ShellExecute 0, "open", strOpenFileName, vbNullString, vbNullString, 1 '開檔
            
            '非電腦中心 (或電腦中心人員操作時,若收受者是自己信件時,可以更新讀取日期時間), 才需更新資料
            If Pub_StrUserSt03 <> "M51" Or _
               (Pub_StrUserSt03 = "M51" And Trim(Left(Combo1, 6)) = strUserNum) Then
               If Trim(GRD1.TextMatrix(dblPrevRow, 9)) = "" Then '無讀取日期時間才需更新資料
                  cnnConnection.BeginTrans: bolConn = True
                  strUpdTime = Right("000000" & ServerTime, 6)
                  strExc(0) = "update InputRecord set " & _
                              " ir05=" & strSrvDate(1) & ",ir06=" & strUpdTime & ",ir07='" & strUserNum & "'" & _
                              " where ir01=" & GRD1.TextMatrix(dblPrevRow, 10) & _
                                " and ir02=" & GRD1.TextMatrix(dblPrevRow, 11) & _
                                " and ir03='" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 16)) & "'" & _
                                " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 13)) & "')"
                  cnnConnection.Execute strExc(0)
                  
                  'Modify By Sindy 2022/8/26 Mark,外專不適用
'                  'Add By Sindy 2019/7/17 副本人員只要有讀取信件就上核銷(刪除資訊)
'                  strExc(0) = "update InputRecord set " & _
'                              " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
'                              " where ir01=" & GRD1.TextMatrix(dblPrevRow, 10) & _
'                                " and ir02=" & GRD1.TextMatrix(dblPrevRow, 11) & _
'                                " and ir03='" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 16)) & "'" & _
'                                " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 13)) & "')" & _
'                                " and ir24='Y'"
'                  cnnConnection.Execute strExc(0)
'                  Call SaveInputRecord(CInt(dblPrevRow), False)
'                  '2019/7/17 END

                  cnnConnection.CommitTrans: bolConn = False
                  
                  GRD1.TextMatrix(dblPrevRow, 9) = ChangeWStringToTDateString(strSrvDate(1)) & " " & Format(strUpdTime, "00:00:00")
               End If
            End If
            'Modify By Sindy 2017/12/22 Mark
'            If Trim(GRD1.TextMatrix(dblPrevRow, 9)) = "" Then '無讀取日期時間代表開啟沒成功
'               Call CancelRowColor(CInt(dblPrevRow)) '清除反白
'            End If
'         Else
'            MsgBox "無此郵件！", vbInformation
         End If
         Screen.MousePointer = vbDefault
      End If
   End If
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans
   If Err.Number = 70 Then
      MsgBox ChgSQL(strOpenFileName) & " 檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Function GetAttachFile(ByVal strPkey1 As String, ByVal strPkey2 As String, _
                               ByVal strPkey3 As String, ByRef pFileName As String, _
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
   
   GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
   
   Exit Function
   
ErrHnd:
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub SaveInputRecord(intRow As Integer, Optional bolSendMail As Boolean = True)
Dim strIR22 As String 'Add By Sindy 2022/8/19
   
   If Trim(GRD1.TextMatrix(intRow, 8)) <> "" Then '已有轉寄資料才須執行下列核銷
      '若信件收受者全部已處理或已刪除,主檔就可以掛上msg檔刪除日期,等待AutoBatchDay一個月後刪除實體檔
      strExc(0) = "select ir01 from InputRecord" & _
                  " where ir01=" & GRD1.TextMatrix(intRow, 10) & _
                    " and ir02=" & GRD1.TextMatrix(intRow, 11) & _
                    " and ir03='" & GRD1.TextMatrix(intRow, 16) & "'" & _
                    " and ir08=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then '信件收受者全部已處理或已刪除
         strExc(0) = "update IPDeptInput set" & _
                     " ii16=" & strSrvDate(1) & _
                     " where Ii01=" & GRD1.TextMatrix(intRow, 10) & _
                       " and Ii02=" & GRD1.TextMatrix(intRow, 11) & _
                       " and Ii03='" & GRD1.TextMatrix(intRow, 16) & "'" & _
                       " and ii16=0"
         cnnConnection.Execute strExc(0)
      End If
      
      'Modify By Sindy 2022/8/19
      If bolSendMail = True Then
         '檢查是否有需主管核准的通知信要發
         strExc(0) = "select * from inputrecord" & _
                     " where ir01=" & GRD1.TextMatrix(intRow, 10) & _
                       " and ir03='" & GRD1.TextMatrix(intRow, 16) & "'" & _
                       " and ir04='" & GRD1.TextMatrix(intRow, 13) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'strIR16 = "" & RsTemp.Fields("IR16")
            strIR22 = "" & RsTemp.Fields("IR22")
         End If
         If strIR22 <> "" Then
            strExc(0) = "select cum02 from CaseUseMemo" & _
                        " where cum05='02'" & _
                          " and cum06=" & CNULL(strUserNum) & _
                          " and cum02='" & strIR22 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                           " values('0','" & strIR22 & "','0','0','02')"
               cnnConnection.Execute strExc(0)
               Frame1.Visible = True '*****
            End If
         End If
      End If
      '2022/8/19 END
   End If
End Sub

Private Sub SetColor(Optional intSetRow As Double = 0)
   Dim ii As Integer, jj As Integer
   
   With GRD1
   If .Rows > 1 Then
      .Visible = False
      For ii = IIf(intSetRow = 0, 1, intSetRow) To IIf(intSetRow = 0, .Rows - 1, intSetRow)
         '若有讀取日期時間時,則變灰色
         .row = ii
         If Trim(.TextMatrix(ii, 9)) <> "" Then
            For jj = 2 To .Cols - 1
               If jj <> 7 Then
                  .col = jj
                  .CellBackColor = &HE0E0E0 '灰
               End If
            Next jj
         End If
         If Trim(.TextMatrix(ii, 23)) = "9" Then
            .col = 7
            .CellBackColor = &HC0FFFF '淺黃色
         ElseIf Trim(.TextMatrix(ii, 23)) = "3" Or Trim(.TextMatrix(ii, 23)) = "8" Then
            .col = 7
            .CellBackColor = &HC0C0FF '淺紅色
         End If
         '標示顏色註記
         If Trim(.TextMatrix(ii, 1)) = "R" Then
            .col = 1
            .CellBackColor = QBColor(12) '淡紅色
         ElseIf Trim(.TextMatrix(ii, 1)) = "Y" Then
            .col = 1
            .CellBackColor = QBColor(14) '淡黃色
         ElseIf Trim(.TextMatrix(ii, 1)) = "G" Then
            .col = 1
            .CellBackColor = QBColor(10) '淡綠色
         ElseIf Trim(.TextMatrix(ii, 1)) = "F" Then
            .col = 1
            .CellBackColor = QBColor(13) '紫紅色
         ElseIf Trim(.TextMatrix(ii, 1)) = "B" Then
            .col = 1
            .CellBackColor = QBColor(11) '淡藍色
         End If
      Next ii
      If intSetRow = 0 Then .TopRow = 1
      .Visible = True
   End If
   End With
End Sub

'點二下可刪除List1資料列
Private Sub List1_DblClick(Cancel As MSForms.ReturnBoolean)
Dim strText As String
   
   If List1.ListIndex >= 0 Then
      strText = List1.List(List1.ListIndex)
      List1.RemoveItem List1.ListIndex
      If List1.Tag <> "" Then
         List1.Tag = Replace(List1.Tag, ";" & strText, "")
         List1.Tag = Replace(List1.Tag, strText, "")
      End If
   End If
End Sub

'轉寄收受者
'Private Sub cboII06_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   cboII06.ToolTipText = "按Tab或點二下加入資料至右邊資料區"
'End Sub
'Private Sub cboII06_DblClick(Cancel As MSForms.ReturnBoolean)
'   If cboII06.ListIndex >= 0 Then
'      If InStr(List1.Tag, cboII06.List(cboII06.ListIndex)) = 0 Then
'         If List1.Tag = "" Then List1.Clear
'         If CheckDataValid(cboII06.List(cboII06.ListIndex)) = False Then GRD1.Visible = True: Exit Sub
'         List1.AddItem cboII06.List(cboII06.ListIndex)
'         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.List(cboII06.ListIndex)
'         List1.SetFocus
'      End If
'      cboII06.ListIndex = 0
'   End If
'End Sub
Private Sub cboII06_Click()
   If cboII06.ListIndex >= 0 Then
      If InStr(List1.Tag, cboII06.List(cboII06.ListIndex)) = 0 Then
         If List1.Tag = "" Then List1.Clear
         If CheckDataValid(cboII06.List(cboII06.ListIndex)) = False Then GRD1.Visible = True: Exit Sub
         List1.AddItem cboII06.List(cboII06.ListIndex)
         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.List(cboII06.ListIndex)
         List1.SetFocus
      End If
      cboII06.ListIndex = 0
   End If
End Sub
Private Sub cboII06_Validate(Cancel As Boolean)
   Call cboII06_LostFocus '收受者輸至第五碼時自動帶出姓名
   If cboII06.Text <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(cboII06, 5)) = True Then
         'cboII06.Text = ""
         cboII06.SetFocus
         Call cboII06_GotFocus
         Exit Sub
      Else
         If InStr(List1.Tag, cboII06.Text) = 0 Then
            If List1.Tag = "" Then List1.Clear
            If CheckDataValid(cboII06.Text) = False Then GRD1.Visible = True: Exit Sub
            List1.AddItem cboII06.Text
            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.Text
         End If
         cboII06.Text = ""
      End If
   End If
End Sub
Private Sub cboII06_GotFocus()
   cboII06.SelStart = 0
   cboII06.SelLength = Len(cboII06.Text)
End Sub
Private Sub cboII06_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cboII06_LostFocus()
Dim strText As String
   
   cboII06.Text = Trim(cboII06.Text) 'Add By Sindy 2021/11/22
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
'轉寄收受者
Private Sub Command3_Click()
Dim ArrStr As Variant
Dim ii As Integer, jj As Integer
Dim bolChk As Boolean, strData As String
   
   If List1.Tag = "" Then List1.Clear
   Call frm880024.SetParent(Me)
   frm880024.Caption = "轉寄收受者"
   frm880024.cboDepName_Click
   frm880024.Show vbModal
   If m_LstEmp <> "" Then
      ArrStr = Split(m_LstEmp, ";")
      For ii = 0 To UBound(ArrStr)
         bolChk = False
         strData = ArrStr(ii)
         If InStr(Combo1.Text, strData) > 0 Then bolChk = True
         For jj = 0 To List1.ListCount - 1
            If InStr(List1.List(jj), strData) > 0 Then
               bolChk = True
               Exit For
            End If
         Next jj
         If bolChk = False Then
            If CheckDataValid(strData) = False Then GRD1.Visible = True: Exit Sub
            'Add By Sindy 2023/3/31
            If GetPrjSalesNM(strData) = "" And InStr(Trim(strData), " ") = 0 Then
               List1.AddItem Replace(strData, "國外部轉信", "") & " " & strData
            Else
            '2023/3/31 END
               List1.AddItem strData & " " & GetPrjSalesNM(strData)
            End If
            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & strData & " " & GetPrjSalesNM(strData)
         End If
      Next ii
   End If
End Sub

'轉寄副本
Private Sub cboCC_Click()
   If cboCC.ListIndex >= 0 Then
      If InStr(List1.Tag, cboCC.List(cboCC.ListIndex)) = 0 Then
         If List1.Tag = "" Then List1.Clear
         If CheckDataValid(cboCC.List(cboCC.ListIndex)) = False Then GRD1.Visible = True: Exit Sub
         List1.AddItem cboCC.List(cboCC.ListIndex) & " (cc)"
         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboCC.List(cboCC.ListIndex) & " (cc)"
         List1.SetFocus
      End If
      cboCC.ListIndex = 0
   End If
End Sub
Private Sub cboCC_Validate(Cancel As Boolean)
   Call cboCC_LostFocus '收受者輸至第五碼時自動帶出姓名
   If cboCC.Text <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(cboCC, 5)) = True Then
         'cboCC.Text = ""
         cboCC.SetFocus
         Call cboCC_GotFocus
         Exit Sub
      Else
         If InStr(List1.Tag, cboCC.Text) = 0 Then
            If List1.Tag = "" Then List1.Clear
            If CheckDataValid(cboCC.Text) = False Then GRD1.Visible = True: Exit Sub
            List1.AddItem cboCC.Text & " (cc)"
            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboCC.Text & " (cc)"
         End If
         cboCC.Text = ""
      End If
   End If
End Sub
Private Sub cboCC_GotFocus()
   cboCC.SelStart = 0
   cboCC.SelLength = Len(cboCC.Text)
End Sub
Private Sub cboCC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cboCC_LostFocus()
Dim strText As String
   
   cboCC.Text = Trim(cboCC.Text) 'Add By Sindy 2021/11/22
   If cboCC.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboCC.Text)
      If strText <> "" Then
         cboCC.Text = strText & " " & cboCC.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboCC.Text, 5))
         If strText <> "" Then
            cboCC.Text = Left(cboCC.Text, 5) & " " & strText
         End If
      End If
   End If
End Sub
'轉寄副本
Private Sub Command4_Click()
Dim ArrStr As Variant
Dim ii As Integer, jj As Integer
Dim bolChk As Boolean, strData As String
   
   If List1.Tag = "" Then List1.Clear
   Call frm880024.SetParent(Me)
   frm880024.Caption = "轉寄副本"
   frm880024.cboDepName_Click
   frm880024.Show vbModal
   If m_LstEmp <> "" Then
      ArrStr = Split(m_LstEmp, ";")
      For ii = 0 To UBound(ArrStr)
         bolChk = False
         strData = ArrStr(ii)
         If InStr(Combo1.Text, strData) > 0 Then bolChk = True
         For jj = 0 To List1.ListCount - 1
            If InStr(List1.List(jj), strData) > 0 Then
               bolChk = True
               Exit For
            End If
         Next jj
         If bolChk = False Then
            If CheckDataValid(strData) = False Then GRD1.Visible = True: Exit Sub
            List1.AddItem strData & " " & GetPrjSalesNM(strData) & " (cc)"
            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & strData & " " & GetPrjSalesNM(strData) & " (cc)"
         End If
      Next ii
   End If
End Sub

Private Sub ClearText()
   'm_TxtIR20 = "可輸入處理原因"
   m_TxtIR20 = ""
   txtIR20 = m_TxtIR20
   txtPI18 = "": txtPI18.Tag = ""
   txtPI19 = ""
   txtPI20 = ""
   txtPI21 = ""
   txtRecvNo = "": Me.Tag = ""
   Label5.Caption = "" '案件性質名稱
   TextContext = vbCrLf & "信件內容參附件！"
   lbl1 = "": txtCOR01 = ""
   'Add By Sindy 2024/7/15 恢復狀態
   txtPI18.Enabled = True
   txtPI19.Enabled = True
   txtPI20.Enabled = True
   txtPI21.Enabled = True
   txtRecvNo.Enabled = True
   '2024/7/15 END
End Sub

'查詢勾選的資料
Private Sub ReadFirstGrd1Text()
   Call ClearText
   With GRD1
      For m_iRow = 1 To .Rows - 1
         .row = m_iRow
         If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
            txtIR20_show.Text = Trim(.TextMatrix(.row, 25)) '處理原因
            
            Option1(0).Value = True
            '將本所案號顯示在畫面上
            'If Trim(.TextMatrix(.row, 4)) <> "" Then
               'txtPI18.Tag = m_iRow
               txtPI18 = SystemNumber(Trim(.TextMatrix(.row, 4)), 1)
               txtPI19 = SystemNumber(Trim(.TextMatrix(.row, 4)), 2)
               txtPI20 = SystemNumber(Trim(.TextMatrix(.row, 4)), 3)
               txtPI21 = SystemNumber(Trim(.TextMatrix(.row, 4)), 4)
               txtRecvNo = Trim(.TextMatrix(.row, 24))
               If txtRecvNo <> "" Then
                  strExc(0) = "select * from caseprogress" & _
                              " where cp09='" & txtRecvNo & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     Label5.Caption = GetPrjState4(txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21, "" & RsTemp.Fields("cp10"))
                  End If
               End If
               
               txtIR20 = Trim(.TextMatrix(.row, 25))
               txtIR20 = m_TxtIR20
'               If txtIR20 = "" And Check2.Value = 0 Then
'                  txtIR20 = m_TxtIR20
'               ElseIf Check2.Value = 1 Then
'                  'Add By Sindy 2019/6/6 商標處提須增待歸檔,讓信件暫緩存放在此資料區,待後續處理
'                  m_TxtIR20 = IIf(txtIR20 <> "", txtIR20 & "; ", "") & "待歸檔後續處理原因:"
'                  txtIR20 = m_TxtIR20
'               End If
               If txtPI18 = "" Then
                  If txtPI18.Enabled = True Then txtPI18.SetFocus
               ElseIf txtRecvNo.Visible = True And txtRecvNo.Enabled = True Then
                  txtRecvNo.SetFocus
               End If
               Exit Sub
            'End If
         End If
      Next m_iRow
   End With
End Sub

Private Function IPDeptInputForm() As Boolean
'Dim strCP13 As String

   IPDeptInputForm = False
   Screen.MousePointer = vbHourglass 'Add By Sindy 2025/6/19
   If txtPI18 <> "" Or txtPI19 <> "" Then '承辦組的輸入會有新案命名,無需輸入本所案號不須檢查
      '人員要從信件切檔案出來,讓系統自動歸入卷宗區
   '   If Dir(strTMCppFilePath, vbDirectory) = "" Then
   '      MkDir strTMCppFilePath
   '   End If
'      strCP13 = PUB_GetAKindSalesNo(txtPI18, txtPI19, txtPI20, txtPI21) '目前智權人員
      
'      "select tm12,tm15,tm10,'T' as sys_type from trademark" & _
'                  " where tm01='" & txtPI18 & "'" & _
'                    " and tm02='" & txtPI19 & "'" & _
'                    " and tm03='" & txtPI20 & "'" & _
'                    " and tm04='" & txtPI21 & "'" & _
'                  " union
      strExc(0) = "select pa11,pa22,pa09,'P' as sys_type from patent" & _
                  " where pa01='" & txtPI18 & "'" & _
                    " and pa02='" & txtPI19 & "'" & _
                    " and pa03='" & txtPI20 & "'" & _
                    " and pa04='" & txtPI21 & "'" & _
                  " union select sp11,sp14,sp09,'S' as sys_type from servicepractice" & _
                  " where sp01='" & txtPI18 & "'" & _
                    " and sp02='" & txtPI19 & "'" & _
                    " and sp03='" & txtPI20 & "'" & _
                    " and sp04='" & txtPI21 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
'         '外專人員
'         If Left(Pub_StrUserSt03, 2) = "F2" Then
'            If txtPI18 <> "P" And txtPI18 <> "FCP" And txtPI18 <> "FG" Then
'               MsgBox "本所案號中的系統別不正確 !", vbExclamation, "警告！"
'               Exit Function
'            ElseIf txtPI18 = "P" And Pub_StrUserSt03 = "F23" Then
'               '判斷是否為FMP案件
'               'If PUB_FMPtoCheck(1, 2, Pub_strUserST05, txtPI18, txtPI19, txtPI20, txtPI21) = False Then 'FMP寰華案件
'               If PUB_ChkIsFMP(txtPI18, txtPI19, IIf(txtPI20 = "", "0", txtPI20), IIf(txtPI21 = "", "00", txtPI21)) = False Then
'                  MsgBox "非FMP案件，不可操作 !", vbExclamation, "警告！"
'                  Exit Function
'               End If
'            End If
'         End If
            
         m_AppNo = "" & RsTemp.Fields("pa11") '申請案號
         m_RegNo = "" & RsTemp.Fields("pa22") '號數
         If Trim(GRD1.TextMatrix(GRD1.row, 4)) <> txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21 Then
            strExc(0) = "update IPDeptInput set " & _
                        "Ii23='" & txtPI18 & "',Ii24='" & txtPI19 & "'," & _
                        "Ii25='" & txtPI20 & "',Ii26='" & txtPI21 & "'" & _
                        " where Ii01=" & m_strIR01 & _
                        " and Ii02=" & m_strIR02 & _
                        " and Ii03='" & m_strIR03 & "'"
            cnnConnection.Execute strExc(0)
            GRD1.TextMatrix(GRD1.row, 4) = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
         End If
      Else
         Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
         MsgBox "基本檔無此案號！", vbExclamation, "警告！"
         Me.txtPI18.SetFocus
         Exit Function
      End If
   'Add By Sindy 2023/2/9 掛錯案號,可以清掉
   Else
      If txtPI20 <> "" Then txtPI20 = ""
      If txtPI21 <> "" Then txtPI21 = ""
      If Not (Trim(GRD1.TextMatrix(GRD1.row, 4)) = "" And Trim(txtPI18 & txtPI19 & txtPI20 & txtPI21) = "") Then
         If Trim(GRD1.TextMatrix(GRD1.row, 4)) <> txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21 Then
            strExc(0) = "update IPDeptInput set " & _
                        "Ii23='" & txtPI18 & "',Ii24='" & txtPI19 & "'," & _
                        "Ii25='" & txtPI20 & "',Ii26='" & txtPI21 & "'" & _
                        " where Ii01=" & m_strIR01 & _
                        " and Ii02=" & m_strIR02 & _
                        " and Ii03='" & m_strIR03 & "'"
            Pub_SeekTbLog strExc(0)
            cnnConnection.Execute strExc(0)
            If Trim(txtPI18 & txtPI19 & txtPI20 & txtPI21) = "" Then
               GRD1.TextMatrix(GRD1.row, 4) = ""
            Else
               GRD1.TextMatrix(GRD1.row, 4) = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
            End If
         End If
      End If
      '2023/2/9 END
   End If
   
   'Add By Sindy 2024/4/24 因要將外來信夾帶進去該筆案件命名追蹤的資料夾裡面，該MSG檔的檔名一律直接改為「ORDER.msg」
   m_strFullFileName_order = Replace(m_strFullFileName, ".msg", ".ORDER.msg")
   If Dir(m_strFullFileName_order) = "" Then
      If GetAttachFile(GRD1.TextMatrix(GRD1.row, 10), GRD1.TextMatrix(GRD1.row, 11), GRD1.TextMatrix(GRD1.row, 16), "", m_strFullFileName_order) = False Then
'         MsgBox "下載檔案失敗，無法轉寄！", vbExclamation, "警告！"
'         Exit Sub
'      Else
'         Exit For
      End If
   End If
   '2024/4/24 END
   
   Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
   If Pub_StrUserSt03 = "F23" Then '承辦組
      Call Forms(0).SetTmpfrm04010519(Me)
      PopupMenu mdiMain.mnuPopEMail2
      
   Else '程序組
      If UCase(Left(App.EXEName, 7)) = "PATPRO1" Or UCase(Left(App.EXEName, 9)) = "TEPATPRO1" Then '使用外專系統
         Call Forms(0).SetTmpfrm04010519(Me)
         PopupMenu mdiMain.mnuPopEMail3
      Else
         Call Forms(0).SetTmpfrm04010519(Me)
         PopupMenu mdiMain.mnuPopEMail1
      End If
   End If
   
   IPDeptInputForm = True
End Function

'輸入
Private Sub cmdInput_Click()
Dim strIR16 As String
   
   With GRD1
      For m_iRow = 1 To .Rows - 1
         .row = m_iRow
         txtPI18.Tag = m_iRow
         If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
            If dblPrevRow <> .row Then
               MsgBox "資料列選取有誤，請重新確認！", vbExclamation, "警告！"
               Exit Sub
            End If
            
            m_AppNo = "": m_RegNo = ""
            If Pub_StrUserSt03 = "F22" Then '程序組
               If txtPI18 = "" Or txtPI19 = "" Then
                  MsgBox "請輸入本所案號！", vbExclamation, "警告！"
                  Me.txtPI18.SetFocus
                  Exit Sub
               End If
            End If
            If txtPI18 <> "" Then
               If ChkCaseNo(False) = False Then
                  Exit Sub
               ElseIf txtPI18 = "P" And Pub_StrUserSt03 = "F23" Then
                  '判斷是否為FMP案件
                  'If PUB_FMPtoCheck(1, 2, Pub_strUserST05, txtPI18, txtPI19, txtPI20, txtPI21) = False Then 'FMP寰華案件
                  If PUB_ChkIsFMP(txtPI18, txtPI19, txtPI20, txtPI21) = False Then
                     MsgBox "非FMP案件，不可操作 !", vbExclamation, "警告！"
                     Exit Sub
                  End If
               End If
            End If
            
            'Added by Morgan 2023/4/17
            If .TextMatrix(.row, 23) = "8" And Left(.TextMatrix(.row, 29), 1) = "C" Then
               MsgBox "此為2次確認退回信件，需先請電腦中心刪除來函後才可重新輸入！", vbCritical
               Exit Sub
            End If
            'end 2023/4/17
            
            m_strIR01 = .TextMatrix(.row, 10)
            m_strIR02 = .TextMatrix(.row, 11)
            m_strIR03 = .TextMatrix(.row, 16)
            m_strIR04 = .TextMatrix(.row, 13)
            'm_strPi12 = .TextMatrix(.row, 21) '收信日期
            m_strPi12 = .TextMatrix(.row, 17) '轉寄日期
            If Val(m_strPi12) > 0 Then
               m_strPi12 = Val(m_strPi12) - 19110000
            End If
            
            '檢查是否有處理狀態
            If PUB_CheckIRStatus(m_strIR01, m_strIR02, ChgSQL(m_strIR03), ChgSQL(m_strIR04), strIR16) = True Then
               MsgBox "此封信件已有人操作【" & strIR16 & "】，請畫面更新！", vbExclamation, "警告！"
               Exit Sub
            End If
            
            If IPDeptInputForm = False Then
               Exit Sub
            End If
            
            Exit For
         End If
      Next m_iRow
   End With
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 1 Then
      cmdSelCp09.Visible = False: FrameRecv.Enabled = False: FrameRecv.Visible = False
      cmdSelCont.Visible = True: FrameCont.Enabled = True: FrameCont.Visible = True '選擇往來記錄
      txtIR20 = "待來函，暫無後續"
   Else
      cmdSelCp09.Visible = True: FrameRecv.Enabled = True: FrameRecv.Visible = True '選擇總收文號
      cmdSelCont.Visible = False: FrameCont.Enabled = False: FrameCont.Visible = False
      txtIR20 = "" 'Add By Sindy 2022/5/3
   End If
End Sub

Private Sub TextContext_Change()
   PUB_RefreshText TextContext
End Sub

Private Sub txtCOR01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtCOR01_GotFocus()
   CloseIme
   TextInverse txtCOR01
End Sub
Private Sub txtCOR01_Validate(Cancel As Boolean)
Dim strCor03Name As String
Dim strCOR03 As String
   
   If txtCOR01 <> "" Then
      strExc(0) = "SELECT * FROM ContactRecord" & _
            " WHERE CR01 = '" & txtCOR01.Text & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strCOR03 = RsTemp.Fields("CR03")
         txtCOR03 = ChangeCustomerL(strCOR03)
         lbl1 = ""
         If PUB_GetCustData(txtCOR03, strCor03Name) = False Then
            txtCOR03 = ""
            Cancel = True
            txtCOR01_GotFocus
            Exit Sub
         Else
            lbl1 = strCor03Name
         End If
         
      Else
         Cancel = True
         MsgBox "無此往來記錄編號，請重新輸入！", vbCritical + vbOKOnly, "檢核資料"
         txtCOR01_GotFocus
         Exit Sub
      End If
   End If
   If Not CheckLengthIsOK(txtCOR01, txtCOR01.MaxLength) Then
      Cancel = True
      txtCOR01_GotFocus
      Exit Sub
   End If
End Sub

Private Sub txtCOR03_GotFocus()
   CloseIme
   TextInverse txtCOR03
End Sub
Private Sub txtCOR03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtCOR03_Validate(Cancel As Boolean)
Dim strName As String
   
   If txtCOR03 <> "" Then
      If Len(txtCOR03) > 5 Then
         txtCOR03 = ChangeCustomerL(txtCOR03)
         lbl1 = ""
         If PUB_GetCustData(txtCOR03, strName) = False Then
            Cancel = True
            txtCOR03_GotFocus
            Exit Sub
         Else
            lbl1 = strName
         End If
      Else
         Cancel = True
         MsgBox "往來對象編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
         txtCOR03_GotFocus
         Exit Sub
      End If
   End If
   If Not CheckLengthIsOK(txtCOR03, txtCOR03.MaxLength) Then
      Cancel = True
      txtCOR03_GotFocus
      Exit Sub
   End If
End Sub

Private Sub txtIR20_Change()
   PUB_RefreshText txtIR20
End Sub

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
'
'      If txtPI18 <> "P" And txtPI18 <> "FCP" And txtPI18 <> "FG" Then
'         MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
'         Cancel = True
'      End If
'   End If
'   If Cancel Then TextInverse txtPI18
'End Sub

Private Sub txtPI19_GotFocus()
   TextInverse txtPI19
End Sub

'Add By Sindy 2022/9/23
Private Sub txtPI19_LostFocus()
   If txtPI18 = "" And txtPI19 = "" Then
      txtPI20 = ""
      txtPI21 = ""
      txtRecvNo = ""
   End If
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
   If txtPI18 <> "" And txtPI19 <> "" Then
      If txtPI20 = "" Then txtPI20 = "0"
      If txtPI21 = "" Then txtPI21 = "00"
   End If
End Sub

Private Sub txtRecvNo_GotFocus()
   TextInverse txtRecvNo
End Sub

'Added by Morgan 2023/4/11
Private Function IsReKeyInCase(pCP09 As String, pIR16nm As String) As Boolean
   Dim strSql As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   'Modified by Morgan 2023/5/23 +pIR16nm
   If Pub_StrUserSt03 = "F22" And pIR16nm = "輸入" Then 'Added by Morgan 2023/4/18
      If Left(pCP09, 1) = "C" Then
         intQ = 1
         'Modified by Morgan 2024/9/26 + or cp142>0 (專利權期限補償核准可能沒有期限但也要輸入新的專利權期滿終止日)
         strSql = "select cp07 from caseprogress where cp09='" & pCP09 & "' and (cp07>0 or cp142>0)"
         Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
         If intQ = 1 Then
            IsReKeyInCase = True
         End If
      End If
   End If
   
   Set rsQuery = Nothing
End Function
