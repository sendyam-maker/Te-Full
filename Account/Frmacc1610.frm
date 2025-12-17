VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1610 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "請款單列印"
   ClientHeight    =   6216
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7164
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   7170
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "測試DOC列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5610
      TabIndex        =   38
      Top             =   5070
      Value           =   1  '核取
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.OptionButton Option2 
      Caption         =   "發票　套表，點陣印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7440
      TabIndex        =   37
      Top             =   270
      Width           =   3345
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   675
      TabIndex        =   31
      Top             =   4860
      Width           =   4755
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "高所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   3600
         TabIndex        =   35
         Top             =   120
         Width           =   705
      End
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "南所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2710
         TabIndex        =   34
         Top             =   120
         Width           =   705
      End
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "中所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1820
         TabIndex        =   33
         Top             =   120
         Width           =   705
      End
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "北所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   930
         TabIndex        =   32
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "所別："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   36
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   7770
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1620
      Width           =   3450
   End
   Begin VB.TextBox txtAdd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "Y"
      Top             =   1950
      Width           =   585
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   405
      Left            =   750
      TabIndex        =   27
      Top             =   240
      Width           =   3585
      Begin VB.OptionButton Option2 
         Caption         =   "請款單　白紙，雷射印表機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   60
         Value           =   -1  'True
         Width           =   3555
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   7770
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   1050
      Width           =   3450
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   990
      Width           =   3450
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2415
      MaxLength       =   10
      TabIndex        =   12
      Top             =   3570
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "重新列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1230
      TabIndex        =   11
      Top             =   3300
      Width           =   1125
   End
   Begin VB.OptionButton Option1 
      Caption         =   "一般列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1230
      TabIndex        =   6
      Top             =   2280
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.CheckBox chk 
      Caption         =   "列印客戶案件案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   750
      TabIndex        =   2
      Top             =   1470
      Value           =   1  '核取
      Width           =   2235
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4335
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2910
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2415
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2910
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1260
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   5610
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   2415
      TabIndex        =   7
      Top             =   2550
      Width           =   1575
      _ExtentX        =   2794
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   4335
      TabIndex        =   8
      Top             =   2550
      Width           =   1575
      _ExtentX        =   2794
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   315
      Left            =   2415
      TabIndex        =   13
      Top             =   3930
      Width           =   1575
      _ExtentX        =   2794
      _ExtentY        =   550
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   315
      Left            =   2415
      TabIndex        =   14
      Top             =   4290
      Width           =   1575
      _ExtentX        =   2794
      _ExtentY        =   550
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox5 
      Height          =   315
      Left            =   4335
      TabIndex        =   15
      Top             =   4290
      Width           =   1575
      _ExtentX        =   2794
      _ExtentY        =   550
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label15 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "發票地址條若要印出                      請至財務Mail資料維護                   勾選(寄紙本)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   7590
      TabIndex        =   30
      Top             =   2760
      Width           =   3210
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   29
      Top             =   1410
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "列印地址條：            (Y : 是)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   28
      Top             =   1980
      Width           =   2805
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "發票印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   26
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "請款單印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   750
      TabIndex        =   25
      Top             =   750
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "列印日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1455
      TabIndex        =   24
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4065
      TabIndex        =   23
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "列印時間"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1455
      TabIndex        =   22
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "編　　號                       以後(含)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1455
      TabIndex        =   21
      Top             =   3600
      Width           =   4440
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2505
      Left            =   675
      Top             =   2220
      Width           =   5850
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "編　　號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1455
      TabIndex        =   20
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4095
      TabIndex        =   19
      Top             =   2940
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   135
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4095
      TabIndex        =   18
      Top             =   2580
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "日　　期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1455
      TabIndex        =   17
      Top             =   2580
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/01 改成Form2.0 ; 地址條改成Excel列印
'Memo By Sindy 2022/1/17 Form2.0已修改
'Create By Sindy 2013/12/4
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Dim strNo As String
Dim strPrinter As String
Dim prnPrint As Printer
Dim strSqlMain As String
Dim bolPrintAddr As Boolean 'Add By Sindy 2014/5/15
Dim strChkDate As String, bolEInvoice As Boolean 'Add By Sindy 2019/7/2
Public ProState As String '權限: 1.全所 2.該所
Dim m_sqlST06 As String
Dim m_FileName As String 'Add By Sindy 2021/12/29


Private Sub Command1_Click()
Dim bolHaveData As Boolean
Dim bolChk As Boolean, ii As Integer
   
   'Add By Sindy 2021/5/21
   bolChk = False
   m_sqlST06 = " and ST06 in("
   For ii = 0 To 3
      If ChkST06(ii).Value = 1 Then
         bolChk = True
         m_sqlST06 = m_sqlST06 & "'" & ii + 1 & "'"
      End If
   Next ii
   If bolChk = False Then
      MsgBox "請勾選所別！", vbExclamation
      ChkST06(0).SetFocus
      m_sqlST06 = ""
      Exit Sub
   End If
   m_sqlST06 = m_sqlST06 & ")"
   m_sqlST06 = Replace(m_sqlST06, "''", "','")
   '2021/5/21 END
   
   Screen.MousePointer = vbHourglass
   If Option2(0).Value = True Then '請款單
      PUB_SetOsDefaultPrinter Combo1 'Add by Sindy 2022/3/14 切換Word/Excel印表機
      PUB_RestorePrinter Combo1
      If Option1(0).Value = True Then
         PrintDoc
      Else
         PrintDocX
      End If
      PUB_SetOsDefaultPrinter strPrinter 'Add by Sindy 2022/3/14 切換Word/Excel印表機
      PUB_RestorePrinter strPrinter
      
   ElseIf Option2(1).Value = True Then '發票
      'Modify By Sindy 2019/7/2
      If Option1(0).Value = True Then
         strChkDate = Val(FCDate(MaskEdBox2.Text))
      Else
         strChkDate = Val(FCDate(MaskEdBox3.Text))
      End If
      If Val(strChkDate) < Val(1080701) Then '紙本發票
         bolEInvoice = False
         PUB_RestorePrinter Combo2
         MsgBox "請放置發票套表紙於選取的印表機!!", vbInformation
      Else
         bolEInvoice = True
      End If
      '2019/7/2 END
      If Option1(0).Value = True Then
         bolHaveData = PrintDoc_Inv
      Else
         bolHaveData = PrintDocX_Inv
      End If
      If bolHaveData = True Then
         '若是否列印地址條上"Y"
         bolPrintAddr = False
         If Me.txtAdd.Text = "Y" Then
            'Modify By Sindy 2014/4/1
            'MsgBox "請放地址條貼紙於選取的印表機!!", vbInformation
            'Modify By Sindy 2014/4/17
            If MsgBox("是否要列印地址條？" & vbCrLf & _
                      "若要印，請放地址條貼紙於選取的印表機!!", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
               bolPrintAddr = True
               PUB_SetOsDefaultPrinter Combo3 'Added by Lydia 2022/03/01 切換Word/Excel印表機
               PUB_RestorePrinter Combo3
               'PrintAddress '列印地址條
            End If
            '2014/4/1 END
         End If
         'Modify By Sindy 2014/5/15 改到外面來,因不管有無要印地址條都需Run此函數,因要更新列印次數
         PrintAddress '列印地址條
         '2014/5/15 END
         Frmacc1611.Hide
         If Frmacc1611.OpenTable = True Then
            Frmacc1611.Show
            Me.Enabled = False
         Else
            Unload Frmacc1611
         End If
      End If
      PUB_SetOsDefaultPrinter strPrinter 'Added by Lydia 2022/03/01 切換Word/Excel印表機
      PUB_RestorePrinter strPrinter
   End If
   EndOfficeAp 'Added by Morgan 2025/9/10 印完要清除物件，否則印表機不會變
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2014/4/1
'*************************************************
'  列印地址條
'
'*************************************************
Private Sub PrintAddress()
   Dim intCounter As Integer
   Dim strAddr As String, strCustName As String
   Dim intR As Integer
   Dim intTimes As Integer
   Dim strA0K04 As String
   Dim intLine As Integer
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
   'Add by Amy 2019/07/25
   'Dim intFCount As Integer '印地址條筆數
   'Dim strCU179 As String '電子發票寄送方式
   Dim strTempAddressList As String 'Added by Lydia 2022/03/01
    
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
'   If InStr(UCase(strSqlMain), UCase("ORDER")) > 0 Then
'      strSqlMain = Mid(strSqlMain, 1, InStr(UCase(strSqlMain), UCase("ORDER")) - 1)
'   End If
'   strSqlMain = strSqlMain & " order by a0k04 asc"
   'Modify By Sindy 2015/1/20 寫暫存
   strSqlMain = "select T02 as a0k03,T15 as a0k04 from ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "' order by T15 asc"
   '2015/1/20 END
   adoacc0k0.Open strSqlMain, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount <> 0 Then
      Screen.MousePointer = vbHourglass
   Else
      adoacc0k0.Close
      ShowNoData
      Exit Sub
   End If
   strA0K04 = ""
   With adoacc0k0
      .MoveFirst
      'Remove by Lydia 2022/03/01 改用Excel列印
      'If bolPrintAddr = True Then 'Modify By Sindy 2014/5/15 +if
      '   'intCounter = 0: intLine = 0 'Modify By Sindy 2016/11/22 Mark
      '   'XP自定紙張需手動設定並將印表機預設為該紙張
      '   '9x
      '   If pub_OS = "1" Then
      '      Printer.Height = 2880
      '      Printer.Width = 13000
      '   Else
      '      Printer.PaperSize = PUB_GetPaperSize(2)
      '   End If
      '
      '   Printer.Font = "@新細明體"
      '   Printer.FontSize = 12
      'End If
      'end 2022/03/01
      Do While .EOF = False
         If bolPrintAddr = True Then 'Modify By Sindy 2014/5/15 +if
            intCounter = 0: intLine = 0 'Add By Sindy 2016/11/22
            If strA0K04 <> .Fields("a0k04") Then
               strAddr = "" '地址
               strCustName = Trim("" & .Fields("a0k04").Value) '客戶名稱
               'Modify by Amy 2019/07/25 改寫至GetInvoiceAddress
'               '收據抬頭為3個字以下(含3個字)以a0k03為主
'               If Len(Trim("" & .Fields("a0k04").Value)) <= 3 Then
'                  strSql = "select cu01,cu02,cu04,cu112,cu23,cu30,cu31 from customer where cu01='" & Left(Trim("" & .Fields("a0k03").Value), 8) & "' and cu02='" & Mid(Trim("" & .Fields("a0k03").Value), 9, 1) & "' "
'                  intR = 1
'                  Set RsTemp = ClsLawReadRstMsg(intR, strSql)
'                  If intR = 1 Then
'                     '聯絡地址優先
'                     If Trim("" & RsTemp.Fields("cu31")) <> "" Then
'                        strAddr = Trim("" & RsTemp.Fields("cu30")) & " " & Trim("" & RsTemp.Fields("cu31"))
'                     '再來中文地址
'                     ElseIf Trim("" & RsTemp.Fields("cu23")) <> "" Then
'                        strAddr = Trim("" & RsTemp.Fields("cu112")) & " " & Trim("" & RsTemp.Fields("cu23"))
'                     End If
'                  End If
'               Else
'                  '以收據抬頭抓客戶檔之CU04,若為客戶資料則抓中文地址
'                  '                         若不存在,則再抓收據抬頭資料檔acc420的營業地址
'                  'Modify By Sindy 2014/8/15 +and (cu80 is null or cu80='其他' or cu80='業務自行處理')
'                  'Modified by Morgan 2015/4/21 收據抬頭取消Trim,否則遇有造字可能會錯
'                  strSql = "select cu01,cu02,cu04,cu112,cu23,cu30,cu31 from customer where cu04='" & .Fields("a0k04").Value & "' and (cu80 is null or cu80='其他' or cu80='業務自行處理') and cu02=0 "
'                  intR = 1
'                  Set RsTemp = ClsLawReadRstMsg(intR, strSql)
'                  If intR = 1 Then
'                     '聯絡地址優先
'                     If Trim("" & RsTemp.Fields("cu31")) <> "" Then
'                        strAddr = Trim("" & RsTemp.Fields("cu30")) & " " & Trim("" & RsTemp.Fields("cu31"))
'                     '再來中文地址
'                     ElseIf Trim("" & RsTemp.Fields("cu23")) <> "" Then
'                        strAddr = Trim("" & RsTemp.Fields("cu112")) & " " & Trim("" & RsTemp.Fields("cu23"))
'                     End If
'                  Else
'                     'Modified by Morgan 2015/4/21 收據抬頭取消Trim,否則遇有造字可能會錯
'                     strSql = "select * from acc420 where a4201='" & .Fields("a0k04").Value & "' "
'                     intR = 1
'                     Set RsTemp = ClsLawReadRstMsg(intR, strSql)
'                     If intR = 1 Then
'                        '郵寄地址優先
'                        If Trim("" & RsTemp.Fields("a4203")) <> "" Then
'                           strAddr = Trim("" & RsTemp.Fields("a4203"))
'                        '再來營業地址
'                        ElseIf Trim("" & RsTemp.Fields("a4215")) <> "" Then
'                           strAddr = Trim("" & RsTemp.Fields("a4215"))
'                        End If
'                     End If
'                  End If
'               End If
                'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
                'strAddr = GetInvoiceAddress(Trim("" & .Fields("a0k03").Value), "" & .Fields("a0k04").Value, False, strCU179)
               'end 2019/07/25
               
               'Modify By Sindy 2016/11/28
'               'Modify by Amy 2019/07/25 電子發票寄送方式為 1.紙本才印地址條
'               If strCU179 = "1" Then
'                    intFCount = intFCount + 1
'                    'Modified by Lydia 2022/03/01 傳入多張地址條的內容；用|區隔不同張地址條，同一張地址條用$區隔地址和收件人
'                    'Call PUB_PrintAccAddress(strAddr, strCustName)
'                    If strAddr & strCustName <> "" Then strTempAddressList = strTempAddressList & Trim(strAddr) & "$" & Trim(strCustName) & "|"
'               End If
                'end 2019/07/25
                'end 2024/06/13
'               '地址
'               Printer.CurrentX = 100
'               Printer.CurrentY = 300 + 2200 * intCounter
'               If strAddr = "" Then
'                  Printer.Print ""
'               Else
'                  '控制折行
'                  'PUB_PrintAddress strAddr, intCounter, 0
'                  PUB_PrintAddress strAddr, intCounter, intLine
'               End If
'               '收件人
'               Printer.CurrentX = 100
'               'Printer.CurrentY = 1000 + 2200 * intCounter
'               Printer.CurrentY = 900 + 2200 * intCounter + 300 * intLine
'               If strCustName = "" Then
'                  Printer.Print ""
'               Else
'                  'Printer.Print "　　　" & strCustName & MsgText(104)
'                  '控制折行
'                  PUB_PrintTitle strCustName & MsgText(104), intCounter, intLine
'               End If
'               Printer.NewPage
               '2016/11/28 END
            End If
            strA0K04 = .Fields("a0k04")
         End If
         
         'Modify By Sindy 2015/1/20 改回到aacc_fun更新
'         If Option1(0).Value = True Then
'            '列印次數
'            If IsNull(.Fields("a4306").Value) Then
'               intTimes = 1
'            Else
'               intTimes = Val(.Fields("a4306").Value) + 1
'            End If
'            '更新發票的列印次數及列印時間
'            adoTaie.Execute "update acc430 set a4306=" & intTimes & ",a4307=sysdate where a4301='" & .Fields("a4301").Value & "'"
'         End If
         
         .MoveNext
      Loop
   End With
   'Modify By Sindy 2016/11/28 Mark
'   If bolPrintAddr = True Then 'Modify By Sindy 2014/5/15 +if
'      Printer.Font = "新細明體"
'      Printer.EndDoc
'   End If
   '2016/11/28 END
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
   'Add by Amy 2019/07/25 無地址條要印彈息訊
'   If intFCount = 0 Then
'        ShowNoData
'   'Added by Lydia 2022/03/01 改用Execl列印地址條
'   Else
'        If strTempAddressList <> "" Then
'            If PUB_XlsAccAddress(strTempAddressList) = False Then
'                MsgBox "列印失敗！", vbCritical
'            End If
'        End If
'   'end 2022/03/01
'   End If
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2021/5/21
Private Sub Form_Activate()
Dim intST06 As Integer, ii As Integer
   
   For ii = 0 To 3
      ChkST06(ii).Value = 0
   Next ii
   
   intST06 = Val(PUB_GetST06(strUserNum)) - 1
   ChkST06(intST06).Value = 1
   '權限: 1.全所 2.該所
   If ProState = "2" Then
      For ii = 0 To 3
         ChkST06(ii).Enabled = False
      Next ii
   Else
      For ii = 0 To 3
         ChkST06(ii).Enabled = True
      Next ii
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   '改單線固定(調整大小不用再設定)
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
      For intY = 0 To Int(ScaleHeight / sglHeight)
         PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
      Next
   Next
   
   'Add By Sindy 2021/12/29
   m_FileName = "$$J公司請款單.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileName) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileName
   End If
   Call PUB_GetSampleFile(m_FileName, "M31-000012-0-00", , App.path & "\" & strUserNum & "\")
   '2021/12/29 END
   
'   'MODIFY BY SONIA 2014/3/21 預設上一個工作日至當天
'   'MaskEdBox1.Text = CFDate(strSrvDate(2))
'   MaskEdBox1.Text = CFDate(TransDate(PUB_GetWorkDay1(strSrvDate(1) - 1, 1), 1))
'   '2014/3/21 END
'   MaskEdBox1.Mask = DFormat
'   MaskEdBox2.Text = CFDate(strSrvDate(2))
'   MaskEdBox2.Mask = DFormat
'
'   MaskEdBox3.Mask = DFormat
'   MaskEdBox4.Mask = Tformat
'   MaskEdBox5.Mask = Tformat
   Call Option2_Click(0) 'Modify By Sindy 2015/3/24
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   PUB_SetPrinter Me.Name, Combo2
   PUB_SetPrinter Me.Name, Combo3
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '若印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '若印表機變動, 則更新列印設定
   If Me.Combo3.Text <> Me.Combo3.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo3.Name, "0", "0", Me.Combo3.Text
   End If
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1610 = Nothing
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Me.Text2(0).Text = Empty
   Me.Text2(1).Text = Empty
   Text2(2).Text = Empty
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = ""
   MaskEdBox4.Text = ""
   MaskEdBox4.Mask = Tformat
   MaskEdBox5.Mask = ""
   MaskEdBox5.Text = ""
   MaskEdBox5.Mask = Tformat
   Option1(0).Value = True
   'Option2(0).Value = True 'Add By Sindy 2014/8/15 瑞婷說不要還原
   'txtAdd.Text = ""
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = MaskEdBox1.Text
   MaskEdBox2.Mask = DFormat
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      MaskEdBox1.Enabled = True
      MaskEdBox2.Enabled = True
      Text2(0).Enabled = True
      Text2(1).Enabled = True
      
      Text2(2).Enabled = False
      MaskEdBox3.Enabled = False
      MaskEdBox4.Enabled = False
      MaskEdBox5.Enabled = False
      'Add By Sindy 2014/4/1
      If Option2(1).Value = True Then
         txtAdd = "Y"
      Else
         txtAdd = ""
      End If
      '2014/4/1 END
   Else
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = False
      Text2(0).Enabled = False
      Text2(1).Enabled = False
      
      Text2(2).Enabled = True
      MaskEdBox3.Enabled = True
      MaskEdBox4.Enabled = True
      MaskEdBox5.Enabled = True
      
      Text2(2).Text = ""
      MaskEdBox3.Mask = ""
      MaskEdBox3.Text = ""
      MaskEdBox3.Mask = DFormat
      MaskEdBox4.Mask = ""
      MaskEdBox4.Text = ""
      MaskEdBox4.Mask = Tformat
      MaskEdBox5.Mask = ""
      MaskEdBox5.Text = ""
      MaskEdBox5.Mask = Tformat
      Text2(2).SetFocus
      txtAdd = "" 'Add By Sindy 2014/4/1
   End If
End Sub

Private Sub Option2_Click(Index As Integer)
   '預設上一個工作日至當天
   MaskEdBox1.Text = CFDate(TransDate(PUB_GetWorkDay1(strSrvDate(1) - 1, 1), 1))
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Text = CFDate(strSrvDate(2))
   MaskEdBox2.Mask = DFormat

   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = Tformat
   MaskEdBox5.Mask = Tformat
   '2015/3/24 END
   
   Text2(0) = ""
   Text2(1) = ""
   Text2(2) = ""
   If Index = 0 Then
      Text2(0).MaxLength = 9
      Text2(1).MaxLength = 9
      Text2(2).MaxLength = 9
      txtAdd = ""
   Else
      Text2(0).MaxLength = 10
      Text2(1).MaxLength = 10
      Text2(2).MaxLength = 10
      txtAdd = "Y"
   End If
End Sub

Private Sub Text2_Change(Index As Integer)
   If Index = 2 Then
      If Option2(0).Value = True Then '請款單
         If Len(Text2(Index)) <> 9 Then Exit Sub
         strExc(0) = "select to_char(a0k14,'yyyymmdd'),to_char(a0k14,'hh24miss'),to_char(sysdate,'hh24miss'),a0k11 from acc0k0 where a0k01='" & Text2(2) & "' and a0k14 is not null and a0k11='J'"
      Else '發票
         If Len(Text2(Index)) <> 10 Then Exit Sub
         strExc(0) = "select to_char(a4307,'yyyymmdd'),to_char(a4307,'hh24miss'),to_char(sysdate,'hh24miss') from acc430 where a4301='" & Text2(2) & "' and a4307 is not null"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MaskEdBox3.Mask = ""
         MaskEdBox3.Text = Format(TransDate(RsTemp(0), 1), DFormat)
         MaskEdBox3.Mask = DFormat
         MaskEdBox4.Mask = ""
         MaskEdBox4.Text = Format(RsTemp(1), Tformat)
         If Len(MaskEdBox4.Text) = 7 Then MaskEdBox4.Text = "0" & MaskEdBox4.Text
         MaskEdBox4.Mask = Tformat
         If RsTemp(0) < strSrvDate(1) Then
            MaskEdBox5.Mask = ""
            MaskEdBox5.Text = Format("235959", Tformat)
            If Len(MaskEdBox5.Text) = 7 Then MaskEdBox5.Text = "0" & MaskEdBox5.Text
            MaskEdBox5.Mask = Tformat
         Else
            MaskEdBox5.Mask = ""
            MaskEdBox5.Text = Format(RsTemp(2), Tformat)
            If Len(MaskEdBox5.Text) = 7 Then MaskEdBox5.Text = "0" & MaskEdBox5.Text
            MaskEdBox5.Mask = Tformat
         End If
      End If
   End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   TextInverse Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   Me.Text2(Index).Text = UCase(Me.Text2(Index).Text)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text2(0) <> MsgText(601) And Text2(0) <> MsgText(802) Then
      FormCheck = True
      Exit Function
   End If
   If Text2(1) <> MsgText(601) And Text2(1) <> MsgText(802) Then
      FormCheck = True
      Exit Function
   End If
   
   FormCheck = False
End Function

Private Function FormCheck1() As Boolean

   If MaskEdBox3.Text = MsgText(29) Then
      MsgBox "請輸入列印日期！", vbExclamation
      MaskEdBox3.SetFocus
      Exit Function
   ElseIf MaskEdBox3.Text <> CFDate(strSrvDate(2)) Then
      If MsgBox("列印日期不是當天，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
         MaskEdBox3.SetFocus
         Exit Function
      End If
   End If
   If MaskEdBox4.Text = "__:__:__" Then
      MsgBox "請輸入列印時間(起)！", vbExclamation
      MaskEdBox4.SetFocus
      Exit Function
   End If
   If MaskEdBox5.Text = "__:__:__" Then
      MsgBox "請輸入列印時間(迄)！", vbExclamation
      MaskEdBox5.SetFocus
      Exit Function
   End If
   
   FormCheck1 = True
End Function

'一般列印
Private Sub PrintDoc()
   
   If FormCheck = False Then
      MsgBox "請輸入列印條件！", vbExclamation
      Exit Sub
   End If
   
On Error GoTo Checking:
   
   strSqlMain = ""
   strSql = ""
   
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Me.Text2(0).Text <> MsgText(601) And Me.Text2(0).Text <> MsgText(29) Then
      strSql = strSql & " and a0k01 >= '" & Me.Text2(0).Text & "' "
   End If
   If Me.Text2(1).Text <> MsgText(601) And Me.Text2(1).Text <> MsgText(29) Then
      strSql = strSql & " and a0k01 <= '" & Me.Text2(1).Text & "' "
   End If
   
   If strSql = "" Then
      MsgBox "【日期】及【編號】條件不可同時空白！", vbExclamation
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
'Modify By Sindy 2024/5/16:
'在一般列印時會同時將已達可列印條件收據會一併列印出來 (不管收據日期), 請修改:
'1. 選擇收據日期條件時才將收據號碼範圍列印時，才將已達可列印條件收據會一併列印出來(不管收據日期)；
'2. 選擇收據號碼條件時，只列印符合收據號碼且可列印的收據印出，不在此號碼區間之可列印收據也不列印。
   strSqlMain = "select * from ( " & _
                        "select * from acc0k0,customer,acc0j0,caseprogress,staff " & _
                        "where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) " & _
                        "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                        "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) " & _
                        "and a0k32 IS NULL and a0k11='J' and a0j13=a0k01 and a0j01=cp09 and a0k20=st01(+) " & strSql & m_sqlST06
   'Modify By Sindy 2024/5/16
   If Me.Text2(0).Text = MsgText(601) And Me.Text2(1).Text = MsgText(601) Then
   'Sindy 2024/5/16 END
      strSqlMain = strSqlMain & _
                        " Union All " & _
                        "select * from acc0k0,customer,acc0j0,caseprogress,staff " & _
                        "where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) " & _
                        "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                        "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) " & _
                        "and (a0k32 ='Y') and a0k11='J' and a0j13=a0k01 and a0j01=cp09 and a0k20=st01(+) " & m_sqlST06
   End If
      strSqlMain = strSqlMain & _
                        ") order by a0k01 asc"
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open strSqlMain, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      Screen.MousePointer = vbDefault
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
'   '控制 9x 才自訂
'   If pub_OS = "1" Then
'      Printer.Height = 8750
'      Printer.Width = 13000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(1)
'   End If
   Printer.FontSize = 12
   Printer.Font = "標楷體"
   
   '初始化記錄編號的變數
   strNo = ""
   Do While adoacc0k0.EOF = False
      '改呼叫共用(與補印請款單統一)
      If strNo <> adoacc0k0.Fields("a0k01").Value Then
         'Modify By Sindy 2022/1/17
         If Check1.Value = 1 Then
            PUB_PrintCaseReceipt_J_Doc App.path & "\" & strUserNum & "\" & m_FileName, adoacc0k0, 0, Me.chk(1)
         Else
         '2022/1/17 END
'            If strNo <> "" Then Printer.NewPage
'            PUB_PrintCaseReceipt_J adoacc0k0, 0, Me.chk(1)
         End If
         strNo = adoacc0k0.Fields("a0k01").Value
      End If
      adoacc0k0.MoveNext
   Loop
   '改呼叫共用(與補印請款單統一)
   adoacc0k0.Close
   'Modify By Sindy 2022/1/17
   If Check1.Value = 0 Then
      Printer.EndDoc
   End If
   '2022/1/17 END
   Printer.Font = "新細明體"
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   Exit Sub
   
Checking:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

'重新列印
Private Sub PrintDocX()
   
   If FormCheck1() = False Then
      Exit Sub
   End If
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   
   strSqlMain = ""
   strSql = ""
   
   strSql = " and a0k14>=to_date('" & DBDATE(MaskEdBox3.Text) & Replace(MaskEdBox4.Text, ":", "") & "','YYYYMMDDHH24MISS')"
   strSql = strSql & " and a0k14<=to_date('" & DBDATE(MaskEdBox3.Text) & Replace(MaskEdBox5.Text, ":", "") & "','YYYYMMDDHH24MISS')"
   strSqlMain = "select * from acc0k0,acc0j0,caseprogress,staff where a0k19=1 and a0k11='J' and (a0k09 is null or a0k09 = 0) and a0j13(+)=a0k01 and a0j01=cp09(+) and a0k20=st01(+) " & strSql & m_sqlST06
   intI = 1
   Set adoacc0k0 = ClsLawReadRstMsg(intI, strSqlMain)
   If intI = 0 Then
      adoacc0k0.Close
      Screen.MousePointer = vbDefault
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
'   If pub_OS = "1" Then
'      Printer.Height = 8750
'      Printer.Width = 13000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(1)
'   End If
   Printer.FontSize = 12
   Printer.Font = "標楷體"
   
   strNo = ""
   Do While adoacc0k0.EOF = False
      If strNo <> adoacc0k0.Fields("a0k01").Value Then
         'Modify By Sindy 2022/1/17
         If Check1.Value = 1 Then
            PUB_PrintCaseReceipt_J_Doc App.path & "\" & strUserNum & "\" & m_FileName, adoacc0k0, 0, Me.chk(1), , , , , , False
         Else
         '2022/1/17 END
'            If strNo <> "" Then Printer.NewPage
'            PUB_PrintCaseReceipt_J adoacc0k0, 0, Me.chk(1), , , , , , False
         End If
         strNo = adoacc0k0.Fields("a0k01").Value
      End If
      adoacc0k0.MoveNext
   Loop
   adoacc0k0.Close
   'Modify By Sindy 2022/1/17
   If Check1.Value = 0 Then
      Printer.EndDoc
   End If
   '2022/1/17 END
   Printer.Font = "新細明體"
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   Exit Sub
   
ErrHnd:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

Private Sub txtAdd_GotFocus()
   TextInverse Me.txtAdd
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case KeyAscii
   Case 8, 89
       '無動作
   Case Else
       KeyAscii = 0
   End Select
End Sub

'發票一般列印
Private Function PrintDoc_Inv() As Boolean
   
   If FormCheck = False Then
      MsgBox "請輸入列印條件！", vbExclamation
      Exit Function
   End If
   
On Error GoTo Checking
   
   PrintDoc_Inv = False
   strSqlMain = ""
   strSql = ""
   
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a4302 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a4302 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Me.Text2(0).Text <> MsgText(601) And Me.Text2(0).Text <> MsgText(29) Then
      strSql = strSql & " and a4301 >= '" & Me.Text2(0).Text & "' "
   End If
   If Me.Text2(1).Text <> MsgText(601) And Me.Text2(1).Text <> MsgText(29) Then
      strSql = strSql & " and a4301 <= '" & Me.Text2(1).Text & "' "
   End If
   'Add By Sindy 2019/7/15 增加判斷：108/7以後發票的地址條都不考慮列印次數
   If bolEInvoice = False Then '紙本發票
      strSql = strSql & " and nvl(a4306,0)=0 "
   End If
   
   If strSql = "" Then
      MsgBox "【日期】及【編號】條件不可同時空白！", vbExclamation
      MaskEdBox1.SetFocus
      Exit Function
   End If
   
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2016/1/18 +,a0k33
   'Modify By Sindy 2020/11/11 + " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')"
   strSqlMain = "select acc430.*,axc02,a0k20,st02,st15,a0k03,a0k04,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as cuname,a0k33" & _
                " from acc430,acc431,acc0k0,customer,staff" & _
                " where a4301=axc01(+)" & _
                " and axc02=a0k01(+)" & _
                " and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)" & _
                " and a0k20=st01(+) and (a4308 is null or a4308 = 0)" & strSql & m_sqlST06 & _
                " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')" & _
                " order by a4301 asc"
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open strSqlMain, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      Screen.MousePointer = vbDefault
      MsgBox MsgText(28), , MsgText(5)
      Exit Function
   Else
      PrintDoc_Inv = True
   End If
   
'   '控制 9x 才自訂
'   If pub_OS = "1" Then
'      Printer.Height = 8750
'      Printer.Width = 13000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(1)
'   End If
   'Modify By Sindy 2019/7/2 增加紙本發票的判斷
   If bolEInvoice = False Then Printer.FontSize = 12
   If bolEInvoice = False Then Printer.Font = "標楷體"
   
   '初始化記錄編號的變數
   strNo = ""
   'Add By Sindy 2015/1/20 先刪暫存
   adoTaie.Execute "delete from ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   '2015/1/20 END
   Do While adoacc0k0.EOF = False
      '改呼叫共用(與補印發票統一)
      If strNo <> adoacc0k0.Fields("a4301").Value Then
         If strNo <> "" Then
            If bolEInvoice = False Then Printer.NewPage
         End If
         'Add By Sindy 2015/1/20 寫暫存
         'T01 : 發票號碼
         'T02 : 客戶編號
         'T15 : 收據抬頭
         'Modified by Morgan 2015/4/21 收據抬頭取消Trim,否則遇有造字可能會錯
         adoTaie.Execute "insert into ACCTMP08(T01,T02,T05,T14,T15) values('" & adoacc0k0.Fields("a4301") & "','" & Trim("" & adoacc0k0.Fields("a0k03")) & "'," & _
                                                                          "'" & Me.Name & "','" & strUserNum & "','" & adoacc0k0.Fields("a0k04") & "')"
         '2015/1/20 END
         If bolEInvoice = False Then PUB_PrintCaseReceipt_Inv adoacc0k0, txtAdd
         strNo = adoacc0k0.Fields("a4301").Value
      End If
      adoacc0k0.MoveNext
   Loop
   '改呼叫共用(與補印發票統一)
   adoacc0k0.Close
   If bolEInvoice = False Then Printer.EndDoc
   If bolEInvoice = False Then Printer.Font = "新細明體"
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   Exit Function
   
Checking:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Function

'發票重新列印
Private Function PrintDocX_Inv() As Boolean
   
   If FormCheck1() = False Then
      Exit Function
   End If
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   
   PrintDocX_Inv = False
   strSqlMain = ""
   strSql = ""
   
   strSql = " and a4307>=to_date('" & DBDATE(MaskEdBox3.Text) & Replace(MaskEdBox4.Text, ":", "") & "','YYYYMMDDHH24MISS')"
   strSql = strSql & " and a4307<=to_date('" & DBDATE(MaskEdBox3.Text) & Replace(MaskEdBox5.Text, ":", "") & "','YYYYMMDDHH24MISS')"
   'Add By Sindy 2019/7/15 增加判斷：108/7以後發票的地址條都不考慮列印次數
   If bolEInvoice = False Then '紙本發票
      strSql = strSql & " and a4306=1 "
   End If
   
   'Modify By Sindy 2016/1/18 +,a0k33
   'Modify By Sindy 2020/11/11 + " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')"
   strSqlMain = "select acc430.*,axc02,a0k20,st02,st15,a0k03,a0k04,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as cuname,a0k33" & _
                " from acc430,acc431,acc0k0,customer,staff" & _
                " where a4301=axc01(+)" & _
                " and axc02=a0k01(+)" & _
                " and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)" & _
                " and a0k20=st01(+) and (a4308 is null or a4308 = 0)" & strSql & m_sqlST06 & _
                " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')" & _
                " order by a4301 asc"
   intI = 1
   Set adoacc0k0 = ClsLawReadRstMsg(intI, strSqlMain)
   If intI = 0 Then
      adoacc0k0.Close
      Screen.MousePointer = vbDefault
      MsgBox MsgText(28), , MsgText(5)
      Exit Function
   Else
      PrintDocX_Inv = True
   End If
   
'   If pub_OS = "1" Then
'      Printer.Height = 8750
'      Printer.Width = 13000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(1)
'   End If
   'Modify By Sindy 2019/7/2 增加紙本發票的判斷
   If bolEInvoice = False Then Printer.FontSize = 12
   If bolEInvoice = False Then Printer.Font = "標楷體"
   
   strNo = ""
   'Add By Sindy 2015/1/20 先刪暫存
   adoTaie.Execute "delete from ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   '2015/1/20 END
   Do While adoacc0k0.EOF = False
      If strNo <> adoacc0k0.Fields("a4301").Value Then
         If strNo <> "" Then
            If bolEInvoice = False Then Printer.NewPage
         End If
         'Add By Sindy 2015/1/20 寫暫存
         'T01 : 發票號碼
         'T02 : 客戶編號
         'T15 : 收據抬頭
         'Modified by Morgan 2015/4/21 收據抬頭取消Trim,否則遇有造字可能會錯
         adoTaie.Execute "insert into ACCTMP08(T01,T02,T05,T14,T15) values('" & adoacc0k0.Fields("a4301") & "','" & Trim("" & adoacc0k0.Fields("a0k03")) & "'," & _
                                                                          "'" & Me.Name & "','" & strUserNum & "','" & adoacc0k0.Fields("a0k04") & "')"
         '2015/1/20 END
         If bolEInvoice = False Then PUB_PrintCaseReceipt_Inv adoacc0k0, txtAdd, , False
         strNo = adoacc0k0.Fields("a4301").Value
      End If
      adoacc0k0.MoveNext
   Loop
   adoacc0k0.Close
   If bolEInvoice = False Then Printer.EndDoc
   If bolEInvoice = False Then Printer.Font = "新細明體"
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   Exit Function
   
ErrHnd:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Function
