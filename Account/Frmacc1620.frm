VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1620 
   AutoRedraw      =   -1  'True
   Caption         =   "補開請款單列印"
   ClientHeight    =   5916
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7392
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5916
   ScaleWidth      =   7392
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
      Left            =   7710
      TabIndex        =   36
      Top             =   240
      Width           =   3345
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   390
      TabIndex        =   30
      Top             =   4440
      Width           =   4755
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
         TabIndex        =   32
         Top             =   120
         Width           =   705
      End
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
         TabIndex        =   31
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label16 
         BackStyle       =   0  '透明
         Caption         =   "所別："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   35
         Top             =   180
         Width           =   675
      End
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
      TabIndex        =   23
      Text            =   "Y"
      Top             =   2070
      Width           =   585
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   8070
      Style           =   2  '單純下拉式
      TabIndex        =   22
      Top             =   1410
      Width           =   3450
   End
   Begin VB.CheckBox chk 
      Caption         =   "列印客戶案件案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   840
      TabIndex        =   19
      Top             =   1440
      Value           =   1  '核取
      Width           =   2235
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1230
      Style           =   2  '單純下拉式
      TabIndex        =   18
      Top             =   960
      Width           =   3450
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   8070
      Style           =   2  '單純下拉式
      TabIndex        =   17
      Top             =   870
      Width           =   3450
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   405
      Left            =   840
      TabIndex        =   15
      Top             =   210
      Width           =   3405
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
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Value           =   -1  'True
         Width           =   3555
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1830
      TabIndex        =   13
      Top             =   3330
      Width           =   1815
   End
   Begin VB.TextBox Text4 
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
      Height          =   315
      Left            =   5820
      TabIndex        =   5
      Top             =   3330
      Width           =   735
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
      Left            =   1290
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   5310
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   4980
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4980
      TabIndex        =   1
      Top             =   2550
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   14737632
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
   Begin VB.TextBox Text1 
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
      Left            =   1830
      TabIndex        =   0
      Top             =   2550
      Width           =   1845
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1830
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   1830
      TabIndex        =   2
      Top             =   2940
      Width           =   4725
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "發票地址條若要印出                      請至財務Mail資料維護勾選(寄紙本)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   7680
      TabIndex        =   29
      Top             =   2400
      Width           =   3300
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label14 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "「與正本相符」"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11190
      TabIndex        =   28
      Top             =   2910
      Width           =   1470
   End
   Begin VB.Label Label13 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "的發票底稿，再列印本發票內容！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   11190
      TabIndex        =   27
      Top             =   3165
      Width           =   1395
   End
   Begin VB.Label Label12 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "PS：請先列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   11190
      TabIndex        =   26
      Top             =   2670
      Width           =   1290
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "列印地址條：            (Y : 是)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7710
      TabIndex        =   25
      Top             =   2100
      Width           =   2805
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   1200
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1950
      Left            =   390
      Top             =   2340
      Width           =   6630
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "請款單印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "發票印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   20
      Top             =   630
      Width           =   2205
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "統一編號"
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
      Left            =   870
      TabIndex        =   14
      Top             =   3330
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "是否補開(Y/N)?"
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
      Left            =   4140
      TabIndex        =   12
      Top             =   3330
      Width           =   1740
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "請款單抬頭"
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
      Left            =   660
      TabIndex        =   11
      Top             =   2940
      Width           =   1185
   End
   Begin VB.Label Label4 
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
      Left            =   4020
      TabIndex        =   10
      Top             =   2580
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "列印次數"
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
      Left            =   870
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   4020
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
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
      Left            =   870
      TabIndex        =   7
      Top             =   2580
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1620"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/01 改成Form2.0 ; 地址條改成Excel列印
'Memo By Sindy 2022/1/17 Form2.0已修改
'Create By Sindy 2014/1/7
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Dim strPrinter As String
Dim prnPrint As Printer
Public ProState As String '權限: 1.全所 2.該所
Dim m_FileName As String 'Add By Sindy 2021/12/29


'Private Sub Command1_Click()
'Dim bolChk As Boolean, sqlST06 As String, ii As Integer
'
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'
'   'Add By Sindy 2021/5/21
'   bolChk = False
'   sqlST06 = " and ST06 in("
'   For ii = 0 To 3
'      If ChkST06(ii).Value = 1 Then
'         bolChk = True
'         sqlST06 = sqlST06 & "'" & ii + 1 & "'"
'      End If
'   Next ii
'   If bolChk = False Then
'      MsgBox "請勾選所別！", vbExclamation
'      ChkST06(0).SetFocus
'      sqlST06 = ""
'      Exit Sub
'   End If
'   sqlST06 = sqlST06 & ")"
'   sqlST06 = Replace(sqlST06, "''", "','")
'   '2021/5/21 END
'
'   Screen.MousePointer = vbHourglass
'
'   If Option2(0).Value = True Then '請款單
'      adoacc0k0.CursorLocation = adUseClient
'      'Modify By Sindy 2020/11/12 + ,acc0j0,caseprogress ; and a0j13=a0k01 and a0j01=cp09
'      adoacc0k0.Open "select * from acc0k0,customer,acc0j0,caseprogress,staff where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and a0k11='J' and a0k01 = '" & Text1 & "' and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
'                     " and (a0k09 is null or a0k09 = 0)" & _
'                     " and a0k01 not in (select a0m02 from acc0m0 where a0m02 = '" & Text1 & "' and a0m03 is not null)" & _
'                     " and a0j13=a0k01 and a0j01=cp09 and a0k20=st01(+)" & sqlST06, adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc0k0.RecordCount = 0 Then
'         adoacc0k0.Close
'         adoacc0k0.Open "select * from acc0k0,customer,acc0j0,caseprogress,staff where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and a0k11='J' and a0k01 = '" & Text1 & "' and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
'                     " and (a0k09 is null or a0k09 = 0)" & _
'                     " and a0j13=a0k01 and a0j01=cp09 and a0k20=st01(+)" & sqlST06, adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc0k0.RecordCount = 0 Then
'            adoacc0k0.Close
'            Screen.MousePointer = vbDefault
'            MsgBox MsgText(28), , MsgText(5)
'            Exit Sub
'         Else
'            adoacc0k0.Close
'            Screen.MousePointer = vbDefault
'            MsgBox MsgText(214), , MsgText(5)
'            Exit Sub
'         End If
'      '控制已列印才可補開 and a0k19>0
'      ElseIf Val("" & adoacc0k0("a0k19")) = 0 Then
'         adoacc0k0.Close
'         Screen.MousePointer = vbDefault
'         MsgBox "請款單尚未列印不可補開", , MsgText(5)
'         Exit Sub
'      End If
'
'      If IsNull(adoacc0k0.Fields("a0k11").Value) = False Then
'         If MsgBox("按確定開始列印請款單...", vbOKCancel + vbDefaultButton2, MsgText(5)) = vbCancel Then
'            adoacc0k0.Close
'            Screen.MousePointer = vbDefault
'            Exit Sub
'         End If
'      End If
'
'      PUB_RestorePrinter Combo1
'      'XP自定紙張需手動設定並將印表機預設為該紙張
'      '9x
''      If pub_OS = "1" Then
''         Printer.Height = 8750
''         Printer.Width = 13000
''      Else
''         Printer.PaperSize = PUB_GetPaperSize(1)
''      End If
'      Printer.FontSize = 12
'      Printer.Font = "標楷體"
'
'      Do While adoacc0k0.EOF = False
'         '改呼叫共用(與補印請款單統一)
'         PUB_PrintCaseReceipt_J adoacc0k0, 0, Me.chk(1), 0, Text4, "", "", FCDate(MaskEdBox2.Text)
'         adoacc0k0.MoveNext
'      Loop
'      adoacc0k0.Close
'      Printer.EndDoc
'      PUB_RestorePrinter strPrinter
'
'   Else '發票
'      adoacc0k0.CursorLocation = adUseClient
'      'Modify By Sindy 2016/1/18 +,a0k33
'      'Modify By Sindy 2020/11/12 + " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')"
'      strSql = "select acc430.*,axc02,a0k20,st02,st15,a0k03,a0k04,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as cuname,a0k33" & _
'               " from acc430,acc431,acc0k0,customer,staff" & _
'               " where a4301=axc01(+)" & _
'               " and axc02=a0k01(+)" & _
'               " and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)" & _
'               " and a0k20=st01(+)" & sqlST06 & _
'               " and a4301 = '" & Text1 & "' and (a4308 is null or a4308 = 0)" & _
'               " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')"
'      adoacc0k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc0k0.RecordCount = 0 Then
'         adoacc0k0.Close
'         Screen.MousePointer = vbDefault
'         MsgBox MsgText(28), , MsgText(5)
'         Exit Sub
'      End If
'
'      '控制已列印才可補開 and a0k19>0
'      If Val("" & adoacc0k0("a4306")) = 0 Then
'         adoacc0k0.Close
'         Screen.MousePointer = vbDefault
'         MsgBox "發票尚未列印不可補開", , MsgText(5)
'         Exit Sub
'      End If
'
'      'Modify By Sindy 2019/7/2
'      If Val(FCDate(MaskEdBox1.Text)) < Val(1080701) Then '紙本發票
'      '2019/7/2 END
'         If MsgBox("請放置發票套表紙於選取的印表機，按確定開始列印...", vbOKCancel + vbDefaultButton2, MsgText(5)) = vbCancel Then
'            adoacc0k0.Close
'            Screen.MousePointer = vbDefault
'            Exit Sub
'         End If
'
'         PUB_RestorePrinter Combo2
'         'XP自定紙張需手動設定並將印表機預設為該紙張
'         '9x
'   '      If pub_OS = "1" Then
'   '         Printer.Height = 8750
'   '         Printer.Width = 13000
'   '      Else
'   '         Printer.PaperSize = PUB_GetPaperSize(1)
'   '      End If
'         Printer.FontSize = 12
'         Printer.Font = "標楷體"
'
'         Do While adoacc0k0.EOF = False
'            '改呼叫共用(與補印發票統一)
'            PUB_PrintCaseReceipt_Inv adoacc0k0, "", Text4
'            adoacc0k0.MoveNext
'         Loop
'         Printer.EndDoc
'      End If
'
'      'Add By Sindy 2016/12/13 + 列印地址條
'      '若是否列印地址條上"Y"
'      If Me.txtAdd.Text = "Y" Then
'         If MsgBox("是否要列印地址條？" & vbCrLf & _
'                   "若要印，請放地址條貼紙於選取的印表機!!", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
'            PUB_SetOsDefaultPrinter Combo3 'Added by Lydia 2022/03/01 切換Word/Excel印表機
'            PUB_RestorePrinter Combo3
'            PrintAddress '列印地址條
'         End If
'      End If
'      '2016/12/13 END
'
'      adoacc0k0.Close
'      PUB_SetOsDefaultPrinter strPrinter  'Added by Lydia 2022/03/01 切換Word/Excel印表機
'      PUB_RestorePrinter strPrinter
'   End If
'
'   'Printer.Font = "新細明體"
'   Screen.MousePointer = vbDefault
'   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
'End Sub

'Add By Sindy 2016/12/13
'*************************************************
'  列印地址條
'
'*************************************************
Private Sub PrintAddress()
   Dim intCounter As Integer
   Dim strAddr As String, strCustName As String
   Dim intR As Integer
   Dim strA0K04 As String
   Dim intLine As Integer
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'   'Add by Amy 2019/07/25
'   Dim intFCount As Integer '印地址條筆數
'   Dim strCU179 As String '電子發票寄送方式
   Dim strTempAddressList As String 'Added by Lydia 2022/03/01
   
   strA0K04 = ""
   With adoacc0k0
      .MoveFirst
      'XP自定紙張需手動設定並將印表機預設為該紙張
      '9x
      'Remove by Lydia 2022/03/01 改用Excel列印
      'If pub_OS = "1" Then
      '   Printer.Height = 2880
      '   Printer.Width = 13000
      'Else
      '   Printer.PaperSize = PUB_GetPaperSize(2)
      'End If
      '
      'Printer.Font = "@新細明體"
      'Printer.FontSize = 12
      'end 2022/03/01
      Do While .EOF = False
         intCounter = 0: intLine = 0
         If strA0K04 <> .Fields("a0k04") Then
            strAddr = "" '地址
            strCustName = Trim("" & .Fields("a0k04").Value) '客戶名稱
            'Mark by Amy 2019/07/25 改寫至GetInvoiceAddress
'            '收據抬頭為3個字以下(含3個字)以a0k03為主
'            If Len(Trim("" & .Fields("a0k04").Value)) <= 3 Then
'               strSql = "select cu01,cu02,cu04,cu112,cu23,cu30,cu31 from customer where cu01='" & Left(Trim("" & .Fields("a0k03").Value), 8) & "' and cu02='" & Mid(Trim("" & .Fields("a0k03").Value), 9, 1) & "'"
'               intR = 1
'               Set RsTemp = ClsLawReadRstMsg(intR, strSql)
'               If intR = 1 Then
'                  '聯絡地址優先
'                  If Trim("" & RsTemp.Fields("cu31")) <> "" Then
'                     strAddr = Trim("" & RsTemp.Fields("cu30")) & " " & Trim("" & RsTemp.Fields("cu31"))
'                  '再來中文地址
'                  ElseIf Trim("" & RsTemp.Fields("cu23")) <> "" Then
'                     strAddr = Trim("" & RsTemp.Fields("cu112")) & " " & Trim("" & RsTemp.Fields("cu23"))
'                  End If
'               End If
'            Else
'               '以收據抬頭抓客戶檔之CU04,若為客戶資料則抓中文地址
'               '                         若不存在,則再抓收據抬頭資料檔acc420的營業地址
'               '收據抬頭取消Trim,否則遇有造字可能會錯
'               strSql = "select cu01,cu02,cu04,cu112,cu23,cu30,cu31 from customer where cu04='" & .Fields("a0k04").Value & "' and (cu80 is null or cu80='其他' or cu80='業務自行處理') and cu02=0"
'               intR = 1
'               Set RsTemp = ClsLawReadRstMsg(intR, strSql)
'               If intR = 1 Then
'                  '聯絡地址優先
'                  If Trim("" & RsTemp.Fields("cu31")) <> "" Then
'                     strAddr = Trim("" & RsTemp.Fields("cu30")) & " " & Trim("" & RsTemp.Fields("cu31"))
'                  '再來中文地址
'                  ElseIf Trim("" & RsTemp.Fields("cu23")) <> "" Then
'                     strAddr = Trim("" & RsTemp.Fields("cu112")) & " " & Trim("" & RsTemp.Fields("cu23"))
'                  End If
'               Else
'                  '收據抬頭取消Trim,否則遇有造字可能會錯
'                  strSql = "select * from acc420 where a4201='" & .Fields("a0k04").Value & "'"
'                  intR = 1
'                  Set RsTemp = ClsLawReadRstMsg(intR, strSql)
'                  If intR = 1 Then
'                     '郵寄地址優先
'                     If Trim("" & RsTemp.Fields("a4203")) <> "" Then
'                        strAddr = Trim("" & RsTemp.Fields("a4203"))
'                     '再來營業地址
'                     ElseIf Trim("" & RsTemp.Fields("a4215")) <> "" Then
'                        strAddr = Trim("" & RsTemp.Fields("a4215"))
'                     End If
'                  End If
'               End If
'            End If
            'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
            'strAddr = GetInvoiceAddress(Trim("" & .Fields("a0k03").Value), "" & .Fields("a0k04").Value, False, strCU179)
            'end 2019/07/25
            'Modify by Amy 2019/07/25 電子發票寄送方式為 1.紙本才印地址條
'            If strCU179 = "1" Then
'                intFCount = intFCount + 1
'                'Modified by Lydia 2022/03/01 傳入多張地址條的內容；用|區隔不同張地址條，同一張地址條用$區隔地址和收件人
'                'Call PUB_PrintAccAddress(strAddr, strCustName)
'                If strAddr & strCustName <> "" Then strTempAddressList = strTempAddressList & Trim(strAddr) & "$" & Trim(strCustName) & "|"
'            End If
            'end 2024/06/13
         End If
         strA0K04 = .Fields("a0k04")
         .MoveNext
      Loop
   End With
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

'Add By Sindy 2021/12/28
Private Sub Command1_Click()
Dim bolChk As Boolean, sqlST06 As String, ii As Integer

   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   'Add By Sindy 2021/5/21
   bolChk = False
   sqlST06 = " and ST06 in("
   For ii = 0 To 3
      If ChkST06(ii).Value = 1 Then
         bolChk = True
         sqlST06 = sqlST06 & "'" & ii + 1 & "'"
      End If
   Next ii
   If bolChk = False Then
      MsgBox "請勾選所別！", vbExclamation
      ChkST06(0).SetFocus
      sqlST06 = ""
      Exit Sub
   End If
   sqlST06 = sqlST06 & ")"
   sqlST06 = Replace(sqlST06, "''", "','")
   '2021/5/21 END
   
   Screen.MousePointer = vbHourglass
   
   If Option2(0).Value = True Then '請款單
      adoacc0k0.CursorLocation = adUseClient
      'Modify By Sindy 2020/11/12 + ,acc0j0,caseprogress ; and a0j13=a0k01 and a0j01=cp09
      '==> 2022/1/17 Mark
      ',acc0j0,caseprogress
      '& " and a0j13=a0k01 and a0j01=cp09"
      'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據=> + and nvl(a0k32,'Y') <> 'Z'
      adoacc0k0.Open "select * from acc0k0,customer,staff where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and a0k11='J' and a0k01 = '" & Text1 & "' and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
                     " and (a0k09 is null or a0k09 = 0) and nvl(a0k32,'Y') <> 'Z'" & _
                     " and a0k01 not in (select a0m02 from acc0m0 where a0m02 = '" & Text1 & "' and a0m03 is not null)" & _
                     " and a0k20=st01(+)" & sqlST06, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0k0.RecordCount = 0 Then
         adoacc0k0.Close
         adoacc0k0.Open "select * from acc0k0,customer,acc0j0,caseprogress,staff where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and a0k11='J' and a0k01 = '" & Text1 & "' and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
                     " and (a0k09 is null or a0k09 = 0)" & _
                     " and a0j13=a0k01 and a0j01=cp09 and a0k20=st01(+)" & sqlST06, adoTaie, adOpenStatic, adLockReadOnly
         If adoacc0k0.RecordCount = 0 Then
            adoacc0k0.Close
            Screen.MousePointer = vbDefault
            MsgBox MsgText(28), , MsgText(5)
            Exit Sub
         Else
            adoacc0k0.Close
            Screen.MousePointer = vbDefault
            MsgBox MsgText(214), , MsgText(5)
            Exit Sub
         End If
      '控制已列印才可補開 and a0k19>0
      ElseIf Val("" & adoacc0k0("a0k19")) = 0 Then
         adoacc0k0.Close
         Screen.MousePointer = vbDefault
         MsgBox "請款單尚未列印不可補開", , MsgText(5)
         Exit Sub
      End If
      
      If IsNull(adoacc0k0.Fields("a0k11").Value) = False Then
         If MsgBox("按確定開始列印請款單...", vbOKCancel + vbDefaultButton2, MsgText(5)) = vbCancel Then
            adoacc0k0.Close
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
      
      '切換印表機
      PUB_SetOsDefaultPrinter Combo1
      PUB_RestorePrinter Combo1
      Do While adoacc0k0.EOF = False
         '改呼叫共用(與補印請款單統一)
         PUB_PrintCaseReceipt_J_Doc App.path & "\" & strUserNum & "\" & m_FileName, adoacc0k0, 0, Me.chk(1), 0, Text4, "", "", FCDate(MaskEdBox2.Text)
         adoacc0k0.MoveNext
      Loop
      adoacc0k0.Close
      '還原印表機
      PUB_SetOsDefaultPrinter strPrinter
      PUB_RestorePrinter strPrinter
      
      '舊請款單列印:
'      PUB_RestorePrinter Combo1
'      'XP自定紙張需手動設定並將印表機預設為該紙張
'      '9x
''      If pub_OS = "1" Then
''         Printer.Height = 8750
''         Printer.Width = 13000
''      Else
''         Printer.PaperSize = PUB_GetPaperSize(1)
''      End If
'      Printer.FontSize = 12
'      Printer.Font = "標楷體"
'
'      Do While adoacc0k0.EOF = False
'         '改呼叫共用(與補印請款單統一)
'         PUB_PrintCaseReceipt_J adoacc0k0, 0, Me.chk(1), 0, Text4, "", "", FCDate(MaskEdBox2.Text)
'         adoacc0k0.MoveNext
'      Loop
'      adoacc0k0.Close
'      Printer.EndDoc
'      PUB_RestorePrinter strPrinter
      
   Else '發票
      adoacc0k0.CursorLocation = adUseClient
      'Modify By Sindy 2016/1/18 +,a0k33
      'Modify By Sindy 2020/11/12 + " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')"
      strSql = "select acc430.*,axc02,a0k20,st02,st15,a0k03,a0k04,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as cuname,a0k33" & _
               " from acc430,acc431,acc0k0,customer,staff" & _
               " where a4301=axc01(+)" & _
               " and axc02=a0k01(+)" & _
               " and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)" & _
               " and a0k20=st01(+)" & sqlST06 & _
               " and a4301 = '" & Text1 & "' and (a4308 is null or a4308 = 0)" & _
               " and not exists(select * from acc0j0,caseprogress where a0j13=a0k01 and a0j01=cp09 and a0j02 like 'ACS%' and a0j03='706')"
      adoacc0k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0k0.RecordCount = 0 Then
         adoacc0k0.Close
         Screen.MousePointer = vbDefault
         MsgBox MsgText(28), , MsgText(5)
         Exit Sub
      End If
      
      '控制已列印才可補開 and a0k19>0
      If Val("" & adoacc0k0("a4306")) = 0 Then
         adoacc0k0.Close
         Screen.MousePointer = vbDefault
         MsgBox "發票尚未列印不可補開", , MsgText(5)
         Exit Sub
      End If
      
      'Modify By Sindy 2019/7/2
      If Val(FCDate(MaskEdBox1.Text)) < Val(1080701) Then '紙本發票
      '2019/7/2 END
         If MsgBox("請放置發票套表紙於選取的印表機，按確定開始列印...", vbOKCancel + vbDefaultButton2, MsgText(5)) = vbCancel Then
            adoacc0k0.Close
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
         PUB_RestorePrinter Combo2
         'XP自定紙張需手動設定並將印表機預設為該紙張
         '9x
   '      If pub_OS = "1" Then
   '         Printer.Height = 8750
   '         Printer.Width = 13000
   '      Else
   '         Printer.PaperSize = PUB_GetPaperSize(1)
   '      End If
         Printer.FontSize = 12
         Printer.Font = "標楷體"
         
         Do While adoacc0k0.EOF = False
            '改呼叫共用(與補印發票統一)
            PUB_PrintCaseReceipt_Inv adoacc0k0, "", Text4
            adoacc0k0.MoveNext
         Loop
         Printer.EndDoc
      End If
      
      'Add By Sindy 2016/12/13 + 列印地址條
      '若是否列印地址條上"Y"
      If Me.txtAdd.Text = "Y" Then
         If MsgBox("是否要列印地址條？" & vbCrLf & _
                   "若要印，請放地址條貼紙於選取的印表機!!", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            PUB_RestorePrinter Combo3
            PrintAddress '列印地址條
         End If
      End If
      '2016/12/13 END
      
      adoacc0k0.Close
      PUB_RestorePrinter strPrinter
   End If
   EndOfficeAp 'Added by Morgan 2025/9/10 印完要清除物件，否則印表機不會變
   'Printer.Font = "新細明體"
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
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
   Me.Width = 7515 '9045
   Me.Height = 6330
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
   
   'Text1 = MsgText(802)
   MaskEdBox1.Mask = DFormat
   'MODIFY by sonia 2014/3/31 列印日期不預設
   'MaskEdBox2.Text = CFDate(ACDate(ServerDate))
   MaskEdBox2.Enabled = False
   MaskEdBox2.Visible = False
   Label2.Visible = False
   '2014/3/31 END
   MaskEdBox2.Mask = DFormat
   Text4 = MsgText(602)
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   PUB_SetPrinter Me.Name, Combo2
   PUB_SetPrinter Me.Name, Combo3 'Add By Sindy 2016/12/13
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
   'Add By Sindy 2016/12/13
   '若印表機變動, 則更新列印設定
   If Me.Combo3.Text <> Me.Combo3.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo3.Name, "0", "0", Me.Combo3.Text
   End If
   '2016/12/13 END
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1620 = Nothing
End Sub

Private Sub Option2_GotFocus(Index As Integer)
   If Option2(Index).Value = False Then
      Text1 = ""
   End If
   If Index = 0 Then
      Text1.MaxLength = 9
   Else
      Text1.MaxLength = 10
   End If
End Sub

Private Sub Text1_Change()
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'   'Add by Amy 2019/07/25
'   Dim strAddr As String
'   Dim strCU179 As String '電子發票寄送方式
   
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text2 = ""
   Text3 = ""
   Text5 = ""
   'add by sonia 2014/3/31
   MaskEdBox2.Enabled = False
   MaskEdBox2.Visible = False
   Label2.Visible = False
   '2014/3/31 END
   If Option2(0).Value = True Then '請款單
      If Len(Text1) < 9 Then Exit Sub
      adoacc0k0.CursorLocation = adUseClient
      adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & Text1 & "' and a0k11='J'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0k0.RecordCount <> 0 Then
         MaskEdBox1.Mask = ""
         If IsNull(adoacc0k0.Fields("a0k02").Value) = False Then
            MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a0k02").Value)
         Else
            MaskEdBox1.Text = ""
         End If
         MaskEdBox1.Mask = DFormat
         If IsNull(adoacc0k0.Fields("a0k19").Value) = False Then
            Text2 = adoacc0k0.Fields("a0k19").Value
         Else
            Text2 = ""
         End If
         If IsNull(adoacc0k0.Fields("a0k04").Value) = False Then
            Text3 = adoacc0k0.Fields("a0k04").Value
         Else
            Text3 = ""
         End If
         'add by sonia 2014/3/31
         MaskEdBox2.Enabled = True
         MaskEdBox2.Visible = True
         MaskEdBox2.Text = CFDate(ACDate(ServerDate))
         Label2.Visible = True
         '2014/3/31 end
      End If
      adoacc0k0.Close
   Else '發票
      If Len(Text1) < 10 Then Exit Sub
      adoacc0k0.CursorLocation = adUseClient
      adoacc0k0.Open "select acc430.*,a0k03,a0k04 from acc430,acc431,acc0k0 where a4301 = '" & Text1 & "' and a4301=axc01(+) and axc02=a0k01(+)", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0k0.RecordCount <> 0 Then
         MaskEdBox1.Mask = ""
         If IsNull(adoacc0k0.Fields("a4302").Value) = False Then
            MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a4302").Value)
         Else
            MaskEdBox1.Text = ""
         End If
         MaskEdBox1.Mask = DFormat
         If IsNull(adoacc0k0.Fields("a4306").Value) = False Then
            Text2 = adoacc0k0.Fields("a4306").Value
         Else
            Text2 = ""
         End If
         If IsNull(adoacc0k0.Fields("a0k04").Value) = False Then
            Text3 = adoacc0k0.Fields("a0k04").Value
         Else
            Text3 = ""
         End If
         If IsNull(adoacc0k0.Fields("a4303").Value) = False Then
            Text5 = adoacc0k0.Fields("a4303").Value
         Else
            Text5 = ""
         End If
         'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'         'Add by Amy 2019/07/25 電子發票寄送方式為電子檔,顯示訊息且不印地址條
'         If Option2(1).Value = True And txtAdd = "Y" And Text1 <> MsgText(601) Then
'            strAddr = GetInvoiceAddress(Trim("" & adoacc0k0.Fields("a0k03").Value), "" & adoacc0k0.Fields("a0k04").Value, False, strCU179)
'            If strCU179 = "2" Then
'                MsgBox Text1 & "電子發票寄送方式為電子檔不會產生地址條！", , MsgText(5)
'                Exit Sub
'            End If
'         End If
         'add by sonia 2018/9/14
         If IsNull(adoacc0k0.Fields("a4309").Value) = False Then
            MsgBox "本發票已有銷退事宜，請確認要補開嗎？", , MsgText(5)
         End If
         'end 2018/9/14
      End If
      adoacc0k0.Close
   End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   'MODIFY by sonia 2014/3/31 列印日期不預設
   'MaskEdBox2.Text = CFDate(ACDate(ServerDate))
   MaskEdBox2.Enabled = False
   MaskEdBox2.Visible = False
   '2014/3/31 END
   MaskEdBox2.Mask = DFormat
   Text2 = ""
   Text3 = ""
   'MODIFY BY SONIA 2014/3/31
   'Text4 = ""
   Text4 = MsgText(602)
   '2014/3/31 END
   Text5 = ""
   Text1.SetFocus
   Me.chk(1).Value = vbUnchecked
   Option2(0).Value = True
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Option2(0).Value = True Then   'ADD BY SONIA 2014/3/31  請款單才檢查
      '列印日期檢查
      If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
         MsgBox "列印日期格式錯誤！", vbExclamation
         If MaskEdBox2.Enabled = True Then MaskEdBox2.SetFocus
         Exit Function
      End If
   End If     '2014/3/31 END
   
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add By Sindy 2016/12/13
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
