VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24j0 
   AutoRedraw      =   -1  'True
   Caption         =   "國外應收帳款分析表"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1245
   ScaleWidth      =   5160
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   180
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   720
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
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
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   90
      Top             =   150
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收款日起迄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   255
      TabIndex        =   4
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc24j0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit
Const SDFormat As String = "###/##"
Const SDFormatCheck As String = "___/__"

Dim lngDate(8) As Long
Dim iRowHeight As Integer
Dim iTopMargin As Integer
Dim iLeftMargin As Integer
Dim iTBWidth As Integer
Dim iTBHeight As Integer
Dim iColWidth(0 To 1) As Integer

Private Sub Command2_Click()
   Screen.MousePointer = vbHourglass
   If FormCheck = True Then
      ProduceData
   End If
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 1650
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = SDFormat
   MaskEdBox2.Mask = SDFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc24j0 = Nothing
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()

   Dim stVTB0 As String, stVTB1 As String, stSQL As String
   Dim stDate1 As String, stDate2 As String, ii As Integer
   Const cFCT As String = "@CFT@CFC@FCT@S@T@TB@TC@TD@TF@TM@TR@TS@TT@"
   Const cFCP As String = "@P@PS@FCP@FG@CFP@CPS@"
   Const cFCL As String = "@FCL@L@CFL@LA@"
   
   lngDate(0) = Val(Replace(MaskEdBox2, "/", "") & "99") '收款迄日 <
   lngDate(1) = Val(Replace(MaskEdBox1, "/", "") & "00") '收款起日 >
   If lngDate(1) \ 10000 = 100 Then
      lngDate(2) = Val(lngDate(1) \ 10000 - 2 & "1299") '前1年 >
   Else
      lngDate(2) = Val(lngDate(1) \ 10000 - 1 & "1299") '當年前幾月 >
   End If
   lngDate(3) = Val(lngDate(2) \ 10000 - 1 & "1299")
   lngDate(4) = Val(lngDate(3) \ 10000 - 1 & "1299")
   lngDate(5) = Val(lngDate(4) \ 10000 - 1 & "1299")
   lngDate(6) = Val(lngDate(5) \ 10000 - 1 & "1299")
   lngDate(7) = Val(lngDate(6) \ 10000 - 5 & "1299")
   
   '先不考慮部分收款--婧瑄
   '請款資料[未作廢 & 未銷帳 & (未結清 or 當期結清)]
   '已結清的都是當期收款
   stVTB0 = "SELECT A1K01 FROM ACC1K0" & _
      " Where nvl(a1k12, 0) = 0 And a1k25 Is Null And a1k29 Is Null And nvl(a1k30,0)=0" & _
      " Union select A1K01 from acc0y0,acc0z0,acc1k0" & _
      " Where a0y02>" & lngDate(1) & " and a0y02<" & lngDate(0) & " and a0z01(+)=a0y01 and a1k01(+)=a0z02 and a1k29='Y'" & _
      " Union select a1k01 from acc1p0, acc1k0 a" & _
      " where a1p18>" & lngDate(1) & " and a1p18<" & lngDate(0) & " and substr(a1p04,1,1)='Z' and a1k17=a1p04 and a1k29='Y'"
   
   '本月新增應收款資料-當期
   stSQL = "SELECT" & _
      "  SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K08))) FCT_US_XREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K11))) FCT_NT_XREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K08))) FCP_US_XREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K11))) FCP_NT_XREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K08))) FCL_US_XREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K11))) FCL_NT_XREC_1" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K08)) US_XREC_1" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(1) & "),1,A1K11)) NT_XREC_1"
   
   '本月收款資料-當期
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K08)))) FCT_US_REC_1" & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K30)))) FCT_NT_REC_1" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K08)))) FCP_US_REC_1" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K30)))) FCP_NT_REC_1" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K08)))) FCL_US_REC_1" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K30)))) FCL_NT_REC_1" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K08))) US_REC_1" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,'Y',A1K30))) NT_REC_1"
      
   '本月應收帳款資料-當期
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K08)))) FCT_US_UREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K11)))) FCT_NT_UREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K08)))) FCP_US_UREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K11)))) FCP_NT_UREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K08)))) FCL_US_UREC_1" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K11)))) FCL_NT_UREC_1" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K08))) US_UREC_1" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(1) & "),1,DECODE(a1k29,NULL,A1K11))) NT_UREC_1"
      
   For ii = 2 To 7
   '上月應收款資料-中間
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K08)))) FCT_US_XREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K11)))) FCT_NT_XREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K08)))) FCP_US_XREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K11)))) FCP_NT_XREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K08)))) FCL_US_XREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K11)))) FCL_NT_XREC_" & ii & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K08))) US_XREC_" & ii & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,A1K11))) NT_XREC_" & ii
   
   '本月收款資料-中間
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K08))))) FCT_US_REC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K30))))) FCT_NT_REC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K08))))) FCP_US_REC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K30))))) FCP_NT_REC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K08))))) FCL_US_REC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K30))))) FCL_NT_REC_" & ii & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K08)))) US_REC_" & ii & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,'Y',A1K30)))) NT_REC_" & ii
   
   '本月應收帳款資料-中間
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K08))))) FCT_US_UREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K11))))) FCT_NT_UREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K08))))) FCP_US_UREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K11))))) FCP_NT_UREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K08))))) FCL_US_UREC_" & ii & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K11))))) FCL_NT_UREC_" & ii & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K08)))) US_UREC_" & ii & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(ii) & "),1,DECODE(SIGN(A1K02-" & lngDate(ii - 1) & "),-1,DECODE(a1k29,NULL,A1K11)))) NT_UREC_" & ii
   
   Next
   
   '上月應收款資料-早期
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K08))) FCT_US_XREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K11))) FCT_NT_XREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K08))) FCP_US_XREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K11))) FCP_NT_XREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K08))) FCL_US_XREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K11))) FCL_NT_XREC_8" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K08)) US_XREC_8" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,A1K11)) NT_XREC_8"
   
   '本月收款資料-早期
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K08)))) FCT_US_REC_8" & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K30)))) FCT_NT_REC_8" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K08)))) FCP_US_REC_8" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K30)))) FCP_NT_REC_8" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K08)))) FCL_US_REC_8" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K30)))) FCL_NT_REC_8" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K08))) US_REC_8" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,'Y',A1K30))) NT_REC_8"
      
   '本月應收帳款資料-早期
   stSQL = stSQL & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K08)))) FCT_US_UREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCT & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K11)))) FCT_NT_UREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K08)))) FCP_US_UREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCP & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K11)))) FCP_NT_UREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K08)))) FCL_US_UREC_8" & _
      ", SUM(DECODE(INSTR('" & cFCL & "','@'||A1K13||'@'),0,0,DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K11)))) FCL_NT_UREC_8" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K08))) US_UREC_8" & _
      ", SUM(DECODE(SIGN(A1K02-" & lngDate(7) & "),-1,DECODE(a1k29,NULL,A1K11))) NT_UREC_8"
      
   stSQL = stSQL & _
      " FROM (" & stVTB0 & ") X, ACC1K0 Y" & _
      " WHERE Y.A1K01(+)=X.A1K01"
      
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, stSQL)
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      DoPrint RsTemp
   End If

End Sub

Private Sub DoPrint(ByRef p_Rst As ADODB.Recordset)
   Dim strTmp As String, ii As Integer, iRow As Integer
   
   iRowHeight = 7
   iTopMargin = 7
   iLeftMargin = 3
   iColWidth(0) = 61
   iColWidth(1) = 16
   iTBWidth = iColWidth(0) + iColWidth(1) * 8
   iTBHeight = iRowHeight * 39
   
   Printer.PaperSize = vbPRPSA4 '9
   Printer.ScaleMode = vbMillimeters '6 公厘
   Printer.Orientation = vbPRORPortrait '1 直印
   
   printTable
   
   Printer.FontBold = True
   Printer.FontSize = 16
   Printer.CurrentY = iTopMargin + 2.5
   Printer.CurrentX = iLeftMargin + iTBWidth / 2 - Printer.TextWidth(Me.Caption) / 2
   Printer.Print Me.Caption
   
   Printer.FontBold = False
   Printer.FontSize = 12
   strExc(0) = Val(Left(MaskEdBox1.Text, 3)) & " 年 "
   If MaskEdBox1.Text <> MaskEdBox2.Text Then
      strExc(0) = strExc(0) & Val(Right(MaskEdBox1.Text, 2)) & " ～ " & Val(Right(MaskEdBox2.Text, 2)) & " 月"
   Else
      strExc(0) = strExc(0) & Val(Right(MaskEdBox1.Text, 2)) & " 月"
   End If
   
   iRow = 2
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iTBWidth / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "帳款年度"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = "~ " & lngDate(7) \ 10000
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = lngDate(7) \ 10000 + 1 & "-" & lngDate(6) \ 10000
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 1 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = lngDate(5) \ 10000
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 2 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = lngDate(4) \ 10000
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 3 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = lngDate(3) \ 10000
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 4 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = lngDate(2) \ 10000
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 5 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   If lngDate(1) Mod 10000 = 100 Then
      strExc(0) = lngDate(1) \ 10000 - 1
   Else
      strExc(0) = lngDate(1) \ 10000 & "/1-" & (lngDate(1) \ 100 Mod 100 - 1)
   End If
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 6 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   If lngDate(1) \ 100 = lngDate(0) \ 100 Then
      strExc(0) = lngDate(1) \ 10000 & "/" & (lngDate(1) \ 100 Mod 100)
   Else
      strExc(0) = lngDate(1) \ 10000 & "/" & (lngDate(1) \ 100 Mod 100) & "-" & ((lngDate(0) \ 100 Mod 100))
   End If
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 7 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   iLeftMargin = iLeftMargin + 1 '右移1厘米
   '上月應收
   iRow = iRow + 1
   strExc(0) = "上月FCT應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "上月FCT應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "上月FCP應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "上月FCP應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "上月FCL應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "上月FCL應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "上月國外應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "上月國外應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   
   '本月新增應收帳款
   iRow = iRow + 2
   strExc(0) = "本月新增FCT應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月新增FCT應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月新增FCP應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月新增FCP應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月新增FCL應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月新增FCL應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月新增國外應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月新增國外應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 2
   '本月收款
   strExc(0) = "本月FCT收款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCT收款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCP收款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCP收款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCL收款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCL收款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月國外收款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月國外收款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   '本月應收帳款
   iRow = iRow + 2
   strExc(0) = "本月FCT應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCT應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCP應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCP應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCL應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月FCL應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月國外應收帳款美金合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
   
   iRow = iRow + 1
   strExc(0) = "本月國外應收帳款台幣合計"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow
   Printer.CurrentX = iLeftMargin
   Printer.Print strExc(0)
      
   Printer.FontSize = 9
   iTopMargin = iTopMargin + 1 '下移1厘米
   iLeftMargin = iLeftMargin - 2 '左移1厘米
   With p_Rst
      '上月應收
      iRow = 4
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("FCT_US_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("FCT_NT_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("FCP_US_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("FCP_NT_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("FCL_US_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("FCL_NT_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("US_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 2 Step -1
         strExc(0) = Format(Val("" & .Fields("NT_XREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 2
      '本月新增應收帳款
      strExc(0) = Format(Val("" & .Fields("FCT_US_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      iRow = iRow + 1
      strExc(0) = Format(Val("" & .Fields("FCT_NT_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      iRow = iRow + 1
      strExc(0) = Format(Val("" & .Fields("FCP_US_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      iRow = iRow + 1
      strExc(0) = Format(Val("" & .Fields("FCP_NT_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      iRow = iRow + 1
      strExc(0) = Format(Val("" & .Fields("FCL_US_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      iRow = iRow + 1
      strExc(0) = Format(Val("" & .Fields("FCL_NT_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      iRow = iRow + 1
      strExc(0) = Format(Val("" & .Fields("US_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      iRow = iRow + 1
      strExc(0) = Format(Val("" & .Fields("NT_XREC_1")), DDollar)
      Printer.CurrentY = iTopMargin + iRowHeight * iRow
      Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 8 - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
            
      '本月收款
      iRow = iRow + 2
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCT_US_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCT_NT_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCP_US_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCP_NT_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCL_US_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCL_NT_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("US_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("NT_REC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      '本月應收帳款
      iRow = iRow + 2
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCT_US_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCT_NT_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCP_US_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCP_NT_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCL_US_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("FCL_NT_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("US_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
      iRow = iRow + 1
      For ii = 8 To 1 Step -1
         strExc(0) = Format(Val("" & .Fields("NT_UREC_" & ii)), DDollar)
         Printer.CurrentY = iTopMargin + iRowHeight * iRow
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * (9 - ii) - Printer.TextWidth(strExc(0))
         Printer.Print strExc(0)
      Next
      
   End With
   
   Printer.EndDoc
End Sub

Private Sub printTable()
   Dim ii As Integer
   
   Printer.DrawWidth = 6
   Printer.Line (iLeftMargin, iTopMargin - 1)-(iLeftMargin + iTBWidth, iTopMargin - 1)
   
   '橫線
   Printer.DrawWidth = 3
   For ii = 2 To 38
      Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * ii)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * ii)
   Next
   
   Printer.DrawWidth = 6
   Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * ii)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * ii)
   
   '直線
   Printer.DrawWidth = 6
   Printer.Line (iLeftMargin, iTopMargin - 1)-(iLeftMargin, iTopMargin - 1 + iRowHeight * 39)
   
   Printer.DrawWidth = 3
   Printer.Line (iLeftMargin + iColWidth(0), iTopMargin - 1 + iRowHeight * 3)-(iLeftMargin + iColWidth(0), iTopMargin - 1 + iRowHeight * 39)
   For ii = 1 To 7
      Printer.Line (iLeftMargin + iColWidth(0) + iColWidth(1) * ii, iTopMargin - 1 + iRowHeight * 3)-(iLeftMargin + iColWidth(0) + iColWidth(1) * ii, iTopMargin - 1 + iRowHeight * 39)
   Next
   
   Printer.DrawWidth = 6
   Printer.Line (iLeftMargin + iColWidth(0) + iColWidth(1) * 8, iTopMargin - 1)-(iLeftMargin + iColWidth(0) + iColWidth(1) * 8, iTopMargin - 1 + iRowHeight * 39)
   
End Sub
'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = SDFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = SDFormat
   MaskEdBox1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If MaskEdBox1.Text = SDFormatCheck Then
      MsgBox "請輸入收款區間起！"
      Exit Function
   End If
   If MaskEdBox2.Text = SDFormatCheck Then
      MsgBox "請輸入收款區間迄！"
      Exit Function
   End If
   If IsDate(MaskEdBox1.Text) = False Then
      MsgBox "收款區間起格式錯誤！"
      Exit Function
   End If
   If IsDate(MaskEdBox1.Text) = False Then
      MsgBox "收款區間迄格式錯誤！"
      Exit Function
   End If
   If MaskEdBox2.Text < MaskEdBox1.Text Then
      MsgBox "收款區間起迄錯誤！"
      Exit Function
   End If
   FormCheck = True
   
End Function

