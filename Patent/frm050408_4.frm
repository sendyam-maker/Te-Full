VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050408_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人案件統計表(收文明細)"
   ClientHeight    =   5976
   ClientLeft      =   168
   ClientTop       =   960
   ClientWidth     =   6708
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5976
   ScaleWidth      =   6708
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5415
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4425
      Left            =   90
      TabIndex        =   1
      Top             =   990
      Width           =   6510
      _ExtentX        =   11494
      _ExtentY        =   7811
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   4
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblCaseName 
      Height          =   300
      Left            =   990
      TabIndex        =   10
      Top             =   375
      Width           =   4065
      VariousPropertyBits=   27
      Size            =   "7170;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSecurFound 
      Height          =   255
      Left            =   3735
      TabIndex        =   9
      Top             =   630
      Width           =   2025
   End
   Begin VB.Label Label4 
      Caption         =   "PS:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   5490
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "因舊系統帳單無法串連，92/2/1以前收文資料無法計算程序盈虧，但案件盈虧之計算依然正確不受影響。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   450
      TabIndex        =   7
      Top             =   5490
      Width           =   5400
   End
   Begin VB.Label lblNetTot 
      Height          =   255
      Left            =   990
      TabIndex        =   6
      Top             =   630
      Width           =   1530
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件盈虧:"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   630
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   375
      Width           =   765
   End
   Begin VB.Label lblCaseNo 
      Height          =   255
      Left            =   990
      TabIndex        =   3
      Top             =   120
      Width           =   3225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frm050408_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/11 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblCaseName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/6/5
Option Explicit

'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Public Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      .Cols = 5
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 2: .FixedRows = 1: .FixedCols = 0
      End If
      .row = 0
      .RowHeight(0) = 250
      .RowHeightMin = 250
      ii = 0
      .col = ii: .ColWidth(.col) = 1000: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1140: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 2000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1140: .Text = "收文盈虧"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      'add by sonia 2018/4/10
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1000: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      'end 2018/4/10
      .Refresh
      .Visible = True
   End With
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Added by Lydia 2025/06/06
   If frm050408.Tag <> "" Then
      Me.Caption = "互惠期間統計表(收文明細)"
   End If
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050408_4 = Nothing
End Sub
