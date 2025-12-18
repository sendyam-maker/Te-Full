VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060320_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專未完成核稿明細查詢"
   ClientHeight    =   5730
   ClientLeft      =   165
   ClientTop       =   960
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8544
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7320
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5208
      Left            =   36
      TabIndex        =   2
      Top             =   468
      Width           =   9264
      _ExtentX        =   16351
      _ExtentY        =   9181
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frm060320_1.frx":0000
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
      _Band(0).Cols   =   14
   End
   Begin VB.Label Label1 
      Caption         =   "* 為已延期, ** 為未完稿無核稿期限"
      Height          =   255
      Left            =   45
      TabIndex        =   3
      Top             =   180
      Width           =   3930
   End
End
Attribute VB_Name = "frm060320_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/5/15
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 1
         Unload frm060320
   End Select
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060320_1 = Nothing
End Sub

