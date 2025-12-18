VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050325_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "美國發明退公開費案件查詢"
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
      Caption         =   "列印(&P)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   45
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   30
      Width           =   756
   End
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
      Left            =   7290
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
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "代理人　　　　　　　　　|本所案號　　　　|智權人員　　　|收據抬頭　　　　　　|公告日　　|公開日　　"
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
End
Attribute VB_Name = "frm050325_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Create by Morgan 2007/12/27
Option Explicit

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         frm050325.Show
         Unload Me
      Case 1
         Unload frm050325
         Unload Me
      Case 2
         Set RsTemp = grdDataList.Recordset.Clone
         RsTemp.Sort = "Srt,cp44,pa01,pa02,pa03,pa04"
         frm050325.DoPrint RsTemp
   End Select
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050325_1 = Nothing
End Sub

