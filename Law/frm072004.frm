VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm072004 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦/協辦人員案件查詢"
   ClientHeight    =   5820
   ClientLeft      =   195
   ClientTop       =   630
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8352
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Left            =   7224
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   70
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5112
      Left            =   72
      TabIndex        =   2
      Top             =   552
      Width           =   9084
      _ExtentX        =   16034
      _ExtentY        =   9022
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm072004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer

Private Sub cmdEnd_Click()
Unload frm072003
Unload Me
End Sub
Private Sub cmdGoInput_Click()
frm072003.Show
Unload Me

End Sub
Private Sub Form_Load()
MoveFormToCenter Me
GridHead
Set MSHFlexGrid1.Recordset = RsTemp
Set RsTemp = Nothing

End Sub
Private Sub GridHead()
Dim i As Integer

With MSHFlexGrid1

.Visible = False
.row = 0
.col = 0
.MergeCells = flexMergeRestrictRows
.MergeRow(0) = True
.col = 0
.ColWidth(0) = 1500
.col = 1
.ColWidth(1) = 1000
.col = 2
.ColWidth(2) = 1000
.col = 3
.ColWidth(3) = 1000
.col = 4
.ColWidth(4) = 1000
.col = 5
.ColWidth(5) = 1000
.col = 6
.ColWidth(6) = 1000
.col = 7
.ColWidth(7) = 0


.Visible = True
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm072004 = Nothing
End Sub
