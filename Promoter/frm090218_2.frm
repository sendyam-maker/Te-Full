VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090218_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "暫停核稿記錄查詢"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5940
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Left            =   4740
      TabIndex        =   0
      Top             =   30
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4365
      Left            =   60
      TabIndex        =   1
      Top             =   750
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7699
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   1
   End
   Begin VB.Label lbl1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   3795
   End
End
Attribute VB_Name = "frm090218_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Private Sub cmdOK_Click()
Me.Hide
frm090218_1.Show
frm090218_1.PubShowNextData
Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090218_2 = Nothing
End Sub

Public Sub StrMenu()
Screen.MousePointer = vbHourglass
grd1.MousePointer = flexArrowHourGlass
DoEvents
grd1.Clear
grd1.Rows = 2
SetGrd
lbl1.Caption = "收文號：" & Me.Tag
strSql = "select sqldatet(em03),decode(Em04,0,'繼續',1,'暫停',''),st02,em02 from EngManLog,staff  where em01='" & Trim(Me.Tag) & "' and em05=st01(+)  order by em02 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    Set grd1.Recordset = adoRecordset
    SetGrd
Else
    ShowNoData
    cmdOK_Click
    Exit Sub
End If
CheckOC
grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("異動日期", "異動內容", "異動人員", "")
   arrGridHeadWidth = Array(800, 800, 800, 0)
                        
   grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
End Sub

