VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210112_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件性質選擇畫面"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdok 
      Caption         =   "取消(&C)"
      Height          =   300
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Height          =   300
      Index           =   0
      Left            =   2745
      TabIndex        =   1
      Top             =   45
      Width           =   900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   4365
      Left            =   75
      TabIndex        =   0
      Top             =   405
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   7699
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm210112_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; Grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim i As Integer
Dim j As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
         For i = 1 To Grd1.Rows - 1
            Grd1.row = i
            Grd1.col = 0
            If Grd1.CellBackColor = &HFFC0C0 Then
               If Trim(frm210112.txt_CP10(frm210112.SeekIdx).Text) <> "" Then
                  If Trim(frm210112.txt_CP10(frm210112.SeekIdx).Text) = "ALL" Then
                     frm210112.txt_CP10(frm210112.SeekIdx).Text = Grd1.Text
                      For j = 0 To Grd1.Cols - 1
                           Grd1.col = j
                           Grd1.CellBackColor = &H80000018
                     Next j
                  Else
                     frm210112.txt_CP10(frm210112.SeekIdx).Text = frm210112.txt_CP10(frm210112.SeekIdx).Text & "," & Grd1.Text
                      For j = 0 To Grd1.Cols - 1
                           Grd1.col = j
                           Grd1.CellBackColor = &H80000018
                     Next j
                  End If
               Else
                     frm210112.txt_CP10(frm210112.SeekIdx).Text = Grd1.Text
                      For j = 0 To Grd1.Cols - 1
                           Grd1.col = j
                           Grd1.CellBackColor = &H80000018
                     Next j
               End If
            End If
         Next i
         Unload Me
Case 1
         Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Select Case frm210112.SeekIdx
Case 0
        strSql = "select cpm02,decode(instr(cpm03,'（無）'),0,cpm03,cpm04) as cpm03 from casepropertymap where cpm01 in ('P','PS','CPS','CFP') and length(cpm02)=3  "
Case 1
        strSql = "select cpm02,decode(instr(cpm03,'（無）'),0,cpm03,cpm04) as cpm03 from casepropertymap where cpm01 in ('T','TF','CFT') and length(cpm02)=3  "
Case 2
        strSql = "select cpm02,decode(instr(cpm03,'（無）'),0,cpm03,cpm04) as cpm03 from casepropertymap where cpm01 in ('L','LA') and length(cpm02)=3  "
Case 3
        strSql = "select cpm02,decode(instr(cpm03,'（無）'),0,cpm03,cpm04) as cpm03 from casepropertymap where cpm01 in ('CFC','TC') and length(cpm02)=3  "
Case Else
End Select
CheckOC3
With AdoRecordSet3
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
       Set Grd1.Recordset = AdoRecordSet3
       SetDataListWidth
   End If
End With
End Sub

Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

arrGridHeadText = Array("代號", "中文說明")
                     
arrGridHeadWidth = Array(1000, 2000)
                     
Grd1.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To Grd1.Cols - 1
   Grd1.row = 0
   Grd1.col = iRow
   Grd1.Text = arrGridHeadText(iRow)
   Grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
   Grd1.CellAlignment = flexAlignCenterCenter
Next
End Sub



Private Sub Form_Unload(Cancel As Integer)

Set frm210112_1 = Nothing
End Sub

Private Sub grd1_SelChange()
Dim ClickRow As Integer
Grd1.Visible = False
ClickRow = Grd1.MouseRow
If ClickRow >= 1 Then
   Grd1.row = ClickRow
     For i = 0 To Grd1.Cols - 1
         Grd1.col = i
         Grd1.CellBackColor = &HFFC0C0
     Next i
End If
Grd1.Visible = True
End Sub
