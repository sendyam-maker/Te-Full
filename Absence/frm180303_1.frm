VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180303_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "打卡明細資料"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5745
   Tag             =   "加班資料"
   Begin VB.CommandButton cmdClose 
      Caption         =   "關閉(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4650
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   330
      Left            =   3450
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   5025
      Left            =   90
      TabIndex        =   1
      Top             =   480
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   8864
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "員工代號|姓名|刷卡日期|刷卡時間|人事補登|刷卡機"
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
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1.Caption = 接收打卡時間公告"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   5580
      Width           =   2700
   End
End
Attribute VB_Name = "frm180303_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Created by Sindy 2013/7/3
Option Explicit

Public m_B1401 As String '員工代號
Public m_B1402 As String '日期
Dim m_PrevForm As Form '前一畫面


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdClose_Click()
   m_PrevForm.bolClose = True
   Unload Me
End Sub

Private Sub cmdBack_Click()
   m_PrevForm.bolClose = False
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   If UCase(m_PrevForm.Name) = UCase("frm180303") Then
      cmdBack.Visible = True
   Else
      cmdBack.Visible = False
   End If
   
   Label1.Caption = 接收打卡時間公告
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm180303_1 = Nothing
End Sub

Public Function QueryData() As Boolean
   Dim stSQL As String
   
   QueryData = False
   
   Screen.MousePointer = vbHourglass
   Me.grdList.MousePointer = flexHourglass
   InitialGridList
   stSQL = "select st01 as 員工代號,st02 as 姓名,sqldatet(pr01) as 刷卡日期,sqltime6(pr02) as 刷卡時間,decode(pr08,999,'Y','') as 人事補登,decode(OMAN,null,pr09,OMAN) 刷卡機"
   stSQL = stSQL & " from staff, staffcarddata, pollrecord,setSpecMan where scd01(+)=st01 and pr03(+)=scd02 and pr01>0" & _
                    " and st01='" & m_B1401 & "' and pr01=" & DBDATE(m_B1402) & " and ocode(+)=pr09"
   stSQL = stSQL & " order by pr02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      Set grdList.Recordset = RsTemp
      grdList.row = 1
      QueryData = True
   Else
      ShowNoData
      Me.grdList.MousePointer = flexDefault
      Screen.MousePointer = vbDefault
      Exit Function
   End If
   Me.grdList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Function

' 初始化列表
Private Sub InitialGridList()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("員工代號", "姓名", "刷卡日期", "刷卡時間", "人事補登", "刷卡機")
   arrGridHeadWidth = Array(800, 800, 850, 850, 800, 1000)
   grdList.Visible = False
   grdList.Cols = UBound(arrGridHeadText) + 1
   grdList.Rows = 2
   For iRow = 0 To grdList.Cols - 1
      grdList.row = 0
      grdList.col = iRow
      grdList.Text = arrGridHeadText(iRow)
      grdList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdList.CellAlignment = flexAlignCenterCenter
   Next
   grdList.Visible = True
End Sub
