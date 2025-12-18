VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm140404_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "往來記錄回覆清單"
   ClientHeight    =   5430
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8745
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Index           =   0
      Left            =   7425
      TabIndex        =   2
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   1
      Left            =   6255
      TabIndex        =   0
      Top             =   30
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   225
      Top             =   90
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4635
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   8176
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "選擇|往來記錄編號|往來日期|回覆期限|主旨|往來類別"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm140404_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/11 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2007/12/4
Option Explicit

Dim m_iSelRow As Integer
Public fmParent As Form

Private Sub cmdOK_Click(Index As Integer)
   Dim stRefNo2 As String
   Select Case Index
      Case 1
         If CheckCheck(stRefNo2) = False Then
            If fmParent.Tag = "" Then
               strExc(1) = "是否確定不回覆？"
            Else
               strExc(1) = "是否確定取消所有回覆？"
            End If
            If MsgBox(strExc(1), vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         fmParent.Tag = stRefNo2
   End Select
   Unload Me
End Sub

Private Function CheckCheck(Optional p_No2 As String) As Boolean
   Dim ii As Integer
   With grdDataList
      p_No2 = ""
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            p_No2 = p_No2 & "," & .TextMatrix(ii, 1)
            CheckCheck = True
         End If
      Next
      If p_No2 <> "" Then
         p_No2 = Mid(p_No2, 2)
      End If
   End With
End Function

Private Sub Form_Activate()
   SetDataListWidth
   grdPaintSelected
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140404_1 = Nothing
End Sub

Private Sub SetDataListWidth()
   With grdDataList
      .FormatString = .FormatString
      .ColWidth(0) = 465
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(1) = 1305
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 885
      .ColAlignment(2) = flexAlignLeftCenter
      .ColWidth(3) = 885
      .ColAlignment(3) = flexAlignLeftCenter
      .ColWidth(4) = 3150
      .ColAlignment(4) = flexAlignLeftCenter
      .ColWidth(5) = 1350
      .ColAlignment(5) = flexAlignLeftCenter
   End With
End Sub

Private Sub grdSelected(p_iRow As Integer)
   Dim lColor As Long, ii As Integer
   With grdDataList
      .row = p_iRow
      .col = 0
      If .Text = "" Then
         .Text = "V"
         lColor = &HFFC0C0
      Else
         .Text = ""
         lColor = &H80000018
      End If
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = lColor
      Next
   End With
End Sub

Private Sub GrdDataList_Click()
   Dim iRow As Integer
   With grdDataList
      If .MouseRow > 0 And .MouseRow < .Rows Then
         .Visible = False
         iRow = .MouseRow
         grdSelected iRow
'         If m_iSelRow <> 0 And m_iSelRow <> iRow Then
'            grdSelected m_iSelRow
'         End If
         m_iSelRow = iRow
         .Visible = True
      End If
   End With
End Sub

Private Sub grdPaintSelected()
   Dim lColor As Long, ii As Integer, jj As Integer
   With grdDataList
      For jj = 1 To .Rows - 1
         .row = jj
         .col = 0
         If .Text = "V" Then
            lColor = &HFFC0C0
         Else
            lColor = &H80000018
         End If
         For ii = 0 To .Cols - 1
            .col = ii
            .CellBackColor = lColor
         Next ii
      Next jj
   End With
End Sub
