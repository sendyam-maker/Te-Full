VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090629 
   BorderStyle     =   1  '單線固定
   Caption         =   "特殊加乘註記維護"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7740
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8385
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090629.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4605
      Left            =   15
      TabIndex        =   7
      Top             =   675
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   8123
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090629.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblSales"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt1(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt1(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt1(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090629.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdList"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdQuery"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtQry"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbl2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(82)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3675
         Left            =   -74910
         TabIndex        =   19
         Top             =   840
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   6482
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   0
         AllowUserResizing=   1
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
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   400
         Left            =   -68340
         TabIndex        =   18
         Top             =   390
         Width           =   912
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   2070
         TabIndex        =   5
         Top             =   2055
         Width           =   705
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   2070
         TabIndex        =   4
         Top             =   1731
         Width           =   705
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   2070
         MaxLength       =   9
         TabIndex        =   0
         Top             =   435
         Width           =   945
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   2070
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1083
         Width           =   945
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   2070
         MaxLength       =   7
         TabIndex        =   1
         Top             =   759
         Width           =   945
      End
      Begin VB.TextBox txtQry 
         Height          =   270
         Left            =   -73000
         MaxLength       =   9
         TabIndex        =   6
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   2070
         TabIndex        =   3
         Top             =   1407
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "墨圖加乘註記："
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   17
         Top             =   2100
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "草圖加乘註記："
         Height          =   180
         Index           =   3
         Left            =   570
         TabIndex        =   16
         Top             =   1770
         Width           =   1260
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -72405
         TabIndex        =   15
         Top             =   465
         Width           =   45
      End
      Begin MSForms.Label lblSales 
         Height          =   255
         Left            =   3105
         TabIndex        =   13
         Top             =   480
         Width           =   4425
         VariousPropertyBits=   27
         Caption         =   "lblSales"
         Size            =   "7805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶或智權人員編號："
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   12
         Top             =   465
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "結束日期："
         Height          =   180
         Index           =   2
         Left            =   570
         TabIndex        =   11
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "起始日期："
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   10
         Top             =   795
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶或智權人員編號："
         Height          =   180
         Index           =   82
         Left            =   -74850
         TabIndex        =   9
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人加乘註記："
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   8
         Top             =   1440
         Width           =   1440
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm090629"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (grdList,lblSales)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'Create by Nickc 2004/03/16
Option Explicit
'前次紀錄KEY
Dim lst_OG01 As String
'本次紀錄KEY
Dim cur_OG01 As String
'目前狀態
Dim iCurState As Integer
'使用者權限設定
Dim bolInsert As Boolean
Dim bolUpdate As Boolean
Dim bolDelete As Boolean
Dim bolSelect As Boolean

'檢查查詢條件
Private Function CheckQueryData() As Boolean

   Dim bolCancel As Boolean, i As Integer
   
   If txtQry.Text = "" Then
        MsgBox "請輸入客戶或智權人員編號!!!", vbExclamation + vbOKOnly
        txtQry.SetFocus
        Exit Function
   End If
   
      Call txtQry_Validate(bolCancel)
      If bolCancel = True Then
         txtQry.SetFocus
         Exit Function
      End If
   CheckQueryData = True
   
End Function

Private Sub InitGrid()

   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer

   arrGridHeadText = Array("客戶或智權人員", "編號", "開始日期", "結束日期", "承辦人加乘註記" _
                     , "草圖加乘註記", "墨圖加乘註記")

   arrGridHeadWidth = Array(1000, 700, 850, 850, 1400 _
                     , 1400, 1400)

   With grdList
      .Cols = UBound(arrGridHeadText) + 1
      .row = 0
      For iCol = 0 To .Cols - 1
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignCenterCenter
      Next
   End With
   
   
End Sub

Private Function QueryData() As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset
   Dim strCon As String
   
On Error GoTo ErrHand

   strCon = ""
   If txtQry <> "" Then
      strCon = strCon & " AND ef03='" & txtQry.Text & "' "
   End If
   
   'Modify by Morgan 2010/8/17 百年蟲 " & SQLDate("ef01") & "-->substrb(' '||sqldatet(ef01),-9)," & SQLDate("ef02") & "-->substrb(' '||sqldatet(ef02),-9)
   strSql = "select decode(substr(ef03,1,1),'X',cu05,st02),ef03,substrb(' '||sqldatet(ef01),-9) AS ef01,substrb(' '||sqldatet(ef02),-9) as ef02,ef04,ef05,ef06 " & _
            " from exflag, staff,customer " & _
            " where ef03=st01(+) and substr(ef03,1,8)=cu01(+) and substr(ef03,9,1)=cu02(+) " & strCon & " ORDER BY ef03,ef01,ef02 "
            
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly

   If rsQuery.RecordCount > 0 Then
      QueryData = True
      Set Me.grdList.Recordset = rsQuery
      InitGrid
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
            
End Function



Private Sub cmdQuery_Click()
   
   If TxtValidate(1) = False Then Exit Sub
      If CheckQueryData = True Then
         Screen.MousePointer = vbHourglass
         grdList.MousePointer = flexHourglass
         grdList.Rows = 2
         grdList.Clear
         grdList.Refresh
         InitGrid
         If QueryData() = False Then
             MsgBox "無資料", vbOKOnly, "查詢資料"
             txtQry.SetFocus
         End If
         grdList.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
      End If

      
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
      '新增
         If SSTab1.Tab = 0 And TBar1.Buttons(1).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(1))
         End If
      Case vbKeyF3
      '修改
         If SSTab1.Tab = 0 And TBar1.Buttons(2).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(2))
         End If
      Case vbKeyF5
      '刪除
         If SSTab1.Tab = 0 And TBar1.Buttons(3).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(3))
         End If
      Case vbKeyF4
      '查詢
         If SSTab1.Tab = 0 And TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
         End If
      Case vbKeyHome
      '第一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(6).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(6))
         End If

      Case vbKeyPageUp
      '上一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(7).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(7))
         End If
      Case vbKeyPageDown
      '下一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(8).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(8))
         End If
      Case vbKeyEnd
      '最後筆
         If SSTab1.Tab = 0 And TBar1.Buttons(9).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(9))
         End If
      Case vbKeyF9
      '存檔
         If SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      Case vbKeyReturn
      '確定
         If SSTab1.Tab = 1 Then
            Call cmdQuery_Click
         ElseIf SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      Case vbKeyF10
      '取消
         If SSTab1.Tab = 0 And TBar1.Buttons(12).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
        If TBar1.Buttons(14).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(14))
         End If
    End Select
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   Me.Show
   setAuthority
   Call FormReset(0)
   Call InitGrid
   '預設為瀏覽
   If doQuery(6) = True Then
      iCurState = 0
   Else
      iCurState = 9
   End If
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
   
End Sub
'使用者權限設定
Private Sub setAuthority()
      bolInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
      bolUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
      bolDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
      bolSelect = IsUserHasRightOfFunction(Me.Name, strFind, False)
End Sub

Private Function TxtValidate(Optional ByVal iTab As Integer = 0) As Boolean

   Dim oText As TextBox, bolCancel As Boolean, arrText
   
   TxtValidate = False
   bolCancel = False
   
   Select Case iTab
      Case 0
         SSTab1.Tab = 0
         If bolCancel = False Then
            For Each oText In txt1
               If oText.Locked = False Then
                  txt1_Validate oText.Index, bolCancel
                  If bolCancel = True Then
                     oText.SetFocus
                     TextInverse oText
                     Exit For
                  End If
               End If
            Next
         End If
      Case 1
         If bolCancel = False Then
            If txtQry.Locked = False Then
               txtQry_Validate bolCancel
               If bolCancel = True Then
                  txtQry.SetFocus
                  TextInverse txtQry
               End If
            End If
         End If
   End Select
   

   
   If bolCancel = False Then TxtValidate = True
   
End Function




Private Function CheckConfirm() As Boolean
   
   CheckConfirm = False
   
   Select Case iCurState
      '1:新增;2:修改
      Case 1, 2
      
         If TxtValidate = False Then Exit Function
         
         If txt1(0) = "" Then
            MsgBox "客戶或智權人員不可空白！", vbCritical
            txt1(0).SetFocus
            Call txt1_GotFocus(0)
            Exit Function
         ElseIf txt1(1) = "" Then
            MsgBox "開始日期不可空白！", vbCritical
            txt1(1).SetFocus
            Call txt1_GotFocus(1)
            Exit Function
         ElseIf txt1(2) = "" Then
            MsgBox "結束日期不可空白！", vbCritical
            txt1(2).SetFocus
            Call txt1_GotFocus(2)
            Exit Function
         ElseIf txt1(3) = "" And txt1(4) = "" And txt1(5) = "" Then
               MsgBox "加乘註記至少輸入一個！", vbCritical
               txt1(3).SetFocus
               Call txt1_GotFocus(3)
               Exit Function
         End If
         
      '查詢
      Case 4
         If txt1(0) = "" Then
            MsgBox "序號不可空白！", vbCritical
            txt1(0).SetFocus
            Call txt1_GotFocus(0)
            Exit Function
         End If
   End Select
   CheckConfirm = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090629 = Nothing
End Sub

Private Sub grdList_DblClick()

   Dim lRow As Long, lCurRow As Long, iCol As Integer
   
   lCurRow = grdList.row
   '呼叫查詢
   If lCurRow > 0 Then
      If TBar1.Buttons(4).Enabled = True Then
         Call Tbar1_ButtonClick(TBar1.Buttons(4))
         If txt1(0).Locked = False Then
            txt1(0).Text = grdList.TextMatrix(lCurRow, 1)
            txt1(1).Text = Replace(grdList.TextMatrix(lCurRow, 2), "/", "")
            txt1(2).Text = Replace(grdList.TextMatrix(lCurRow, 3), "/", "")
            If TBar1.Buttons(11).Enabled = True Then
               Call Tbar1_ButtonClick(TBar1.Buttons(11))
            End If
         End If
      End If
   End If
   
End Sub

Private Sub grdList_Click()
      
   Dim lRow As Long, lCurRow As Long, iCol As Integer
   
   With grdList
      lCurRow = .row
      If lCurRow > 0 Then
         '還原
         For lRow = 1 To .Rows - 1
            .row = lRow: iCol = 1
            If .CellBackColor <> &H80000005 Then
               For iCol = 0 To .Cols - 1
                   .col = iCol
                   .CellBackColor = QBColor(15)
               Next iCol
            End If
         Next lRow
         '反白
         .row = lCurRow
         For iCol = 0 To .Cols - 1
             .col = iCol
             .CellBackColor = &HFFC0C0
         Next iCol
      End If
      
   End With
   
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case PreviousTab
      Case 0
         If iCurState = 0 Then txtQry.SetFocus
   End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
      '新增
         iCurState = 1
      Case 2
      '修改
         iCurState = 2
      Case 3
      '刪除
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            If DeleteData = True Then
               If doQuery(8, False) = True Then
                  iCurState = 0
               ElseIf doQuery(9) = True Then
                  iCurState = 0
               Else
                  cur_OG01 = ""
                  iCurState = 9
               End If
            End If
         End If
      Case 4
      '查詢
         iCurState = 4
      Case 6
      '第一筆
         Call doQuery(6)
      Case 7
      '上一筆
         Call doQuery(7)
      Case 8
      '下一筆
         Call doQuery(8)
      Case 9
      '最後筆
         Call doQuery(9)
      Case 11
      '確定
         If CheckConfirm = False Then Exit Sub
         Select Case iCurState
            '新增
            Case 1
               If insertdata() = False Then
                  txt1(1).SetFocus
                  txt1_GotFocus (1)
                  Exit Sub
               End If
            '查詢
            Case 4
               cur_OG01 = ChangeTStringToWString(txt1(1)) & ChangeTStringToWString(txt1(2)) & txt1(0)
               
            '修改
            Case 2
               If UpdateData() = False Then
                  Exit Sub
               End If
               
         End Select
         '重新查詢
         If doQuery(4) = True Then
            Call SetToolBar(0)
            Call SetInputs
         Else
            If iCurState = 4 Then
               'Call SetToolBar(0)
               Call Tbar1_ButtonClick(TBar1.Buttons(12))
               txt1(0).SetFocus
               Call txt1_GotFocus(0)
            End If
            Exit Sub
         End If
         iCurState = 0
      Case 12
      '取消
         Select Case iCurState
            
            '1:新增
            Case 1
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               ElseIf cur_OG01 = "" Then
                  If doQuery(6) = True Then
                     iCurState = 0
                  Else
                     iCurState = 9
                  End If
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
            '2:修改
            Case 2
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
            '查詢
            Case 4
               cur_OG01 = lst_OG01
               If cur_OG01 = "" Then
                  If doQuery(6) = True Then
                     iCurState = 0
                  Else
                     iCurState = 9
                  End If
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  'Call SetToolBar(0)
                  Call Tbar1_ButtonClick(TBar1.Buttons(12))
                  Exit Sub
               End If
         End Select
      Case 14
      '結束
         If iCurState = 2 Or iCurState = 1 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               Unload Me
               Exit Sub
            End If
         Else
            Unload Me
            Exit Sub
         End If
         
   End Select
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
   lst_OG01 = cur_OG01
   
End Sub
'清除畫面
Private Sub FormReset(Optional ByVal iTab As Integer = 0)

   Dim oText As TextBox
   
   Select Case iTab
   
      Case 0
      '頁籤0
         For Each oText In txt1
            oText.Text = ""
         Next
         lblSales.Caption = ""
         If txt1(0).Enabled = True Then
            txt1(0).SetFocus
         End If
      Case 1
      '頁籤1
         txtQry.Text = ""
   End Select
End Sub
'工具列控制
Private Sub SetToolBar(Optional ByVal iStatus As Integer)

   Dim i As Integer
   For i = 1 To 13
      TBar1.Buttons(i).Enabled = False
   Next
   TBar1.Buttons(14).Enabled = True
   
   Select Case iStatus
   
      Case 0
      '瀏覽
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         If bolUpdate Then
            TBar1.Buttons(2).Enabled = True
         End If
         If bolDelete Then
            TBar1.Buttons(3).Enabled = True
         End If
         If bolSelect Then
            TBar1.Buttons(4).Enabled = True
         End If
         TBar1.Buttons(6).Enabled = True
         TBar1.Buttons(7).Enabled = True
         TBar1.Buttons(8).Enabled = True
         TBar1.Buttons(9).Enabled = True
         
      Case 1, 2, 4
      '1:新增  '2:修改  '4查詢
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
               
      Case 9
      '無資料
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         
   End Select
   
End Sub
'設定文字框
Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)

   Dim oText As TextBox, oLabel As Label, oMaskEdBox As MaskEdBox
   
   Select Case iStatus
      
      Case 0
      '瀏覽
         For Each oText In txt1
            oText.Enabled = True
            oText.Locked = True
         Next
         If txt1(0).Enabled = True Then
            txt1(0).SetFocus
         End If
      Case 1
      '新增
         SSTab1.Tab = 0
         For Each oText In txt1
            oText.Text = ""
            oText.Locked = False
            oText.Enabled = True
         Next
         Call FormReset(0)
      Case 2
      '修改
         SSTab1.Tab = 0
         For Each oText In txt1
            oText.Locked = False
            oText.Enabled = True
         Next
         txt1(0).Locked = True
         txt1(1).Locked = True
         txt1(2).Locked = True
      Case 4
      '查詢
         SSTab1.Tab = 0
         For Each oText In txt1
            oText.Locked = False
            oText.Enabled = False
         Next
         Call FormReset(0)
         txt1(0).Enabled = True
         txt1(1).Enabled = True
         txt1(2).Enabled = True
         txt1(0).SetFocus
      Case 9
      '無資料
         For Each oText In txt1
            oText.Enabled = False
            oText.Locked = True
         Next
         Call FormReset(0)
   End Select
   
End Sub
'讀取資料
Private Function doQuery(ByVal iAct As Integer, Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery = False
   
   Select Case iAct
      Case 4
      '查詢
         strSql = "Select ef01,ef02,ef03 From exflag where to_char(ef01)||to_char(ef02)||ef03='" & cur_OG01 & "' "
         stMessage = "查無資料！"
   
      Case 6
      '第一筆
         strSql = "Select ef01,ef02,ef03 From exflag ORDER BY 1 ASC,2 asc,3 asc"
         stMessage = "無特殊加乘註記紀錄！"
      Case 7
      '上一筆
         strSql = "Select ef01,ef02,ef03 From exflag where to_char(ef01)||to_char(ef02)||ef03<'" & cur_OG01 & "' " & _
            " ORDER BY 1 DESC,2 desc,3 desc"
         stMessage = "已是第一筆了！"

      Case 8
      '下一筆
         strSql = "Select ef01,ef02,ef03 From exflag where to_char(ef01)||to_char(ef02)||ef03>'" & cur_OG01 & "' " & _
            " ORDER BY 1 ASC,2 asc,3 asc"
         stMessage = "已是最後一筆了！"

      Case 9
      '最後筆
         strSql = "Select ef01,ef02,ef03 From exflag" & _
            " ORDER BY 1 DESC,2 desc,3 desc"
         stMessage = "無特殊加乘註記紀錄！"
        
   End Select
   
On Error GoTo ErrHand

   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
         lst_OG01 = cur_OG01
         cur_OG01 = "" & CheckStr(rsQuery.Fields(0).Value) & CheckStr(rsQuery.Fields(1).Value) & CheckStr(rsQuery.Fields(2).Value)
         If ReQuery() = True Then doQuery = True
   ElseIf bolMsg Then
      MsgBox stMessage, vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
   
End Function


'完整資料查詢
Private Function ReQuery(Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, intI As Integer
   
On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   
   ReQuery = False
   'Modify by Morgan 2010/8/17 百年蟲 " & SQLDate("ef01") & "-->substrb(' '||sqldatet(ef01),-9)," & SQLDate("ef02") & "-->substrb(' '||sqldatet(ef02),-9)
   strSql = "select decode(substr(ef03,1,1),'X',cu05,st02),ef03,substrb(' '||sqldatet(ef01),-9) AS ef01,substrb(' '||sqldatet(ef02),-9) as ef02,ef04,ef05,ef06 " & _
            " from exflag, staff,customer " & _
         " where ef03=st01(+) and substr(ef03,1,8)=cu01(+) and substr(ef03,9,1)=cu02(+) and to_char(ef01)||to_char(ef02)||ef03='" & cur_OG01 & "' ORDER BY ef03,ef01,ef02 "

   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      txt1(0).Text = "" & rsQuery.Fields("ef03").Value
      lblSales.Caption = "" & rsQuery.Fields(0).Value
      txt1(1).Text = Replace(CheckStr(rsQuery.Fields("ef01").Value), "/", "")
      txt1(2).Text = Replace(CheckStr(rsQuery.Fields("ef02").Value), "/", "")
      txt1(3).Text = "" & rsQuery.Fields("ef04").Value
      txt1(4).Text = "" & rsQuery.Fields("ef05").Value
      txt1(5).Text = "" & rsQuery.Fields("ef06").Value
      ReQuery = True
   ElseIf bolMsg Then
      MsgBox "特殊加乘註記〔" & cur_OG01 & "〕已被刪除！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Screen.MousePointer = vbDefault
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
   Screen.MousePointer = vbDefault
   
End Function


Private Function DeleteData() As Boolean
   Dim strSql As String, lngEffRec As Long
   
   strSql = "Delete exflag Where ef01=" & ChangeTStringToWString(txt1(1).Text) & " and ef02=" & ChangeTStringToWString(txt1(2).Text) & " and ef03='" & txt1(0).Text & "'   "
   
   DeleteData = False
   
On Error GoTo ErrHnd

   cnnConnection.Execute strSql
   DeleteData = True
   
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Function UpdateData() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, ef(4 To 6) As String
   Dim rsQuery As New ADODB.Recordset, strUpdSQL As String, lngEffRec As Long
   
   ef(4) = "ef04=" & txt1(3).Text
   ef(5) = "ef05=" & txt1(4).Text
   ef(6) = "ef06=" & txt1(5).Text

   
   strUpdSQL = Join(ef, ",")
   
   strSql = "Update exflag Set " & strUpdSQL & " Where ef01=" & ChangeTStringToWString(txt1(1)) & " and ef02=" & ChangeTStringToWString(txt1(2)) & " and ef03='" & txt1(0).Text & "' "
         
   UpdateData = False
   
On Error GoTo ErrHnd

   cnnConnection.Execute strSql
   UpdateData = True
   
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Function insertdata() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, ef(1 To 6) As String
   Dim strCols As String, strValues As String, lngEffRec As Long
   Dim rsQuery As New ADODB.Recordset
   
   ' 檢查是否區間重複
   strSql = "select * from exflag where ef01<=" & ChangeTStringToWString(txt1(1)) & " and ef02>=" & ChangeTStringToWString(txt1(1)) & " and ef03='" & txt1(0) & "' "
   strSql = strSql & " union select * from exflag where ef01<=" & ChangeTStringToWString(txt1(2)) & " and ef02>=" & ChangeTStringToWString(txt1(2)) & " and ef03='" & txt1(0) & "' "
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 Then
      MsgBox "此客戶或智權人員在 " & ChangeTStringToTDateString(txt1(1)) & " 到 " & ChangeTStringToTDateString(txt1(2)) & " 已經設定，無法再重複設定！", vbCritical
      insertdata = False
      CheckOC3
      Exit Function
   End If
   CheckOC3
   strCols = "ef01"
   For intI = 2 To 6
      strCols = strCols & ",ef" & Format(intI, "00")
   Next intI
   
   ef(1) = ChangeTStringToWString(txt1(1).Text)
   ef(2) = ChangeTStringToWString(txt1(2).Text)
   ef(3) = "'" & txt1(0).Text & "'"
   ef(4) = Val(txt1(3).Text)
   ef(5) = Val(txt1(4).Text)
   ef(6) = Val(txt1(5).Text)
   
   strValues = Join(ef, ",")
   
   strSql = "INSERT INTO exflag (" & strCols & ") VALUES(" & strValues & ")"
         
   insertdata = False
   
On Error GoTo ErrHnd

   cnnConnection.Execute strSql

   cur_OG01 = ChangeTStringToWString(txt1(1)) & ChangeTStringToWString(txt1(2)) & txt1(0)
   insertdata = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1, 2
      If Not (KeyAscii = 13 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
         KeyAscii = 0
      End If
Case 3, 4, 5
      If Not (KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 43 Or KeyAscii = 45 Or (KeyAscii > 47 And KeyAscii < 58)) Then
         KeyAscii = 0
      End If
Case Else
End Select
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)

If Len(Trim(txt1(Index))) = 0 Then Exit Sub
Select Case Index
Case 0
         If Mid(txt1(Index), 1, 1) = "X" Then txt1(Index) = txt1(Index) & "000000000"
         strSql = "select st02 from staff where st01='" & txt1(Index) & "' "
         strSql = strSql & " union select cu04 from customer where cu01='" & Mid(txt1(Index), 1, 8) & "' and cu02='" & Mid(txt1(Index), 9, 1) & "' "
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount <> 0 Then
            lblSales.Caption = CheckStr(AdoRecordSet3.Fields(0).Value)
         Else
            MsgBox "查無此客戶或智權人員！", vbCritical
            txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
         End If
         CheckOC3
Case 1
         If PUB_CheckKeyInDate(txt1(Index)) < 0 Then Cancel = True: txt1(Index).SetFocus: txt1_GotFocus Index
Case 2
         If PUB_CheckKeyInDate(txt1(Index)) = 0 Then
            'Modify by Morgan 2010/8/17 百年蟲
            'If txt1(1) <> "" And (txt1(2) < txt1(1)) Then
            If txt1(1) <> "" And Val(txt1(2)) < Val(txt1(1)) Then
               MsgBox "結束日期必需大於開始日期！", vbCritical
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
            End If
         Else
            txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
         End If
Case 3, 4, 5
         If Val(txt1(Index)) > 3 Or Val(txt1(Index)) < 0 Then
            MsgBox "加乘註記範圍請介於 0 ~3 ！", vbCritical
            txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
         End If
Case Else
End Select
End Sub

Private Sub txtQry_GotFocus()
txtQry.SelStart = 0
txtQry.SelLength = Len(txtQry)
End Sub

Private Sub txtQry_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub


Private Sub txtQry_Validate(Cancel As Boolean)
If Len(Trim(txtQry)) = 0 Then Exit Sub
         strSql = "select st02 from staff where st01='" & txtQry & "' "
         strSql = strSql & " union select cu05 from customer where cu01='" & Mid(txtQry, 1, 8) & "' and cu02='" & Mid(txtQry, 9, 1) & "' "
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount <> 0 Then
            'lblSales.Caption = CheckStr(AdoRecordSet3.Fields(0).Value)
         Else
            MsgBox "查無此客戶或智權人員！", vbCritical
            Cancel = True
         End If
         CheckOC3
End Sub
