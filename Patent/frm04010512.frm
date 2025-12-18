VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010512 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人信件收達管制"
   ClientHeight    =   5508
   ClientLeft      =   1752
   ClientTop       =   1860
   ClientWidth     =   9144
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5508
   ScaleWidth      =   9144
   Begin TabDlg.SSTab SSTab1 
      Height          =   4605
      Left            =   180
      TabIndex        =   27
      Top             =   750
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   8128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm04010512.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDisp(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDisp(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDisp(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDisp(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDisp(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDisp(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblFMP"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtFM(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtFM(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtFM(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtFM(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtFM(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFM(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm04010512.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(82)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line1(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "grdList"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdQuery(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtQry(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtQry(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdQuery(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtQry(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtQry(4)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtQry(5)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtQry(6)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Check1"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -71880
         TabIndex        =   30
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtQry 
         Height          =   270
         Index           =   6
         Left            =   -68895
         MaxLength       =   2
         TabIndex        =   11
         Top             =   450
         Width           =   435
      End
      Begin VB.TextBox txtQry 
         Height          =   270
         Index           =   5
         Left            =   -69300
         MaxLength       =   1
         TabIndex        =   10
         Top             =   450
         Width           =   315
      End
      Begin VB.TextBox txtQry 
         Height          =   270
         Index           =   4
         Left            =   -70305
         MaxLength       =   6
         TabIndex        =   9
         Top             =   450
         Width           =   915
      End
      Begin VB.TextBox txtQry 
         Height          =   270
         Index           =   3
         Left            =   -70935
         MaxLength       =   3
         TabIndex        =   8
         Top             =   450
         Width           =   525
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "列印(&P)"
         Height          =   400
         Index           =   1
         Left            =   -67380
         TabIndex        =   13
         Top             =   390
         Width           =   912
      End
      Begin VB.TextBox txtFM 
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   3495
         MaxLength       =   2
         TabIndex        =   1
         Top             =   600
         Width           =   705
      End
      Begin VB.TextBox txtQry 
         Height          =   270
         Index           =   2
         Left            =   -72990
         MaxLength       =   7
         TabIndex        =   7
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtQry 
         Height          =   270
         Index           =   1
         Left            =   -74040
         MaxLength       =   7
         TabIndex        =   6
         Top             =   450
         Width           =   945
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   400
         Index           =   0
         Left            =   -68340
         TabIndex        =   12
         Top             =   390
         Width           =   912
      End
      Begin VB.TextBox txtFM 
         Height          =   270
         Index           =   1
         Left            =   1530
         MaxLength       =   8
         TabIndex        =   0
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox txtFM 
         Height          =   270
         Index           =   3
         Left            =   1515
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1020
         Width           =   525
      End
      Begin VB.TextBox txtFM 
         Height          =   270
         Index           =   4
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtFM 
         Height          =   270
         Index           =   5
         Left            =   3150
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1020
         Width           =   315
      End
      Begin VB.TextBox txtFM 
         Height          =   270
         Index           =   6
         Left            =   3555
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1020
         Width           =   435
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3372
         Left            =   -74952
         TabIndex        =   31
         Top             =   1128
         Width           =   8652
         _ExtentX        =   15261
         _ExtentY        =   5948
         _Version        =   393216
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
      Begin VB.Label lblFMP 
         Caption         =   "此為寰華案件"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   600
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -70920
         X2              =   -68790
         Y1              =   580
         Y2              =   580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   4
         Left            =   -71880
         TabIndex        =   28
         Top             =   480
         Width           =   900
      End
      Begin MSForms.Label lblDisp 
         Height          =   180
         Index           =   4
         Left            =   4140
         TabIndex        =   26
         Top             =   4320
         Width           =   885
         VariousPropertyBits=   27
         Caption         =   "lblName"
         Size            =   "1561;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Update："
         Height          =   180
         Index           =   7
         Left            =   3330
         TabIndex        =   25
         Top             =   4320
         Width           =   645
      End
      Begin MSForms.Label lblDisp 
         Height          =   180
         Index           =   5
         Left            =   5040
         TabIndex        =   24
         Top             =   4320
         Width           =   795
         VariousPropertyBits=   27
         Caption         =   "lblDate"
         Size            =   "1402;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblDisp 
         Height          =   180
         Index           =   6
         Left            =   5730
         TabIndex        =   23
         Top             =   4320
         Width           =   840
         VariousPropertyBits=   27
         Caption         =   "lblTime"
         Size            =   "1482;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblDisp 
         Height          =   180
         Index           =   3
         Left            =   2520
         TabIndex        =   21
         Top             =   4320
         Width           =   750
         VariousPropertyBits=   27
         Caption         =   "lblTime"
         Size            =   "1323;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblDisp 
         Height          =   180
         Index           =   2
         Left            =   1680
         TabIndex        =   20
         Top             =   4320
         Width           =   795
         VariousPropertyBits=   27
         Caption         =   "lblDate"
         Size            =   "1402;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "序        號："
         Height          =   180
         Index           =   1
         Left            =   2550
         TabIndex        =   19
         Top             =   645
         Width           =   900
      End
      Begin VB.Line Line2 
         X1              =   -73260
         X2              =   -72840
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收信日期:"
         Height          =   180
         Index           =   82
         Left            =   -74850
         TabIndex        =   18
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收信日期："
         Height          =   180
         Index           =   0
         Left            =   630
         TabIndex        =   17
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   3
         Left            =   630
         TabIndex        =   16
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Creat："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   4320
         Width           =   555
      End
      Begin MSForms.Label lblDisp 
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   14
         Top             =   4320
         Width           =   795
         VariousPropertyBits=   27
         Caption         =   "lblName"
         Size            =   "1402;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1860
         X2              =   3990
         Y1              =   1170
         Y2              =   1170
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8385
      Top             =   330
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":084C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":0E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04010512.frx":1E10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9144
      _ExtentX        =   16129
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
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
Attribute VB_Name = "frm04010512"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Morgan 2021/12/10 改成Form2.0 (lblDisp,grdList)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Toni 2008/8/19
Option Explicit
'前次紀錄KEY
Dim lst_FM01 As String      '日期
Dim lst_FM02 As String      '序號
'本次紀錄KEY
Dim cur_FM01 As String      '日期
Dim cur_FM02 As String      '序號
'目前狀態
Dim iCurState As Integer
'使用者權限設定
Dim bolInsert As Boolean
Dim bolUpdate As Boolean
Dim bolDelete As Boolean
Dim bolSelect As Boolean
'列印控制
Dim PLeft(0 To 3) As Integer
Dim iPrint As Integer
Dim Page As Integer
Dim m_bNoTab As Boolean
'暫存日期 2008/8/28 add by Toni 自己入收信日期
Dim Tem_Date As String
'Added by Lydia 2015/04/28
Dim strQFMP As String '控制查詢
Dim strFMP As String '控制部門可新增的案件
Dim intCno As Integer 'Grid案號的起始欄位
Dim bolCaseMsg As Boolean '防止2次訊息
Dim CIDst03 ' Create ID所屬部門
'檢查查詢條件
Private Function CheckQueryData() As Boolean

   Dim bolCancel As Boolean, i As Integer
   
   '2008/9/10 modify by sonia
   If txtQry(1).Text = "" And txtQry(2).Text = "" And txtQry(3).Text = "" And txtQry(4).Text = "" And txtQry(5).Text = "" And txtQry(6).Text = "" Then
        MsgBox "請輸入查詢條件!!!", vbExclamation + vbOKOnly
        txtQry(2).SetFocus
        Exit Function
   End If
   '2008/9/10 end
   
   If txtQry(2) <> "" And txtQry(1).Text = "" Then
        MsgBox "請輸入收信起日!!!", vbExclamation + vbOKOnly
        txtQry(1).SetFocus
        Exit Function
   End If
   If txtQry(1) <> "" And txtQry(2).Text = "" Then
        MsgBox "請輸入收信迄日!!!", vbExclamation + vbOKOnly
        txtQry(2).SetFocus
        Exit Function
   End If
   
   For i = 1 To 6
      Call txtQry_Validate(i, bolCancel)
      If bolCancel = True Then
         txtQry(i).SetFocus
         Exit Function
      End If
   Next
   CheckQueryData = True
   
End Function

Private Sub InitGrid()

   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   'Added by Lydia 2015/04/28 +寰華案
'   arrGridHeadText = Array("", "收信日期", "序號", "本所案號")
'   arrGridHeadWidth = Array(300, 1000, 1000, 1500)
   arrGridHeadText = Array("", "收信日期", "序號", "本所案號", "FM03", "FM04", "FM05", "FM06", "寰華案")
   arrGridHeadWidth = Array(300, 1000, 1000, 1500, 0, 0, 0, 0, 800)
   intCno = 4
   
   With grdList
      .row = 0
      .Cols = UBound(arrGridHeadText) + 1
      For iCol = 0 To .Cols - 1
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignCenterCenter
      Next
      .Rows = 1
   End With
   
   
End Sub

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)

   Dim iRow As Integer, iCol As Integer
   Dim PP(1 To 4) As String 'Added by Lydia 2015/04/28
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
   
      With grdList
         .Rows = .Rows + 1
         iRow = .Rows - 1
         'Added by Lydia 2015/04/28 +寰華案
'        For iCol = 1 To grdList.Cols - 1
'            .TextMatrix(iRow, iCol) = "" & rsTmp.Fields(iCol - 1).Value
'        Next iCol
        For iCol = 1 To grdList.Cols - 2
            .TextMatrix(iRow, iCol) = "" & rsTmp.Fields(iCol - 1).Value
            If iCol >= intCno Then
               PP(iCol - intCno + 1) = "" & rsTmp.Fields(iCol - 1).Value
            End If
        Next iCol
        If PUB_FMPtoCheck(1, 1, "31", PP(1), PP(2), PP(3), PP(4)) = True Then
           .TextMatrix(iRow, grdList.Cols - 1) = "Y"
        End If
       'end 2015/04/28
      End With
      rsTmp.MoveNext
   Loop
   'Added by Lydia 2023/10/17
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/17
End Sub

Private Function QueryData() As Boolean
Dim strSql As String, strSQL1 As String
Dim rsQuery As New ADODB.Recordset
Dim strCon As String
   
On Error GoTo ErrHand
                                                                                                                                                                         
   '2008/9/10 modify by sonia  查詢條件加本所案號
   'strSQL = "select SUBSTRB(FM01,1,4)-1911||'/'||SUBSTRB(FM01,5,2)||'/'||SUBSTRB(FM01,7,2)AS FM01, FM02,DECODE(FM03,NULL,NULL,FM03||'-'||FM04||'-'||FM05||'-'||FM06)" & _
   '          " from FagentMail where  FM01>='" & ChangeTStringToWString(txtQry(1)) & "'  and FM01 <= '" & ChangeTStringToWString(txtQry(2)) & "' ORDER BY FM01,FM02"
   strSQL1 = ""
   If txtQry(1) <> "" Then
      strSQL1 = "FM01>='" & ChangeTStringToWString(txtQry(1)) & "' and FM01 <= '" & ChangeTStringToWString(txtQry(2)) & "' "
   End If
   If txtQry(3) <> "" Then
      If strSQL1 <> "" Then strSQL1 = strSQL1 & "and "
      strSQL1 = strSQL1 & "fm03='" & txtQry(3) & "' and fm04='" & txtQry(4) & "' and fm05='" & txtQry(5) & "' and fm06='" & txtQry(6) & "' "
   End If
   'Added by Lydia 2015/04/28 +寰華案
   '外專人員+勾選只顯示寰華案件 ,非外專人員+不勾選含寰華案件 ,
   If (Left(Pub_StrUserSt03, 1) = "F" And Check1.Value = 1) Or (Left(Pub_StrUserSt03, 1) <> "F" And Check1.Value = 0) Then
      strSQL1 = strSQL1 + strQFMP
   End If
   'end 2015/04/28
   
   'Modify by Morgan 2010/8/11 百年蟲
   'strSql = "select SUBSTRB(FM01,1,4)-1911||'/'||SUBSTRB(FM01,5,2)||'/'||SUBSTRB(FM01,7,2)AS FM01, FM02,DECODE(FM03,NULL,NULL,FM03||'-'||FM04||'-'||FM05||'-'||FM06)" & _
             " from FagentMail where " & strSQL1 & " ORDER BY FM01,FM02"
   'Added by Lydia 2015/04/28 +本所案號
   strSql = "select substrb(' '||sqldatet(FM01),-9) AS FM01, FM02,DECODE(FM03,NULL,NULL,FM03||'-'||FM04||'-'||FM05||'-'||FM06) CasNO,FM03,FM04,FM05,FM06" & _
             " from FagentMail where " & strSQL1 & " ORDER BY FM01,FM02"
   '2008/9/10 end
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   
   If rsQuery.RecordCount > 0 Then
      QueryData = True
      Call UpdateGridList(rsQuery)
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
            
End Function

'報表列印
Private Sub PrintData()

   Dim ii As Integer
   
   Page = 1
   PrintTitle
   
   With grdList
      For ii = 1 To .Rows - 1
      
         '收信日期
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 1)
         '序號
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 2)
         '本所案號
         'Printer.CurrentX = PLeft(3)
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 3)
         'Added by Lydia 2015/04/28
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 8)
         'end 2015/04/28
         iPrint = iPrint + 300
         If iPrint > 10000 And ii <> .Rows - 1 Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.NewPage
            Page = Page + 1
            PrintTitle
         End If
        
       Next ii
   End With
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.EndDoc
   
End Sub

Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = PLeft(0) + 1250
   PLeft(2) = PLeft(1) + 1250
   'Added by Lydia 2015/04/28
   'PLeft(3) = PLeft(2)
   PLeft(3) = PLeft(2) + 2500
End Sub

Sub PrintTitle()
   GetPleft
   
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "代理人信件收達管制記錄明細表"

   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "收信日期：" & Format(ChangeTStringToTDateString(Me.txtQry(1).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txtQry(2).Text)

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "收信日期"
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "序號"
  
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   
   'Added by Lydia 2015/04/28
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "寰華案"
   'end 2015/04/28
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   
End Sub

Private Sub cmdQuery_Click(Index As Integer)
   
   If TxtValidate(1) = False Then Exit Sub
   '查詢
   If Index = 0 Then
      grdList.Rows = 1
      If CheckQueryData = True Then
         Screen.MousePointer = vbHourglass
         grdList.MousePointer = flexHourglass
         If QueryData() = False Then
             MsgBox "無資料", vbOKOnly, "查詢資料"
             txtQry(1).SetFocus
         End If
         grdList.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
      End If
   '列印
   Else
      If grdList.Rows > 1 Then
         Screen.MousePointer = vbHourglass
         PrintData
         ShowPrintOk
         Screen.MousePointer = vbDefault
      Else
         ShowNoData
      End If
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
            Call cmdQuery_Click(0)
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
   'Added by Lydia 2015/04/28
   If Left(Pub_StrUserSt03, 1) = "F" Then
      Check1.Value = 1: Check1.Caption = "只顯示寰華案件"
      FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
      strQFMP = FMP2openSQL '寰華案件
   Else
      Check1.Value = 0: Check1.Caption = "含寰華案件"
      FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
      strQFMP = FMP2openSQL '排除寰華案件
   End If
   '電腦中心不設限
   If Pub_StrUserSt03 = "M51" Then
      strQFMP = "": FMP2openSQL = ""
   End If
   strFMP = FMP2openSQL '控制部門可新增的案件
   strQFMP = Replace(strQFMP, "f0.CP01", "FM03") '控制查詢
   strQFMP = Replace(strQFMP, "f0.CP02", "FM04")
   strQFMP = Replace(strQFMP, "f0.CP03", "FM05")
   strQFMP = Replace(strQFMP, "f0.CP04", "FM06")
   'end 2015/04/28
   
   MoveFormToCenter Me
   Me.Show
   Me.SSTab1.Tab = 0
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
'檢查本所案號
Private Function CheckCaseNo() As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset
   
On Error GoTo ErrHnd

   CheckCaseNo = False
   'Added by Lydia 2015/04/28 編輯資料控制+寰華案
   bolCaseMsg = False
   If txtFM(2).Enabled = False And Pub_StrUserSt03 <> "M51" Then
      strSql = "Select PA01,PA09 From Patent Where PA01='" & txtFM(3) & "' AND PA02='" & txtFM(4) & "' AND PA03='" & txtFM(5) & "' AND PA04='" & txtFM(6) & "'" & strFMP
      strSql = Replace(strSql, "f0.CP", "PA")
   Else
      strSql = "Select PA01,PA09 From Patent Where PA01='" & txtFM(3) & "' AND PA02='" & txtFM(4) & "' AND PA03='" & txtFM(5) & "' AND PA04='" & txtFM(6) & "'"
   End If
   strSql = strSql & " Union Select TM01,TM10 From Trademark Where TM01='" & txtFM(3) & "' AND TM02='" & txtFM(4) & "' AND TM03='" & txtFM(5) & "' AND TM04='" & txtFM(6) & "'"
   strSql = strSql & " Union Select LC01,LC15 From Lawcase Where LC01='" & txtFM(3) & "' AND LC02='" & txtFM(4) & "' AND LC03='" & txtFM(5) & "' AND LC04='" & txtFM(6) & "'"
   strSql = strSql & " Union Select HC01,'000' From Hirecase Where HC01='" & txtFM(3) & "' AND HC02='" & txtFM(4) & "' AND HC03='" & txtFM(5) & "' AND HC04='" & txtFM(6) & "'"
   strSql = strSql & " Union Select SP01,SP09 From Servicepractice Where SP01='" & txtFM(3) & "' AND SP02='" & txtFM(4) & "' AND SP03='" & txtFM(5) & "' AND SP04='" & txtFM(6) & "'"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   If rsQuery.RecordCount > 0 Then
      If rsQuery.Fields(1) = "000" Then
         MsgBox "此案為台灣案", vbCritical
         bolCaseMsg = True 'Added by Lydia 2015/04/28
         Call txtFM_GotFocus(3)
      Else
         CheckCaseNo = True
      End If
   'Added by Lydia 2015/04/28
   ElseIf txtFM(2).Enabled = False And (Pub_strUserST05 = "31" Or Pub_strUserST05 = "33") Then
         MsgBox "請輸入寰華案!!", vbCritical
         bolCaseMsg = True 'Added by Lydia 2015/04/28
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   Exit Function
   
ErrHnd:

   MsgBox Err.Description
   
End Function

Private Function TxtValidate(Optional ByVal iTab As Integer = 0) As Boolean
 
   Dim oText As TextBox, bolCancel As Boolean, arrText, oMaskEdBox As MaskEdBox
   
   TxtValidate = False
   bolCancel = False
   
   Select Case iTab
      Case 0
         SSTab1.Tab = 0
         Set arrText = txtFM
         
      Case 1
         Set arrText = txtQry
   End Select
   
   If bolCancel = False Then
      For Each oText In arrText
         If oText.Locked = False Then
            txtFM_Validate oText.Index, bolCancel
            If bolCancel = True Then
               oText.SetFocus
               TextInverse oText
               Exit For
            End If
         End If
      Next
   End If
   
   If bolCancel = False Then TxtValidate = True
   
End Function

Private Function CheckConfirm() As Boolean
   
   CheckConfirm = False
   
   
   Select Case iCurState
   
      '1:新增;
      Case 1
         
         If TxtValidate = False Then Exit Function
         
         '收信日期
         If txtFM(1) = "" Then
             MsgBox "收信日期不可空白！", vbCritical
             txtFM(1).SetFocus
             Call txtFM_GotFocus(1)
             Exit Function
         End If
        
         '沒有打本所案號第一碼
         If txtFM(3) = "" And (txtFM(4) = "" Or txtFM(5) = "" Or txtFM(6) = "") Then
                MsgBox "本所案不可空白！", vbCritical
                txtFM(3).SetFocus
                Call txtFM_GotFocus(3)
                Exit Function
         ElseIf txtFM(3) = "" And (txtFM(4) <> "" Or txtFM(5) <> "" Or txtFM(6) <> "") Then
               MsgBox "本所案號錯誤！", vbCritical
               txtFM(3).SetFocus
               Call txtFM_GotFocus(3)
               Exit Function
         '有打本所案號第一碼
         ElseIf txtFM(3) <> "" Then
               If CheckCaseNo() = False Then
                   'Added by Lydia 2015/04/28
                  'MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                  'txtFM(3).SetFocus
                  'Call txtFM_GotFocus(3)
                  If bolCaseMsg = False Then
                     MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                     txtFM(3).SetFocus
                     Call txtFM_GotFocus(3)
                  Else
                     txtFM(4).SetFocus
                     Call txtFM_GotFocus(4)
                  End If
                  
                  Exit Function
               End If
         End If
      
      '修改
      Case 2
      
         If TxtValidate = False Then Exit Function
      
         If txtFM(1) = "" And txtFM(2) <> "" Then
            MsgBox "收信日期不可空白！", vbCritical
            txtFM(1).SetFocus
            Call txtFM_GotFocus(1)
            Exit Function
         ElseIf txtFM(1) <> "" And txtFM(2) = "" Then
            MsgBox "序號不可空白！", vbCritical
            txtFM(2).SetFocus
            Call txtFM_GotFocus(2)
            Exit Function
         End If
         
         '沒有打本所案號第一碼
         If txtFM(3) = "" And (txtFM(4) = "" Or txtFM(5) = "" Or txtFM(6) = "") Then
                MsgBox "本所案不可空白！", vbCritical
                txtFM(3).SetFocus
                Call txtFM_GotFocus(3)
                Exit Function
         ElseIf txtFM(3) = "" And (txtFM(4) <> "" Or txtFM(5) <> "" Or txtFM(6) <> "") Then
               MsgBox "本所案號錯誤！", vbCritical
               txtFM(3).SetFocus
               Call txtFM_GotFocus(3)
               Exit Function
         '有打本所案號第一碼
         ElseIf txtFM(3) <> "" Then
               If CheckCaseNo() = False Then
                  'Added by Lydia 2015/04/28
                  'MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                  'txtFM(3).SetFocus
                  'Call txtFM_GotFocus(3)
                  If bolCaseMsg = False Then
                     MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                     txtFM(3).SetFocus
                     Call txtFM_GotFocus(3)
                  Else
                     txtFM(4).SetFocus
                     Call txtFM_GotFocus(4)
                  End If
                  
                  Exit Function
               End If
         End If
         
         '查詢
         Case 4
         
            If txtFM(1) = "" And txtFM(2) <> "" Then
               MsgBox "收信日期不可空白！", vbCritical
               txtFM(1).SetFocus
               Call txtFM_GotFocus(1)
               Exit Function
            ElseIf txtFM(1) <> "" And txtFM(2) = "" Then
               MsgBox "序號不可空白！", vbCritical
               txtFM(2).SetFocus
               Call txtFM_GotFocus(2)
               Exit Function
            End If
   End Select
   CheckConfirm = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010512 = Nothing
End Sub

Private Sub grdList_DblClick()

   Dim lRow As Long, lCurRow As Long, iCol As Integer
   
   lCurRow = grdList.row
   '呼叫查詢
   If lCurRow > 0 Then
      If TBar1.Buttons(4).Enabled = True Then
         Call Tbar1_ButtonClick(TBar1.Buttons(4))
         If txtFM(1).Locked = False Then
                            
            txtFM(1) = ChangeTDateStringToTString(grdList.TextMatrix(lCurRow, 1))
            txtFM(2) = grdList.TextMatrix(lCurRow, 2)
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
               For iCol = 1 To .Cols - 1
                   .col = iCol
                   .CellBackColor = &H80000005
                   .CellForeColor = &H80000008
               Next iCol
            End If
         Next lRow
         '反白
         .row = lCurRow
         For iCol = 1 To .Cols - 1
             .col = iCol
             .CellBackColor = &H8000000D
             .CellForeColor = &H80000005
         Next iCol
      End If
      
      m_bNoTab = True
      grdList_DblClick
      m_bNoTab = False
      
   End With
   
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case PreviousTab
      Case 0
         If iCurState = 0 Then txtQry(1).SetFocus
   End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strCon As String

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
                  cur_FM01 = ""
                  cur_FM02 = ""
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
                  Exit Sub
               End If
            '查詢
            Case 4
               cur_FM01 = ChangeTStringToWString(txtFM(1))
               cur_FM02 = txtFM(2)
               
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
               txtFM(1).SetFocus
               Call txtFM_GotFocus(1)
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
               ElseIf cur_FM01 = "" And cur_FM02 = "" Then
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
               cur_FM01 = lst_FM01
               cur_FM02 = lst_FM02
               
               If cur_FM01 = "" And cur_FM02 = "" Then
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

   lst_FM01 = cur_FM01
   lst_FM02 = cur_FM02

End Sub
'清除畫面
Private Sub FormReset(Optional ByVal iTab As Integer = 0)

   Dim oText As TextBox, oLabel As Object, oMaskEdBox As MaskEdBox
   
   Select Case iTab
   
      Case 0
      '頁籤0
         For Each oText In txtFM
            oText.Text = ""
         Next
         
         For Each oLabel In lblDisp
            oLabel.Caption = ""
         Next
         lblFMP.Visible = False 'Added by Lydia 2015/04/28
      Case 1
      '頁籤1
      
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
   'Added by Lydia 2105/04/28
   If CheckAuthority Then
   End If
   
End Sub
'設定文字框
Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)

   Dim oText As TextBox, oLabel As Label, oMaskEdBox As MaskEdBox
   
   Select Case iStatus
      
      Case 0
      '瀏覽
         For Each oText In txtFM
            oText.Enabled = True
            oText.Locked = True
         Next
         
         txtFM(1).SetFocus
      Case 1
      '新增
         SSTab1.Tab = 0
         For Each oText In txtFM
            oText.Text = ""
            oText.Locked = False
            oText.Enabled = True
         Next
         
         Call FormReset(0)
'         txtFM(1).Text = strSrvDate(2)
         txtFM(1).Text = Tem_Date
         txtFM(3).Text = StrStartSystemByNick '2008/8/29 add by sonia 依系統別預設
         
         '序號
         txtFM(2).Enabled = False
         txtFM(1).SetFocus
         Call txtFM_GotFocus(1)
         
         
      Case 2
      '修改
         SSTab1.Tab = 0
         For Each oText In txtFM
            oText.Locked = False
            oText.Enabled = True
         Next
         
         '序號
         txtFM(1).Enabled = False
         txtFM(2).Enabled = False
         txtFM(3).SetFocus
         Call txtFM_GotFocus(3)
         
      Case 4
      '查詢
         If m_bNoTab = False Then
            SSTab1.Tab = 0
         End If
         For Each oText In txtFM
            oText.Locked = False
            oText.Enabled = False
         Next
         
         Call FormReset(0)
         txtFM(1).Enabled = True
         txtFM(2).Enabled = True
         txtFM(1).SetFocus
         
      Case 9
      '無資料
         For Each oText In txtFM
            oText.Enabled = False
            oText.Locked = True
         Next
         
         Call FormReset(0)
   End Select
   
End Sub
'讀取資料
Private Function doQuery(ByVal iAct As Integer, Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   Dim strSQL1 As String 'Added by Lydia 2015/04/28
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery = False
   
'   'Added by Lydia 2015/04/28 +寰華案
'   '外專人員+勾選只顯示寰華案件 ,非外專人員+不勾選含寰華案件 ,
'   If (Left(Pub_StrUserSt03, 1) = "F" And Check1.Value = 1) Or (Left(Pub_StrUserSt03, 1) <> "F" And Check1.Value = 0) Then
'      strSQL1 = strSQL1 + strQFMP
'   End If
   
   'Modified by Lydia 2015/04/28 條件+寰華案
   Select Case iAct
      Case 4
      '查詢
         strSql = "Select FM01,FM02 From FagentMail where FM01='" & cur_FM01 & "' AND FM02='" & cur_FM02 & "'" & strSQL1
         stMessage = "查無資料！"
   
      Case 6
      '第一筆
         If Len(strSQL1) > 0 Then strSQL1 = "where" + Mid(strSQL1, 5, Len(strSQL1) - 4)
         strSql = "Select FM01,FM02 from FagentMail " & strSQL1 & _
                  " ORDER BY 1 ASC"
         stMessage = "無收信記錄！"
      Case 7
      '上一筆
         strSql = "Select FM01,FM02 From FagentMail where FM01||FM02<'" & cur_FM01 & cur_FM02 & "'" & strSQL1 & _
            " ORDER BY 1 DESC"
         stMessage = "已是第一筆了！"

      Case 8
      '下一筆
         strSql = "Select FM01,FM02 From FagentMail where FM01||FM02>'" & cur_FM01 & cur_FM02 & "'" & strSQL1 & _
            " ORDER BY 1 ASC"
         stMessage = "已是最後一筆了！"

      Case 9
      '最後筆
         If Len(strSQL1) > 0 Then strSQL1 = "where" + Mid(strSQL1, 5, Len(strSQL1) - 4)
         strSql = "Select FM01,FM02 From FagentMail " & strSQL1 & _
            " ORDER BY 1 DESC"
         stMessage = "無信件紀錄！"
        
   End Select
   
On Error GoTo ErrHand

   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   
   If rsQuery.RecordCount > 0 Then
      lst_FM01 = cur_FM01
      lst_FM02 = cur_FM02
      
      cur_FM01 = "" & rsQuery.Fields(0).Value
      cur_FM02 = "" & rsQuery.Fields(1).Value
      
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

Private Sub txtFM_GotFocus(Index As Integer)
   
   If txtFM(Index).Locked = False Then
      TextInverse txtFM(Index)
   End If
   
End Sub
'完整資料查詢
Private Function ReQuery(Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, intI As Integer
   
   Dim strTemp As String
   
On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   
   ReQuery = False
   'Added by Lydia 2015/04/28 +Create ID所屬部門
   strSql = "SELECT FM01-19110000 AS FM01,FM02,FM03,FM04,FM05,FM06" & _
            ",A.ST02 AS D01,FM08 as D02,FM09 as D03,B.ST02 as D04,FM11 as D05,FM12 as D06,a.st03 " & _
         " From FagentMail, STAFF A ,STAFF B Where A.ST01(+)=FM07  and B.ST01(+)=FM10 and FM01='" & cur_FM01 & "' AND FM02='" & cur_FM02 & "'"

   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   
   If rsQuery.RecordCount > 0 Then
      txtFM(1) = cur_FM01
      txtFM(2) = cur_FM02
      
      '曰期 序號 本所案號
      For intI = 1 To 6
         txtFM(intI) = "" & rsQuery.Fields("FM" & Format(intI, "00"))
      Next intI
      
     'Added by Lydia 2015/04/28 +寰華案
     If PUB_FMPtoCheck(1, 1, "31", txtFM(3), txtFM(4), txtFM(5), txtFM(6)) = True Then
        lblFMP.Visible = True
     Else
        lblFMP.Visible = False
     End If
     'end 2015/04/28
     
     lblDisp(1) = "" & rsQuery.Fields("D01")
            
     strTemp = ChangeWStringToTString(rsQuery.Fields("D02"))
     lblDisp(2) = Format(strTemp, "###/##/##")
   
     strTemp = rsQuery.Fields("D03")
     lblDisp(3) = Format(strTemp, "##:##")
     
     'Added by Lydia 2015/04/28 +Create ID所屬部門
     CIDst03 = "" & rsQuery.Fields("st03")
     
    'For UpdateFiled
    If rsQuery.Fields("D04") <> "" And rsQuery.Fields("D05") <> "" And rsQuery.Fields("D06") <> "" Then
    
      lblDisp(4) = "" & rsQuery.Fields("D04")
             
      strTemp = ChangeWStringToTString(rsQuery.Fields("D05"))
      lblDisp(5) = Format(strTemp, "###/##/##")
    
      strTemp = rsQuery.Fields("D06")
      lblDisp(6) = Format(strTemp, "##:##")
     Else
        lblDisp(4) = ""
        lblDisp(5) = ""
        lblDisp(6) = ""
      End If
      ReQuery = True
      
   ElseIf bolMsg Then
      MsgBox "收信曰期〔" & cur_FM01 & "〕收信序號〔" & cur_FM02 & "〕已被刪除！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Screen.MousePointer = vbDefault
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
   Screen.MousePointer = vbDefault
   
End Function

Private Sub txtFM_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtFM(Index).Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
      
         Case 1, 2
         '收信日期, 序號
         If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
         
         Case 3
         '本所案號:只可為字母
            If Not (KeyAscii = 8 Or (KeyAscii > 64 And KeyAscii < 91)) Then
               KeyAscii = 0
            End If
         Case 4, 5, 6
         '本所案號:只可為數字
            If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
      
      End Select
   End If
End Sub

Private Sub txtFM_LostFocus(Index As Integer)
   If SSTab1.Tab = 1 Then
      Dim bolCancel As Boolean
      bolCancel = False
      Call txtFM_Validate(Index, bolCancel)
      If bolCancel = True Then
         SSTab1.Tab = 0
         txtFM(Index).SetFocus
      End If
   End If
End Sub

Private Sub txtFM_Validate(Index As Integer, Cancel As Boolean)

   If txtFM(Index).Locked = False Then
      Select Case Index
         Case 1
            If PUB_CheckKeyInDate(txtFM(Index)) <> 0 Then Cancel = True
               '收信日期
               '2008/8/28 add by Toni 自己入收信日期
               'Modify by Morgan 2010/8/11 百年蟲
               'If txtFM(Index) > strSrvDate(2) Then
               If Val(txtFM(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "收信日期不可大於系統日期"
                  'txtFM(1).SetFocus
                  Cancel = True
               End If
'            End If
         '序號
         Case 2
            If txtFM(Index) <> "" Then
               txtFM(Index) = UCase(Right("000000000" & txtFM(Index).Text, 2))
            End If
         Case 3
         '本所案號
            txtFM(Index) = Trim(txtFM(Index))
            'Added by Lydia 2015/04/28
            'If CheckSysKind(txtFM(Index)) = False Then
            If Not IsCorrectSysKind(txtFM(Index)) Then
               MsgBox "系統代碼輸入錯誤！", vbCritical
               Cancel = True
            End If

         Case 4
         '本所案號
            If txtFM(3) <> "" Then
               txtFM(Index) = UCase(Right("000000" & txtFM(Index).Text, 6))
            End If
         Case 5
         '本所案號
            If txtFM(3) <> "" Then
               txtFM(Index) = UCase(Right("0" & txtFM(Index).Text, 1))
            End If
         Case 6
         '本所案號
            If txtFM(3) <> "" Then
               txtFM(Index) = UCase(Right("00" & txtFM(Index).Text, 2))
            End If
      End Select
   End If
   If Cancel = True Then Call txtFM_GotFocus(Index)
End Sub
'Remove by Lydia 2015/04/28 用現存共用模組 IsCorrectSysKind
'Private Function CheckSysKind(ByVal stSys As String, Optional ByVal bolMsg As Boolean) As Boolean
'
'   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
'   Dim i As Integer, arrSys() As String
'
'On Error GoTo ErrHand
'
'   CheckSysKind = False
'   strSql = "Select SK01 from systemkind"
'   rsQuery.CursorLocation = adUseClient
'   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
'
'   If rsQuery.RecordCount > 0 Then
'      arrSys = Split(rsQuery.GetString, Chr(13))
'      For i = 0 To UBound(arrSys)
'         If arrSys(i) = stSys Then
'            CheckSysKind = True
'            Exit For
'         End If
'      Next i
'   Else
'      If bolMsg Then
'         MsgBox "無法取得系統代碼！", vbCritical
'      End If
'   End If
'
'   If rsQuery.State <> adStateClosed Then rsQuery.Close
'   Set rsQuery = Nothing
'
'   Exit Function
'
'ErrHand:
'
'   MsgBox Err.Description, vbCritical
'
'End Function

Private Function DeleteData() As Boolean
   Dim strSql As String, lngEffRec As Long
   
   strSql = "Delete Fagentmail Where FM01='" & cur_FM01 & "' and FM02='" & cur_FM02 & "'"
   
   DeleteData = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   DeleteData = True
   
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function UpdateData() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, fm(1 To 9) As String
   Dim rsQuery As New ADODB.Recordset, strUpdSQL As String, lngEffRec As Long
   
   fm(3) = "FM03='" & txtFM(3).Text & "'"
   fm(4) = "FM04='" & txtFM(4).Text & "'"
   fm(5) = "FM05='" & txtFM(5).Text & "'"
   fm(6) = "FM06='" & txtFM(6).Text & "'"
   
   fm(7) = "FM10='" & strUserNum & "'"
   fm(8) = "FM11=TO_NUMBER(TO_CHAR(SYSDATE,'YYYYMMDD'))"
   fm(9) = "FM12=TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   
   strSql = "Update FagentMail Set " & fm(3) & "," & fm(4) & "," & fm(5) & "," & fm(6) & "," & fm(7) & "," & fm(8) & "," & fm(9) & _
                                                        " Where FM01='" & cur_FM01 & "' and FM02='" & cur_FM02 & "'"
   
   UpdateData = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   UpdateData = True
   
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function insertdata() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, fm(1 To 9) As String
   Dim strCols As String, strValues As String, lngEffRec As Long
   Dim rsQuery As New ADODB.Recordset
   
   strCols = "FM01"
   For intI = 2 To 9
      strCols = strCols & ",FM" & Format(intI, "00")
   Next intI
   
   fm(1) = Val(txtFM(1).Text) + 19110000              '收信日期
   cur_FM01 = fm(1)
   Tem_Date = fm(1) - 19110000
   fm(3) = "'" & txtFM(3).Text & "'"                  '本所NO
   fm(4) = "'" & txtFM(4).Text & "'"
   fm(5) = "'" & txtFM(5).Text & "'"
   fm(6) = "'" & txtFM(6).Text & "'"
      
   fm(7) = "'" & strUserNum & "'"
   fm(8) = "TO_NUMBER(TO_CHAR(SYSDATE,'YYYYMMDD'))"
   fm(9) = "TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   
   '2008/8/28 add by Toni 自己入收信日期
   'Modify by Morgan 2010/8/11 百年蟲
   'If txtFM(1) > strSrvDate(2) Then
   If Val(txtFM(1)) > Val(strSrvDate(2)) Then
      MsgBox "收信日期大於系統日期"
      txtFM(1).SetFocus
      Exit Function
   End If
    
   '讀取序號
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   'Modified by Morgan 2017/6/1 win7 沒資料會錯
   'strSql = "SELECT max(FM02) FROM Fagentmail where FM01='" & ChangeTStringToWString(txtFM(1)) & "'"
   strSql = "SELECT nvl(max(FM02),0) FROM Fagentmail where FM01='" & ChangeTStringToWString(txtFM(1)) & "'"
   'end 2017/6/1
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
 
   If Not (rsQuery.BOF And rsQuery.EOF) Then
      fm(2) = CNULL(Format(rsQuery.Fields(0).Value + 1, "00"))
      cur_FM02 = Format(rsQuery.Fields(0).Value + 1, "00")
   Else
      fm(2) = CNULL("01")
      cur_FM02 = Format(1, "00")
   End If
  
   strValues = Join(fm, ",")
   
  strSql = "INSERT INTO FagentMail (" & strCols & ") VALUES(" & strValues & ")"
   insertdata = False

On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   insertdata = True
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtQry_GotFocus(Index As Integer)
   If txtQry(Index).Locked = False Then
      TextInverse txtQry(Index)
      If txtQry(Index).Locked = False Then
         CloseIme
      End If
   End If
End Sub

Private Sub txtQry_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtQry(Index).Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
         Case 1, 2
         '收信日期:只可為數字
            If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
      End Select
   End If
End Sub

Private Sub txtQry_LostFocus(Index As Integer)
   If SSTab1.Tab = 0 Then
      Dim bolCancel As Boolean
      bolCancel = False
      Call txtQry_Validate(Index, bolCancel)
      If bolCancel = True Then
         SSTab1.Tab = 1
         txtQry(Index).SetFocus
      End If
   End If
End Sub

Private Sub txtQry_Validate(Index As Integer, Cancel As Boolean)
   If txtQry(Index).Locked = False Then
      Select Case Index
         Case 1
            '收信日期起
            If PUB_CheckKeyInDate(txtQry(Index)) <> 0 Then Cancel = True
         Case 2
            '收信日期迄
            If PUB_CheckKeyInDate(txtQry(Index)) = 0 Then
               'Modify by Morgan 2010/8/11 百年蟲
               'If txtQry(1) <> "" And (txtQry(2) < txtQry(1)) Then
               If txtQry(1) <> "" And (Val(txtQry(2)) < Val(txtQry(1))) Then
                  MsgBox "收信日期迄日必需大於起日！", vbCritical
                  Cancel = True
               End If
            Else
                Cancel = True
            End If
         Case 4
            '本所案號
            If txtQry(3) <> "" Then
               txtQry(Index) = UCase(Right("000000" & txtQry(Index).Text, 6))
            End If
         Case 5
         '本所案號
            If txtQry(3) <> "" Then
               txtQry(Index) = UCase(Right("0" & txtQry(Index).Text, 1))
            End If
         Case 6
         '本所案號
            If txtQry(3) <> "" Then
               txtQry(Index) = UCase(Right("00" & txtQry(Index).Text, 2))
            End If
      End Select
      If Cancel = True Then txtQry_GotFocus (Index)
   End If
End Sub

'Added by Lydia 2015/04/28 控制資料維護權限
Private Function CheckAuthority() As Boolean
   CheckAuthority = False
   If Pub_StrUserSt03 = "M51" Then
      CheckAuthority = True
   Else
    '依CREATE ID的部門及操作者的部門檢查，不可修改或刪除非相同部門輸入的資料
       If Pub_StrUserSt03 = CIDst03 Then
          If Left(Pub_StrUserSt03, 1) = "F" Then
             '開放員工等級為'31'、'33'的外專程序組可輸入寰華案件之信函管理
             'modify by sonia 2016/6/30 加等級32
             If InStr("31,32,33", Pub_strUserST05) > 0 Then CheckAuthority = True
          Else
              CheckAuthority = True
          End If
       End If
   End If
   '控制新增,修改,刪除
   If TBar1.Buttons(11).Enabled = False Then
      If CheckAuthority = True Then
         TBar1.Buttons(2).Enabled = True
         TBar1.Buttons(3).Enabled = True
      Else
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
      End If
   End If
   If Left(Pub_StrUserSt03, 1) = "F" And InStr("31,32,33", Pub_strUserST05) = 0 Then
      TBar1.Buttons(1).Enabled = False
   End If
   Exit Function
End Function

