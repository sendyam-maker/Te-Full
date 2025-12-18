VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210110 
   BorderStyle     =   1  '單線固定
   Caption         =   "行事曆"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9390
   Begin TabDlg.SSTab SSTab1 
      Height          =   4635
      Left            =   30
      TabIndex        =   2
      Top             =   675
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm210110.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(55)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(155)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSupName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSS(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtSS(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtSS(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtSS(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm210110.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdList"
      Tab(1).Control(1)=   "cmdQuery"
      Tab(1).Control(2)=   "txtSS(4)"
      Tab(1).Control(3)=   "txtSS(5)"
      Tab(1).Control(4)=   "Label1(4)"
      Tab(1).Control(5)=   "Line2"
      Tab(1).ControlCount=   6
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3540
         Left            =   -74895
         TabIndex        =   15
         Top             =   1035
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6244
         _Version        =   393216
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
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Left            =   -66750
         TabIndex        =   7
         Top             =   510
         Width           =   912
      End
      Begin MSForms.TextBox txtSS 
         Height          =   300
         Index           =   4
         Left            =   -74040
         TabIndex        =   5
         Top             =   510
         Width           =   945
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSS 
         Height          =   300
         Index           =   5
         Left            =   -72990
         TabIndex        =   6
         Top             =   510
         Width           =   945
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSS 
         Height          =   3240
         Index           =   3
         Left            =   1350
         TabIndex        =   1
         Top             =   1275
         Width           =   7875
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13891;5715"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSS 
         Height          =   300
         Index           =   2
         Left            =   4650
         TabIndex        =   4
         Top             =   570
         Width           =   525
         VariousPropertyBits=   671107097
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSS 
         Height          =   300
         Index           =   1
         Left            =   1365
         TabIndex        =   0
         Top             =   570
         Width           =   945
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSS 
         Height          =   300
         Index           =   0
         Left            =   1350
         TabIndex        =   3
         Top             =   930
         Width           =   945
         VariousPropertyBits=   671107097
         MaxLength       =   6
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日期:"
         Height          =   180
         Index           =   4
         Left            =   -74850
         TabIndex        =   13
         Top             =   540
         Width           =   405
      End
      Begin VB.Line Line2 
         X1              =   -73260
         X2              =   -72840
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備忘錄："
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   12
         Top             =   1350
         Width           =   720
      End
      Begin MSForms.Label lblSupName 
         Height          =   255
         Left            =   2430
         TabIndex        =   11
         Top             =   990
         Width           =   2415
         VariousPropertyBits=   27
         Size            =   "4260;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工："
         Height          =   180
         Index           =   155
         Left            =   510
         TabIndex        =   10
         Top             =   990
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Index           =   55
         Left            =   510
         TabIndex        =   9
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "序號:"
         Height          =   180
         Index           =   6
         Left            =   3810
         TabIndex        =   8
         Top             =   630
         Width           =   405
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frm210110.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":084C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":0E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210110.frx":1E10
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm210110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/20 改成Form2.0 (grdList,lblSupName,txtSS)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Memo by Lydia 2019/07/01 表單名稱:個人行事曆維護=>行事曆
'add by nickc 2005/08/16
Option Explicit

Dim ss(1 To 4) As String
Dim strRsStart1 As String, strRsStart2 As String, strRsStart3 As String, strRsEnd1 As String, strRsEnd2 As String, strRsEnd3 As String
Dim rsDefineSize As New ADODB.Recordset
Dim intWhere As Integer
Dim ActionEdit As Integer
Dim intRow As Integer
Dim m_CurrSel As Integer
Dim StrSQLa As String
Dim m_SalesList As String 'Added by Lydia 2019/08/07 增加權限內可看的人員(ex.杜燕文要看同部門離職人員的記錄)

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyF2
                  RsSitu 0
               Case vbKeyF3
                  RsSitu 1
               Case vbKeyF5
                  RsSitu 2
               Case vbKeyF4
                  RsSitu 5
            End Select
            KeyCode = 0
         End If
      Case vbKeyF9, vbKeyF10, vbKeyReturn
         If ActionEdit <> 3 Then
            Select Case KeyCode
               Case vbKeyF9, vbKeyReturn
                  RsSitu 3
               Case vbKeyF10
                  RsSitu 4
            End Select
            KeyCode = 0
         End If
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyHome
                  RsAction 0
               Case vbKeyPageUp
                  RsAction 1
               Case vbKeyPageDown
                  RsAction 2
               Case vbKeyEnd
                  RsAction 3
            End Select
            KeyCode = 0
         End If
    Case vbKeyEscape
        If MsgBox("是否確定結束?", vbYesNo + vbCritical) = vbYes Then Unload Me
    Case Else
        Exit Sub
    End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
 
    MoveFormToCenter Me
    'Modified by Lydia 2019/08/07 改模組
'    '2009/5/22 add by sonia
'    If GetStaffDepartment(strUserNum) = "M51" Then
'       strExc(0) = "SELECT * FROM staff_schedule where ROWNUM<1"
'    Else
'    '2009/5/22 end
'       strExc(0) = "SELECT * FROM staff_schedule where ss01='" & strUserNum & "' and ROWNUM<1"
'    End If
'    intI = 1
'    Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'   strRsStart1 = Empty: strRsStart2 = Empty: strRsStart3 = Empty
'   strRsEnd1 = Empty: strRsEnd2 = Empty: strRsEnd3 = Empty
'    '2009/5/22 add by sonia
'    If GetStaffDepartment(strUserNum) = "M51" Then
'       strExc(0) = "SELECT ss01,ss02,ss03 FROM staff_schedule Order By ss01,ss02,ss03 "
'    Else
'    '2009/5/22 end
'      strExc(0) = "SELECT ss01,ss02,ss03 FROM staff_schedule where ss01='" & strUserNum & "' Order By ss01,ss02,ss03 "
'    End If
'   intI = 1
'   'edit by nickc 2007/02/05 不用 dll 了
'   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
'   If intI = 1 Then
'        RsTemp.MoveFirst
'      strRsStart1 = "" & RsTemp.Fields(0).Value
'      strRsStart2 = "" & RsTemp.Fields(1).Value
'      strRsStart3 = Format("" & RsTemp.Fields(2).Value, "0000")
'        RsTemp.MoveLast
'      strRsEnd1 = "" & RsTemp.Fields(0).Value
'      strRsEnd2 = "" & RsTemp.Fields(1).Value
'      strRsEnd3 = Format("" & RsTemp.Fields(2).Value, "0000")
'      RsAction 0
'   End If
   m_SalesList = PUB_GetSalesList(strUserNum)
   Call GetLimitArea
   RsAction 0
   'end 2019/08/07
   ActionEdit = 3
   CmdSitu True

   TxtLock 3
   InitialGridList
   SSTab1.Tab = 0
   Me.cmdQuery.Default = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210110 = Nothing
End Sub

Private Sub RsSitu(ByVal Situ As Integer)
Dim i As Integer, St1 As String, St2 As String
Dim TBmk As Variant
Dim StrSQLa As String
 
 '911106 nick
 On Error GoTo CheckingErr
 
 Static TmpSS(4) As String
   Select Case Situ
      Case 0 '按下新增add
        TmpSS(1) = Me.txtSS(0).Text
        TmpSS(2) = ChangeTStringToWString(Me.txtSS(1).Text)
        TmpSS(3) = Me.txtSS(2).Text
         CmdSitu False
         TxtLock 0
         ActionEdit = 0
         Me.txtSS(1).SetFocus
        txtSS_GotFocus 1
        Me.txtSS(0).Text = strUserNum
        Me.lblSupName.Caption = GetStaffName(Me.txtSS(0).Text)
      Case 1 '按下修改modi
         CmdSitu False
         TxtLock 1
         ActionEdit = 1
        TmpSS(1) = Me.txtSS(0).Text
        TmpSS(2) = ChangeTStringToWString(Me.txtSS(1).Text)
        TmpSS(3) = Me.txtSS(2).Text
      Case 2 '按下刪除delete
        If Me.txtSS(0).Text = "" Then
            MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
            Exit Sub
        End If
        If DelMsg Then
            StrSQLa = "Delete From staff_schedule Where Ss01=" & CNULL(Me.txtSS(0).Text) & " And Ss02='" & ChangeTStringToWString(Me.txtSS(1).Text) & "' And Ss03='" & Me.txtSS(2).Text & "'  "
            cnnConnection.Execute StrSQLa
            '2008/9/16 MODIFY BY SONIA
            'strExc(0) = "SELECT Ss01, Ss02, Ss03 FROM staff_schedule WHERE SS01||SS02||SS03>='" & ss(1) & ss(2) & ss(3) & "' Order By SS01, SS02, SS03 "
            '2009/5/22 add by sonia
            'Modifeid by Lydia 2019/08/07 全部移到第1筆
'            If GetStaffDepartment(strUserNum) = "M51" Then
'               strExc(0) = "SELECT Ss01, Ss02, Ss03 FROM staff_schedule WHERE Ss01||SS02||SS03>='" & CNULL(Me.txtSS(0).Text) & ss(2) & ss(3) & "' Order By SS01, SS02, SS03 "
'            Else
'            '2009/5/22 end
'               strExc(0) = "SELECT Ss01, Ss02, Ss03 FROM staff_schedule WHERE Ss01=" & CNULL(Me.txtSS(0).Text) & " And SS02||SS03>='" & ss(2) & ss(3) & "' Order By SS01, SS02, SS03 "
'            End If
'            '2008/9/16 END
'             intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(1) = "" & RsTemp.Fields(0).Value
'               strExc(2) = "" & RsTemp.Fields(1).Value
'               strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
'               ReadStaff_Schedule strExc
'            Else
'                '2008/9/16 MODIFY BY SONIA
'                'strExc(0) = "SELECT SS01, SS02, SS03 FROM staff_schedule WHERE SS01||SS02||SS03<='" & ss(1) & ss(2) & ss(3) & "' Order By SS01 Desc , SS02 Desc, SS03 Desc"
'               '2009/5/22 add by sonia
'               If GetStaffDepartment(strUserNum) = "M51" Then
'                  strExc(0) = "SELECT Ss01, Ss02, Ss03 FROM staff_schedule WHERE Ss01||SS02||SS03<='" & CNULL(Me.txtSS(0).Text) & ss(2) & ss(3) & "' Order By SS01 Desc , SS02 Desc, SS03 Desc"
'               Else
'               '2009/5/22 end
'                strExc(0) = "SELECT SS01, SS02, SS03 FROM staff_schedule WHERE Ss01=" & CNULL(Me.txtSS(0).Text) & " And SS02||SS03<='" & ss(2) & ss(3) & "' Order By SS01 Desc , SS02 Desc, SS03 Desc"
'               End If
'                '2008/9/16 END
'                 intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                If intI = 1 Then
'                   strExc(1) = "" & RsTemp.Fields(0).Value
'                   strExc(2) = "" & RsTemp.Fields(1).Value
'                   strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
'                   ReadStaff_Schedule strExc
'                Else
'                   RsAction 0
'                End If
'            End If
'            'Modify By Sindy 2009/07/23
'            'strExc(0) = "SELECT SS01, SS02, SS03 FROM staff_schedule Order By SS01, SS02, SS03 "
'            If GetStaffDepartment(strUserNum) = "M51" Then
'               strExc(0) = "SELECT ss01,ss02,ss03 FROM staff_schedule Order By ss01,ss02,ss03 "
'            Else
'               strExc(0) = "SELECT ss01,ss02,ss03 FROM staff_schedule where ss01='" & strUserNum & "' Order By ss01,ss02,ss03 "
'            End If
'            '2009/07/23 End
'            intI = 1
'            'edit by nickc 2007/02/05 不用 dll 了
'            'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
'            If intI = 1 Then
'                 RsTemp.MoveFirst
'               strRsStart1 = "" & RsTemp.Fields(0).Value
'               strRsStart2 = "" & RsTemp.Fields(1).Value
'               strRsStart3 = Format("" & RsTemp.Fields(2).Value, "0000")
'                 RsTemp.MoveLast
'               strRsEnd1 = "" & RsTemp.Fields(0).Value
'               strRsEnd2 = "" & RsTemp.Fields(1).Value
'               strRsEnd3 = Format("" & RsTemp.Fields(2).Value, "0000")
'            End If
             Call GetLimitArea
             RsAction 0
             'end 2019/08/07
        End If
      Case 3 'update
         If ActionEdit = 0 Then '在新增狀態按Enter鍵
            If Not GetData Then Exit Sub
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            Me.txtSS(2).Text = GetSerialNo(Me.txtSS(0).Text, ChangeTStringToWString(Me.txtSS(1).Text))
            StrSQLa = "Insert Into staff_schedule (Ss01, Ss02, Ss03, Ss04) values ('" & Me.txtSS(0).Text & "'," & ChangeTStringToWString(Me.txtSS(1).Text) & "," & Me.txtSS(2).Text & ",'" & ChgSQL(Me.txtSS(3).Text) & "')"
            cnnConnection.Execute StrSQLa
            'Modifeid by Lydia 2019/08/07 改模組
'            If strRsStart1 & strRsStart2 & strRsStart3 & strRsEnd1 & strRsEnd2 & strRsEnd3 = "" Then
'                   strRsStart1 = Me.txtSS(0).Text
'                   strRsStart2 = ChangeTStringToWString(Me.txtSS(1).Text)
'                   strRsStart3 = Me.txtSS(2).Text
'                   strRsEnd1 = Me.txtSS(0).Text
'                   strRsEnd2 = ChangeTStringToWString(Me.txtSS(1).Text)
'                   strRsEnd3 = Me.txtSS(2).Text
'            Else
'               If Me.txtSS(0).Text & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text < strRsStart1 & strRsStart2 & strRsStart3 Then
'                   strRsStart1 = Me.txtSS(0).Text
'                   strRsStart2 = ChangeTStringToWString(Me.txtSS(1).Text)
'                   strRsStart3 = Me.txtSS(2).Text
'               End If
'               If Me.txtSS(0).Text & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text > strRsEnd1 & strRsEnd2 & strRsEnd3 Then
'                   strRsEnd1 = Me.txtSS(0).Text
'                   strRsEnd2 = ChangeTStringToWString(Me.txtSS(1).Text)
'                   strRsEnd3 = Me.txtSS(2).Text
'               End If
'            End If
            Call GetLimitArea
            'end 2018/08/06
            strExc(1) = Me.txtSS(0).Text
            strExc(2) = ChangeTStringToWString(Me.txtSS(1).Text)
            strExc(3) = Me.txtSS(2).Text
            ReadStaff_Schedule strExc
         ElseIf ActionEdit = 1 Then '在修改狀態按Enter鍵
            If Not GetData Then Exit Sub
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            StrSQLa = ""
            If Me.txtSS(3).Text <> Me.txtSS(3).Tag Then
                StrSQLa = StrSQLa & " SS04='" & ChgSQL(Me.txtSS(3).Text) & "',"
            End If
            If StrSQLa <> "" Then
                StrSQLa = Left(StrSQLa, Len(StrSQLa) - 1)
            Else
                GoTo NoUpdate
            End If
            StrSQLa = "Update staff_schedule Set " & StrSQLa & " Where SS01='" & Me.txtSS(0).Text & "' And SS02=" & Val(ChangeTStringToWString(Me.txtSS(1).Text)) & " And SS03='" & Me.txtSS(2).Text & "' "
            cnnConnection.Execute StrSQLa
NoUpdate:
            '2009/5/22 add by sonia
            'Modifeid by Lydia 2019/08/07 改模組
'            If GetStaffDepartment(strUserNum) = "M51" Then
'               strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule Order By SS01, SS02, SS03"
'            Else
'            '2009/5/22 end
'               strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule where SS01='" & strUserNum & "'  Order By SS01, SS02, SS03"
'            End If
'            intI = 1
'            'edit by nickc 2007/02/05 不用 dll 了
'            'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
'            If intI = 1 Then
'                 RsTemp.MoveFirst
'               strRsStart1 = "" & RsTemp.Fields(0).Value
'               strRsStart2 = "" & RsTemp.Fields(1).Value
'               strRsStart3 = Format("" & RsTemp.Fields(2).Value, "0000")
'                 RsTemp.MoveLast
'               strRsEnd1 = "" & RsTemp.Fields(0).Value
'               strRsEnd2 = "" & RsTemp.Fields(1).Value
'               strRsEnd3 = Format("" & RsTemp.Fields(2).Value, "0000")
'            End If
            Call GetLimitArea
            'end 2019/08/07
            strExc(1) = Me.txtSS(0).Text
            strExc(2) = ChangeTStringToWString(Me.txtSS(1).Text)
            strExc(3) = Me.txtSS(2).Text
            ReadStaff_Schedule strExc
         ElseIf ActionEdit = 2 Then '在查詢狀態按下Enter鍵
            If Me.txtSS(1).Text = "" Then
               MsgBox "日期不可空白，請重新輸入 !", vbCritical
               Me.txtSS(1).SetFocus
               txtSS_GotFocus 1
               Exit Sub
            End If
            intI = 1
            strExc(0) = "SELECT COUNT(*) FROM Staff_schedule WHERE SS01='" & Me.txtSS(0).Text & "' And SS02=" & IIf(Me.txtSS(1).Text <> "", "" & ChangeTStringToWString(Me.txtSS(1).Text) & " ", "SS02") & " And SS03= " & IIf(Me.txtSS(2).Text <> "", "'" & Me.txtSS(2).Text & "'", "SS03") & " "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) = 0 Then
                  MsgBox "查無此排程記錄 !", vbCritical
                    strExc(1) = TmpSS(1)
                    strExc(2) = TmpSS(2)
                    strExc(3) = TmpSS(3)
               Else
                    strExc(1) = Me.txtSS(0).Text
                    strExc(2) = ChangeTStringToWString(Me.txtSS(1).Text)
                    strExc(3) = Me.txtSS(2).Text
               End If
            End If
            ReadStaff_Schedule strExc
         End If
         CmdSitu True
         ActionEdit = 3
         TxtLock 3
      Case 4 'cancel
         If ActionEdit <> 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
         End If
         CmdSitu True
        If TmpSS(1) = "" Then TmpSS(1) = strRsStart1
        If TmpSS(2) = "" Then TmpSS(2) = strRsStart2
        If TmpSS(3) = "" Then TmpSS(3) = strRsStart3
        strExc(1) = TmpSS(1)
        strExc(2) = TmpSS(2)
        strExc(3) = TmpSS(3)
         ActionEdit = 3
         ReadStaff_Schedule strExc
         TxtLock 3
      Case 5 'query
        TmpSS(1) = Me.txtSS(0).Text
        TmpSS(2) = ChangeTStringToWString(Me.txtSS(1).Text)
        TmpSS(3) = Me.txtSS(2).Text
         CmdSitu False
         TxtLock 2
         ActionEdit = 2
         Me.txtSS(1).SetFocus
         txtSS_GotFocus 1
   End Select
   Exit Sub
CheckingErr:
    MsgBox Err.Description
End Sub

Private Sub RsAction(ByVal Sty As Integer)
Dim i As Integer
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   strExc(1) = "": strExc(2) = "": strExc(3) = ""    'Added by Lydia 2019/08/07 清空記錄
   
   Select Case Sty
      Case 0 '第一筆
         'Remove by Lydia 2019/08/07
         'strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule WHERE SS01='" & strRsStart1 & "' And SS02 =" & strRsStart2 & " And SS03= " & strRsStart3 & " "
        ' Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
       '  If intI = 1 Then
        '    strExc(1) = "" & RsTemp.Fields(0).Value
       '     strExc(2) = "" & RsTemp.Fields(1).Value
       '     strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
       ' Else
            'Modified by Lydia 2017/07/25 非電腦中心人員,只能看自己的資料 +  " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & "
            'Modified by Lydia 2019/08/07 增加權限內可看的人員
            'strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & " SS01||SS02||ltrim(to_char(SS03,'0000'))>='" & strRsStart1 & strRsStart2 & strRsStart3 & "' Order By SS01, SS02, SS03 "
            strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule " & _
                             "WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01 in (" & m_SalesList & ") AND") & _
                             " SS02||ltrim(to_char(SS03,'0000'))>='" & strRsStart2 & strRsStart3 & "' " & _
                             " Order By SS02, SS03 "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
                strRsStart1 = strExc(1)
                strRsStart2 = strExc(2)
                strRsStart3 = strExc(3)
            End If
         'End If 'Remove by Lydia 2019/08/07
      Case 1 '前一筆
         'Remove by Lydia 2019/08/07
         'If Me.txtSS(0).Text & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text = strRsStart1 & strRsStart2 & strRsStart3 Then
         '   Beep
         '   Screen.MousePointer = vbDefault
         '   DataErrorMessage 6
         '   Exit Sub
         'Else
            'Modified by Lydia 2017/07/25 非電腦中心人員,只能看自己的資料 + " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & "
            'Modified by Lydia 2019/08/07 增加權限內可看的人員
            'strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & " SS01||SS02||ltrim(to_char(SS03,'0000'))<'" & Me.txtSS(0).Text & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text & "' Order By SS01 Desc, SS02 Desc, SS03 Desc"
            strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule " & _
                             "WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01 in (" & m_SalesList & ") AND") & _
                             " SS02||ltrim(to_char(SS03,'0000'))<'" & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text & "' " & _
                             " Order By SS02 Desc, SS03 Desc"
            'edit by nickc 2007/02/05 不用 dll 了
            'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
            'Added by Lydia 2019/08/07
            Else
                Beep
                Screen.MousePointer = vbDefault
                DataErrorMessage 6
                Exit Sub
            'end 2019/08/07
            End If
         'End If 'Remove by Lydia 2019/08/07
      Case 2 '後一筆
         'Remove by Lydia 2019/08/07
         'If Me.txtSS(0).Text & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text = strRsEnd1 & strRsEnd2 & strRsEnd3 Then
        '    Beep
        '    Screen.MousePointer = vbDefault
        '    DataErrorMessage 7
        '    Exit Sub
        ' Else
            'Modified by Lydia 2017/07/25 非電腦中心人員,只能看自己的資料 + " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & "
            'Modified by Lydia 2019/08/07 增加權限內可看的人員
            'strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & " SS01||SS02||ltrim(to_char(SS03,'0000'))>'" & Me.txtSS(0).Text & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text & "' Order By SS01, SS02, SS03 "
            strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule " & _
                              "WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01 in (" & m_SalesList & ") AND") & _
                              " SS02||ltrim(to_char(SS03,'0000'))>'" & ChangeTStringToWString(Me.txtSS(1).Text) & Me.txtSS(2).Text & "' " & _
                              " Order By SS02, SS03 "
            intI = 1
            'edit by nickc 2007/02/05 不用 dll 了
            'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
            'Added by Lydia 2019/08/07
            Else
                Beep
                Screen.MousePointer = vbDefault
                DataErrorMessage 7
                Exit Sub
            End If
         'End If 'Remove by Lydia 2019/08/07
      Case 3 '最後筆
         'Remove by Lydia 2019/08/07
         'strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule WHERE SS01='" & strRsEnd1 & "' And SS02=" & strRsEnd2 & " And SS03=" & strRsEnd3 & " "
         'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         'If intI = 1 Then
         '   strExc(1) = "" & RsTemp.Fields(0).Value
         '   strExc(2) = "" & RsTemp.Fields(1).Value
         '   strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
        'Else
            'Modified by Lydia 2017/07/25 非電腦中心人員,只能看自己的資料 + " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & "
            'Modified by Lydia 2019/08/07 增加權限內可看的人員
            'strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01='" & strUserNum & "' AND") & " SS01||SS02||ltrim(to_char(SS03,'0000'))<='" & strRsEnd1 & strRsEnd2 & strRsEnd3 & "' Order By SS01 Desc, SS02 Desc, SS03 Desc "
            strExc(0) = "SELECT SS01, SS02, SS03 FROM Staff_schedule " & _
                             "WHERE " & IIf(Pub_StrUserSt03 = "M51", "", "SS01 in (" & m_SalesList & ") AND") & _
                             " SS02||ltrim(to_char(SS03,'0000'))<='" & strRsEnd2 & strRsEnd3 & "' " & _
                             " Order By SS02 Desc, SS03 Desc "
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = Format("" & RsTemp.Fields(2).Value, "0000")
                strRsEnd1 = strExc(1)
                strRsEnd2 = strExc(2)
                strRsEnd3 = strExc(3)
            End If
         'End If　'Remove by Lydia 2019/08/07
   End Select
   ReadStaff_Schedule strExc
   Screen.MousePointer = vbDefault
   Exit Sub
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub CmdSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As TextBox
   If TF = True Then
'      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         If Not IsEmptyText(strRsStart1) And Not IsEmptyText(strRsEnd1) Then
            TBar1.Buttons(i + 5).Enabled = True
         Else
            TBar1.Buttons(i + 5).Enabled = False
         End If
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
'      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
Select Case Lt
Case 0 '新增
    Me.txtSS(0).Locked = True
    Me.txtSS(1).Locked = False
    Me.txtSS(2).Locked = True
    Me.txtSS(0).Text = strUserNum
    Me.txtSS(1).Text = ""
    Me.txtSS(2).Text = ""
    Me.txtSS(3).Text = ""
    Me.txtSS(4).Text = ""
    Me.txtSS(5).Text = ""
    Me.lblSupName.Caption = GetStaffName(Me.txtSS(0).Text)
Case 1 '修改
    Me.txtSS(0).Locked = True
    Me.txtSS(1).Locked = True
    Me.txtSS(2).Locked = True
Case 2 '查詢
    Me.txtSS(0).Locked = False
    Me.txtSS(1).Locked = False
    Me.txtSS(2).Locked = False
    Me.txtSS(0).Text = strUserNum
    Me.txtSS(1).Text = ""
    Me.txtSS(2).Text = ""
    Me.txtSS(3).Text = ""
    Me.txtSS(4).Text = ""
    Me.txtSS(5).Text = ""
    Me.lblSupName.Caption = GetStaffName(Me.txtSS(0).Text)
Case 3 '按下取消後的狀態
    Me.txtSS(0).Locked = True
    Me.txtSS(1).Locked = True
    Me.txtSS(2).Locked = True
End Select
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
Dim nCurrSel As Integer
Dim nCol As Integer
   
    nCurrSel = grdList.row
    ' 與前一選擇的列位置相同則不處理
    If m_CurrSel = grdList.row Then
        GoTo EXITSUB
    End If
    ' 將原先選取的列回復到正常的顏色
    If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
        grdList.row = m_CurrSel
        grdList.col = 1
        If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
                grdList.col = nCol
                If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
        End If
        grdList.col = 0
    End If
    ' 設定成所選取的列
    m_CurrSel = nCurrSel
    ' 將所選取的列反白
    If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
        grdList.row = m_CurrSel
        grdList.col = 1
        For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
        Next nCol
        grdList.col = 0
    End If
EXITSUB:
End Sub

'Add by Morgan 2003/12/26
Private Sub grdList_DblClick()
   'Call grdList_SelChange 'Added by Lydia 2019/08/07
   SSTab1.Tab = 0
End Sub

Private Sub grdList_SelChange()
   Dim nRow As Integer
    grdList_ShowSelection

    If grdList.row > 0 And grdList.row <= grdList.Rows - 1 Then
        nRow = grdList.row
        'Modified by Lydia 2019/08/07
        'strExc(1) = Me.grdList.TextMatrix(nRow, 1)
        strExc(1) = Trim(Mid(Me.grdList.TextMatrix(nRow, 1), 1, InStr(Me.grdList.TextMatrix(nRow, 1), " ")))
        strExc(2) = DBDATE(Me.grdList.TextMatrix(nRow, 2))
        strExc(3) = Format(Me.grdList.TextMatrix(nRow, 3), "0000")
        ReadStaff_Schedule strExc
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
    Select Case Me.SSTab1.Tab
    Case 0
        Me.txtSS(1).SetFocus
        txtSS_GotFocus 1
        Me.cmdQuery.Default = False
    Case 1
        Me.txtSS(4).SetFocus
        txtSS_GotFocus 4
        Me.cmdQuery.Default = True
    End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHand
   Select Case Button.Index
      Case 1 '按下新增
         RsSitu 0
      Case 2 '按下修改
         RsSitu 1
      Case 3 '按下刪除
         RsSitu 2
      Case 4 '按下查詢
         RsSitu 5
      Case 6 '第一筆
         RsAction 0
      Case 7 '前一筆
         RsAction 1
      Case 8 '後一筆
         RsAction 2
      Case 9 '最後筆
         RsAction 3
      Case 11 '按下確定
         RsSitu 3
      Case 12 '按下取消
         RsSitu 4
      Case 14 '結束
         Unload Me
   End Select
   Exit Sub
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

Private Function CheckRule() As Boolean
Dim i As Integer, bolChk As Boolean, j As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   CheckRule = False
   If Me.txtSS(1).Text = "" Then
      MsgBox "日期不可空白 !", vbCritical
      Me.txtSS(1).SetFocus
      txtSS_GotFocus 1
      Exit Function
   End If
   If Me.txtSS(3).Text = "" Then
      MsgBox "備忘錄不可空白 !", vbCritical
      Me.txtSS(3).SetFocus
      txtSS_GotFocus 3
      Exit Function
   End If

   CheckRule = True
End Function

Private Function GetData() As Boolean
Dim i As Integer
    GetData = False
    If CheckRule = False Then Exit Function
    ss(1) = Me.txtSS(0).Text
    ss(2) = ChangeTStringToWString(Me.txtSS(1).Text)
    ss(3) = Me.txtSS(2).Text
    ss(4) = Me.txtSS(3).Text
    GetData = True
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
'Added by Morgan 2022/1/20 檢查畫面輸入欄位是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If
'end 2022/1/20
   
For Each objTxt In Me.txtSS
    If objTxt.Enabled = True Then
       Cancel = False
       txtSS_Validate objTxt.Index, Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
Next
TxtValidate = True
End Function

Private Sub txtSS_GotFocus(Index As Integer)
    TextInverse Me.txtSS(Index)
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSS_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSS_Validate(Index As Integer, Cancel As Boolean)
    If Me.txtSS(Index).Text = "" Then Exit Sub
    If ActionEdit <> 0 And ActionEdit <> 1 Then Exit Sub 'Added by Morgan 2022/1/20
    Select Case Index
    Case 1 '日期
        If CheckIsTaiwanDate(Me.txtSS(Index).Text) = False Then
            Me.txtSS(Index).SetFocus
            txtSS_GotFocus Index
            Cancel = True
            Exit Sub
'edit by nickc 2005/08/24 分所智權人員會有跟客戶約非工作日的情況
'        ElseIf ChkWorkDay(ChangeTStringToWString(Me.txtSS(Index).Text)) = False Then
'            MsgBox "輸入的日期非工作天!!!", vbExclamation + vbOKOnly
'            Me.txtSS(Index).SetFocus
'            txtSS_GotFocus Index
'            Cancel = True
'            Exit Sub
        End If
    Case 4, 5 '支援日期區間
        If CheckIsTaiwanDate(Me.txtSS(Index).Text) = False Then
            Me.txtSS(Index).SetFocus
            txtSS_GotFocus Index
            Cancel = True
            Exit Sub
        End If
        If Me.txtSS(4).Text <> "" And Me.txtSS(5).Text <> "" Then
            If Val(Me.txtSS(4).Text) > Val(Me.txtSS(5).Text) Then
                MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtSS(4).SetFocus
                txtSS_GotFocus 4
                Cancel = True
                Exit Sub
            End If
        End If
    End Select
End Sub

Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim nRow As Integer
   
    QueryData = False
    
    strSql = ""
    
    If Me.txtSS(4).Text <> "" Then
        strSql = strSql & " And Ss02>=" & DBDATE(Me.txtSS(4).Text) & " "
    End If
    If Me.txtSS(5).Text <> "" Then
        strSql = strSql & " And Ss02<=" & DBDATE(Me.txtSS(5).Text) & " "
    End If
    '2009/5/22 add by sonia
    'Modified by Lydia 2019/08/07
    'If GetStaffDepartment(strUserNum) = "M51" Then
    If Pub_StrUserSt03 = "M51" Then
       'Modified by Lydia 2019/08/07 +ST02
       'strSql = "SELECT SS01,sqldatet(SS02), SS03, SS04 FROM Staff_schedule " & _
          "WHERE ss01 IS NOT NULL " & strSql & "Order By  SS02, SS01, SS03, SS04 "
       strSql = "SELECT SS01||' '||ST02 as Sname,sqldatet(SS02), SS03, SS04 FROM Staff_schedule,Staff " & _
          "WHERE ss01 IS NOT NULL " & strSql & " and ss01=st01(+) Order By  SS02, SS01, SS03, SS04 "
    Else
    '2009/5/22 end
         'Modified by Lydia 2019/08/07 增加權限內可看的人員(ex.杜燕文要看同部門離職人員的記錄)
         'strSql = "SELECT SS01,sqldatet(SS02), SS03, SS04 FROM Staff_schedule " & _
                "WHERE ss01='" & strUserNum & "' " & strSql & " Order By SS01, SS02, SS03, SS04 "
         'Modified by Lydia 2019/08/07 +ST02
         'strSql = "SELECT SS01,sqldatet(SS02), SS03, SS04 FROM Staff_schedule " & _
                "WHERE ss01 in (" & m_SalesList & " ) " & strSql & " Order By SS01, SS02, SS03, SS04 "
         strSql = "SELECT SS01||' '||ST02 as Sname,sqldatet(SS02), SS03, SS04 FROM Staff_schedule,Staff " & _
                "WHERE ss01 in (" & m_SalesList & " ) " & strSql & " and ss01=st01(+) Order By SS02, SS01, SS03, SS04 "
    End If
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    grdList.Clear
    grdList.Rows = 2 'Added by Morgan 2022/1/20
    If rsTmp.RecordCount > 0 Then
        QueryData = True
        grdList.Rows = 2
        grdList.Cols = 4
        Set grdList.Recordset = rsTmp
    End If
    InitialGridList
    rsTmp.Close
    Set rsTmp = Nothing
    
    grdList.row = 0    'Added by Morgan 2022/1/20
End Function

' 初始化列表
Public Sub InitialGridList()
    grdList.Cols = 5
    grdList.row = 0
    grdList.ColWidth(0) = 0
    grdList.col = 1
    'Modified by Lydia 2019/08/07
    'grdList.ColWidth(1) = 0
    grdList.Text = "員工"
    grdList.ColWidth(1) = 1300
    'end 2019/08/07
    grdList.ColAlignment(1) = flexAlignLeftCenter
    
    grdList.col = 2
    grdList.Text = "日期"
    grdList.ColWidth(2) = 800
    grdList.ColAlignment(2) = flexAlignLeftCenter
    grdList.col = 3
    grdList.Text = "序號"
    grdList.ColWidth(3) = 500
    grdList.ColAlignment(3) = flexAlignLeftCenter
    grdList.col = 4
    grdList.Text = "備忘錄"
    grdList.ColWidth(4) = 7500
    grdList.ColAlignment(4) = flexAlignLeftCenter
End Sub

'取得序號
Private Function GetSerialNo(strSS01 As String, strSS02 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'Modify By Sindy 2011/3/9
'StrSQLa = "Select max(ss03) as SS03 From staff_schedule Where SS01=" & strSS01 & " And SS02='" & strSS02 & "' "
StrSQLa = "Select NVL(max(ss03),0) as SS03 From staff_schedule Where SS01='" & strSS01 & "' And SS02='" & strSS02 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If Not rsA.EOF And Not rsA.BOF Then
    GetSerialNo = Format(Val(rsA("SS03").Value) + 1, "0000")
Else
    GetSerialNo = "0001"
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

Public Sub SelectToolbarButtom()
Dim btn
    '設定為按下查詢鈕扭
    Set btn = Me.TBar1.Buttons(4)
    Tbar1_ButtonClick btn
End Sub

Private Sub cmdQuery_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
        
'add by nickc 2005/07/06 加入判斷，若頁籤在 基本資料，就不管
If SSTab1.Tab = 0 Then Exit Sub

    If Me.txtSS(4).Text = "" Then
        MsgBox "請輸入日期起日!!!", vbExclamation + vbOKOnly
        Me.txtSS(4).SetFocus
        Exit Sub
    End If
    If CheckIsTaiwanDate(Me.txtSS(4).Text) = False Then
        Me.txtSS(4).SetFocus
        txtSS_GotFocus 4
        Exit Sub
    End If
    If Me.txtSS(5).Text = "" Then
        MsgBox "請輸入支援迄日!!!", vbExclamation + vbOKOnly
        Me.txtSS(5).SetFocus
        Exit Sub
    End If
    If CheckIsTaiwanDate(Me.txtSS(5).Text) = False Then
        Me.txtSS(5).SetFocus
        Exit Sub
    End If
    If Val(Me.txtSS(4).Text) > Val(Me.txtSS(5).Text) Then
        MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txtSS(4).SetFocus
        txtSS_GotFocus 4
        Exit Sub
    End If
    '查詢
      Screen.MousePointer = vbHourglass
      Me.grdList.MousePointer = flexHourglass
      If QueryData() = False Then
          strTit = "查詢資料"
          strMsg = "無資料"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      Me.grdList.MousePointer = flexDefault
      Screen.MousePointer = vbDefault
End Sub

Private Function ReadStaff_Schedule(ByRef tsTmp() As String) As Boolean
Dim i As Integer, j As Integer, Lbl As LABEL, txt As TextBox, strTmp As String
Dim strTxt(0 To 4) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
    strTxt(1) = tsTmp(1): strTxt(2) = tsTmp(2): strTxt(3) = tsTmp(3)
    ss(1) = strTxt(1): ss(2) = strTxt(2): ss(3) = strTxt(3)
    For i = 0 To 3
      Me.txtSS(i).Text = ""
    Next i
    Me.lblSupName.Caption = ""
   If ss(1) = "" Then Exit Function
    StrSQLa = "Select * From staff_schedule Where ss01='" & ss(1) & "' And SS02=" & IIf(ss(2) <> "", "'" & ss(2) & "'", "SS02") & " And SS03=" & IIf(ss(3) <> "", "'" & ss(3) & "'", "SS03") & "  Order By SS01, SS02, SS03 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
        ss(1) = "" & rsA.Fields(0).Value
        ss(2) = "" & rsA.Fields(1).Value
        ss(3) = Format("" & rsA.Fields(2).Value, "0000")
        ss(4) = "" & rsA.Fields(3).Value
    Else
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   Me.txtSS(0).Text = ss(1)
   Me.txtSS(1).Text = ChangeWStringToTString(ss(2))
   Me.lblSupName.Caption = GetStaffName(ss(1), True)
   Me.txtSS(2).Text = ss(3)
   Me.txtSS(3).Text = ss(4)
End Function

'Added by Lydia 2019/08/07 取得最小範圍和最大範圍
Private Sub GetLimitArea()
Dim intQ As Integer
Dim rsB As New ADODB.Recordset
Dim strA1 As String
Dim tmpArr As Variant

    strA1 = "SELECT MIN(SS02||','||LPAD(SS03,4,'0')||','||SS01) MINNO," & _
                "MAX(SS02||','||LPAD(SS03,4,'0')||','||SS01) MAXNO " & _
                "FROM STAFF_SCHEDULE WHERE " & IIf(Pub_StrUserSt03 = "M51", "1=1", "SS01 IN (" & m_SalesList & ") ")
    intQ = 1
    Set rsB = ClsLawReadRstMsg(intQ, strA1)
    strRsStart1 = "": strRsStart2 = "": strRsStart3 = ""
    strRsEnd1 = "": strRsEnd2 = "": strRsEnd3 = ""
    If intQ = 1 Then
        If "" & rsB.Fields("MINNO") <> "" Then
             tmpArr = Split(rsB.Fields("MINNO"), ",")
             strRsStart1 = tmpArr(2)
             strRsStart2 = tmpArr(0)
             strRsStart3 = tmpArr(1)
        End If
        If "" & rsB.Fields("MAXNO") <> "" Then
             tmpArr = Split(rsB.Fields("MAXNO"), ",")
             strRsEnd1 = tmpArr(2)
             strRsEnd2 = tmpArr(0)
             strRsEnd3 = tmpArr(1)
        End If
    End If
    Set rsB = Nothing

End Sub
