VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050711 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人國外ID資料維護"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9150
   Begin TabDlg.SSTab SSTab1 
      Height          =   4932
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   8892
      _ExtentX        =   15690
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050711.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbld1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbld1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm050711.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Combo1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Adodc1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "MFG1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1488
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1008
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   0
         Top             =   528
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   -73740
         MaxLength       =   8
         TabIndex        =   3
         Top             =   408
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查詢"
         Height          =   375
         Left            =   -72000
         TabIndex        =   4
         Top             =   360
         Width           =   1332
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG1 
         Bindings        =   "frm050711.frx":0038
         Height          =   3672
         Left            =   -74820
         TabIndex        =   5
         Top             =   1080
         Width           =   8532
         _ExtentX        =   15055
         _ExtentY        =   6482
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   336
         Left            =   -67440
         Top             =   360
         Visible         =   0   'False
         Width           =   1212
         _ExtentX        =   2143
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
      Begin MSForms.ComboBox Combo1 
         Height          =   300
         Left            =   -73740
         TabIndex        =   15
         Top             =   765
         Width           =   7455
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "13150;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱:"
         Height          =   180
         Index           =   1
         Left            =   -74820
         TabIndex        =   14
         Top             =   768
         Width           =   768
      End
      Begin VB.Label Label5 
         Caption         =   "ID號碼:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1488
         Width           =   975
      End
      Begin MSForms.Label Lbld1 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Top             =   1005
         Width           =   975
         VariousPropertyBits=   27
         Size            =   "1720;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbld1 
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   10
         Top             =   525
         Width           =   2895
         VariousPropertyBits=   27
         Size            =   "5106;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "申請國家:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1008
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "客戶編號:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號:"
         Height          =   180
         Index           =   0
         Left            =   -74820
         TabIndex        =   7
         Top             =   405
         Width           =   765
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8580
      Top             =   600
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
            Picture         =   "frm050711.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050711.frx":1E25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
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
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "frm050711"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/4 改成Form2.0 (MFG1,Combo1,Lbld1)
'Memo By Morgan 2012/12/12 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim Data_Mission As Integer
Dim rs As ADODB.Recordset
Dim strName As String
Dim strTemp As String
Dim AFID(4) As String
'edit by nickc 2007/02/06 不用 dll 了
'Dim obj0701 As Object
Dim Fld1 As String
Dim Fld2 As String
Dim Fld3 As String
Dim Fld4 As String
Dim DelFlg As Boolean
Dim RsCounts As Integer
Dim nRet As Boolean
Dim InitValue As Boolean
Dim GetNowData As Boolean
Dim ChkData As Boolean
Dim BlnULetter As Boolean
Dim blnKeypreview As Boolean
'Add By Sindy 2014/4/23 執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'2014/4/23 END
Dim m_ID As String 'Added by Lydia 2015/04/10


Private Sub Command1_Click()
'Modified by Lydia 2018/05/10
'Dim TempRs As ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim strTemp As String
   strTemp = Trim$(Text2.Text)
   If strTemp <> "" Then strTemp = strTemp + String(8 - Len(strTemp), "0")
   MFG1.ColWidth(0) = 1200
   MFG1.ColWidth(1) = 2000
   strSql = "select afid02||' - ' ||nvl(na03,nvl(na04,nvl(na05,''))) 申請國家,afid03 ID號碼  from applicantforeignid,nation where afid01='" + Trim$(strTemp) + "' and afid02=na01(+)"
   'Modified by Lydia 2018/05/10 debug-O12
'   Adodc1.ConnectionString = cnnConnection.ConnectionString
'   Adodc1.RecordSource = strSql
'   Adodc1.Refresh
'   If Adodc1.Recordset.RecordCount = 0 Then
   intI = 1
   Set rsAD = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
   'end 2018/05/10
      MsgBox "無此客戶編號之資料，請重新輸入 !", vbInformation
      Text2.SetFocus
      Exit Sub
   'Added by Lydia 2018/05/10
   Else
       Set Adodc1.Recordset = rsAD
   'end 2018/05/10
   End If
   With MFG1
      .row = 1
      .ColWidth(0) = 1800
      .ColWidth(1) = 1500
   End With
   MFG1.Refresh
   If strTemp <> "" Then Text2.Text = IIf(Right(strTemp, 2) = "00", Mid(strTemp, 1, 6), strTemp)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF9, vbKeyF10, vbKeyPageUp, vbKeyPageDown
'         UseDatamaintain (KeyCode)
'         KeyCode = 0
'      Case vbKeyEnd, vbKeyHome
'         If Data_Mission <> 1 And Data_Mission <> 3 Then
'            UseDatamaintain (KeyCode)
'            KeyCode = 0
'         End If
'      Case vbKeyEscape
'          If MsgBox("是否確定結束?", vbYesNo + vbCritical) = vbYes Then UseDatamaintain (KeyCode)
'      Case vbKeyReturn
'         UseDatamaintain (vbKeyF9)
'         KeyCode = 0
'   End Select
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If Data_Mission = 0 Then
               UseDatamaintain KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If Data_Mission = 0 Then
               UseDatamaintain KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If Data_Mission = 0 Then
               UseDatamaintain KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If Data_Mission = 0 Then
               UseDatamaintain KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If Data_Mission = 0 Then
               UseDatamaintain KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If Data_Mission <> 0 Then
            UseDatamaintain KeyCode
            KeyCode = 0
         End If
      Case vbKeyReturn:
         If Data_Mission <> 0 Then
            UseDatamaintain vbKeyF9
         End If
      Case vbKeyEscape:
         If Data_Mission = 0 Then
            'UseDatamaintain KeyCode
            If MsgBox("是否確定結束?", vbYesNo + vbCritical) = vbYes Then UseDatamaintain (KeyCode)
         Else
            UseDatamaintain vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If BlnULetter Then
        KeyAscii = UpperCase(KeyAscii)
    End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Add By Sindy 2014/4/23 取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   '2014/4/23 END
   
   OpenTable
   InitValue = True
   ChkData = True
   ShowData
   ChkData = True
   InitValue = False
   DelFlg = False
   OnOffTxt False
   blnKeypreview = True
   BlnULetter = False
   With MFG1
'      .Rows = 1
      .ColWidth(0) = 1800
      .ColWidth(1) = 1500
   End With
   Text1(1) = "011"
   SSTab1.Tab = 0 'Added by Lydia 2018/05/10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   rs.Close
   'Add By Cheng 2002/07/18
   Set frm050711 = Nothing

End Sub

Private Sub MFG1_Click()
   Fld3 = Text2.Text & String(8 - Len(Text2.Text), "0")
   Fld2 = Left(MFG1.TextMatrix(MFG1.row, 0), InStr(MFG1.TextMatrix(MFG1.row, 0), "-") - 2)
   'Added by Lydia 2015/04/10 +afid03
   m_ID = LTrim(RTrim(MFG1.TextMatrix(MFG1.row, 1)))
   rs.MoveFirst
   If Len(m_ID) > 0 Then
     QueryData "afid01=" + CNULL(Fld3), 3
   Else
     QueryData "afid01=" + CNULL(Fld3), 2
   End If

End Sub

Private Sub MFG1_DblClick()
 'Added by Lydia 2015/04/10 切換頁面
    SSTab1.Tab = 0
    Text1(0).SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case SSTab1.Tab
      Case 1
         Text2_GotFocus
         Text2.SetFocus
   End Select
End Sub
Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Select Case Index
   Case 0
        BlnULetter = True
   End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = UpperCase(KeyAscii)
      Case 1
         If KeyAscii = 13 Then
            If Data_Mission = 4 Then UseDatamaintain vbKeyF9
         End If
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
 Dim strTmp As String
   Select Case Index
   Case 0
        BlnULetter = False
   'Modified by Lydia 2015/04/10 +AFID03為KEY(Text1(2))
   Case 1, 2
      'Modified by Lydia 2015/04/10 改為模組
'      If Data_Mission = 1 Then
'         If Text1(0) <> "" And Text1(1) <> "" Then
'            strTmp = Text1(0)
'            strTmp = strTmp + String(8 - Len(strTmp), "0")
'            strExc(0) = "SELECT COUNT(*) FROM APPLICANTFOREIGNID WHERE AFID01='" & strTmp & "' AND AFID02='" & Text1(1) & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            If RsTemp.Fields(0) > 0 Then
'               MsgBox "此筆資料巳存在，請重新輸入 !", vbCritical
'               Text1(0).SetFocus
'            End If
'         End If
      If Data_Mission <> 0 Then
        If CheckExistsAFID(Text1(0), "" & Text1(1), "" & Text1(2)) = False Then
           Text1(Index).SetFocus
        End If
      End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String
Dim strA As String
   Select Case Index
   Case 0
    If Text1(0) = "" Then Me.Lbld1(0).Caption = "": Exit Sub
        If Left(Text1(0), 1) <> 客戶編號 Then ShowMsg MsgText(1101): Cancel = True: Exit Sub
        strA = Text1(0)
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCustomer(Text1(0).Text, strName) Then
        If ClsPDGetCustomer(Text1(0).Text, strName) Then
              Me.Lbld1(0).Caption = strName
              Text1(Index) = IIf(Right(Text1(Index), 2) = "00", Left(Text1(Index), 6), Text1(Index))
        Else
              Me.Lbld1(0).Caption = ""
              Cancel = True
              Text1_GotFocus (Index)
        End If
  Case 1
    If Text1(1) = "" Then Lbld1(1).Caption = "": Exit Sub
    'edit by nickc 2007/02/02 不用 dll 了
    'If objPublicData.GetNation(Text1(1), strName) Then
    If ClsPDGetNation(Text1(1), strName) Then
       Lbld1(1).Caption = strName
    Else
       Lbld1(1).Caption = ""
       Cancel = True
       Text1_GotFocus (Index)
       Exit Sub
    End If
    'Modified by Lydia 2015/04/10 +AFID03為KEY(Text1(2))
   Case 2
      If Data_Mission <> 0 Then
        If CheckExistsAFID(Text1(0), "" & Text1(1), "" & Text1(2)) = False Then
            Cancel = True
            Text1_GotFocus (Index)
            Exit Sub
        End If
      End If
   End Select
End Sub
Private Sub Text2_GotFocus()
    TextInverse Text2
    BlnULetter = True
End Sub

Private Sub Text2_LostFocus()
    BlnULetter = False
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = "" Then Exit Sub
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.GetCusCAJnam(Text2.Text, strExc(0), strExc(1), strExc(2)) Then
   If Not ClsLawGetCusCAJnam(Text2.Text, strExc(0), strExc(1), strExc(2)) Then
      Cancel = True
   Else
      AddCboName Combo1, strExc(0), strExc(1), strExc(2)
   End If
End Sub

Private Sub ShowData()
'edit by nickc 2007/02/06 不用 dll 了 Dim obj01 As Object
   rs.ReQuery
   RsCounts = rs.RecordCount
   If RsCounts > 0 Then
     rs.MoveFirst
   End If
   If rs.EOF Then
     Call Clear_AllTxtAry(Text1, 0, 2)
     Lbld1(0).Caption = ""
     Lbld1(1).Caption = ""
      tlbar.Buttons.Item(1).Enabled = True
      tlbar.Buttons.Item(2).Enabled = False
      tlbar.Buttons.Item(3).Enabled = False
      tlbar.Buttons.Item(4).Enabled = False
      tlbar.Buttons.Item(6).Enabled = False
      tlbar.Buttons.Item(7).Enabled = False
      tlbar.Buttons.Item(8).Enabled = False
      tlbar.Buttons.Item(9).Enabled = False
      tlbar.Buttons.Item(11).Enabled = False
      tlbar.Buttons.Item(12).Enabled = False
      tlbar.Buttons.Item(14).Enabled = True
      Exit Sub
   End If
'      tlbar.Buttons.Item(1).Enabled = True
'      tlbar.Buttons.Item(2).Enabled = True
'      tlbar.Buttons.Item(3).Enabled = True
'      tlbar.Buttons.Item(4).Enabled = True
'      tlbar.Buttons.Item(6).Enabled = True
'      tlbar.Buttons.Item(7).Enabled = True
'      tlbar.Buttons.Item(8).Enabled = True
'      tlbar.Buttons.Item(9).Enabled = True
'      tlbar.Buttons.Item(11).Enabled = False
'      tlbar.Buttons.Item(12).Enabled = False
'      tlbar.Buttons.Item(14).Enabled = True
      If m_bInsert Then
         tlbar.Buttons(1).Enabled = True
      Else
         tlbar.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
         tlbar.Buttons(2).Enabled = True
      Else
         tlbar.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
         tlbar.Buttons(3).Enabled = True
      Else
         tlbar.Buttons(3).Enabled = False
      End If
      If m_bQuery Then
         tlbar.Buttons(4).Enabled = True
      Else
         tlbar.Buttons(4).Enabled = False
      End If
      If m_bQuery Then
         tlbar.Buttons(6).Enabled = True
         tlbar.Buttons(7).Enabled = True
         tlbar.Buttons(8).Enabled = True
         tlbar.Buttons(9).Enabled = True
      Else
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
      End If
      tlbar.Buttons(11).Enabled = False
      tlbar.Buttons(12).Enabled = False
      tlbar.Buttons(14).Enabled = True
      
      If RsCounts > 1 And Not InitValue Then
           If DelFlg Then
               QueryData "afid01=" + CNULL(Fld3), 2
           Else
               QueryData "afid01=" + CNULL(Fld1), 2
           End If
           Exit Sub
      Else
           ShowDetail
           Exit Sub
      End If
End Sub

Private Sub OnOff_Button(TlBarCtrl As Control, ButtonValue As Boolean)
'Dim i As Integer
'   For i = 1 To 4
'      TlBarCtrl.Buttons.Item(i).Enabled = ButtonValue
'   Next
'   For i = 5 To 9
'      TlBarCtrl.Buttons.Item(i).Enabled = ButtonValue
'   Next
   If m_bInsert Then
      tlbar.Buttons(1).Enabled = ButtonValue
   Else
      tlbar.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
      tlbar.Buttons(2).Enabled = ButtonValue
   Else
      tlbar.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
      tlbar.Buttons(3).Enabled = ButtonValue
   Else
      tlbar.Buttons(3).Enabled = False
   End If
   If m_bQuery Then
      tlbar.Buttons(4).Enabled = ButtonValue
   Else
      tlbar.Buttons(4).Enabled = False
   End If
   If m_bQuery Then
      tlbar.Buttons(6).Enabled = ButtonValue
      tlbar.Buttons(7).Enabled = ButtonValue
      tlbar.Buttons(8).Enabled = ButtonValue
      tlbar.Buttons(9).Enabled = ButtonValue
   Else
      tlbar.Buttons(6).Enabled = False
      tlbar.Buttons(7).Enabled = False
      tlbar.Buttons(8).Enabled = False
      tlbar.Buttons(9).Enabled = False
   End If
   If ButtonValue = True Then
      TlBarCtrl.Buttons.Item(11).Enabled = False
      TlBarCtrl.Buttons.Item(12).Enabled = False
   Else
      TlBarCtrl.Buttons.Item(11).Enabled = True
      TlBarCtrl.Buttons.Item(12).Enabled = True
   End If
   TlBarCtrl.Buttons.Item(14).Enabled = ButtonValue
End Sub

Private Function ChkInData() As Boolean
Dim i As Integer
   If Text1(0) <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If Not objPublicData.GetCustomer(Text1(0).Text, strName) Then
        If Not ClsPDGetCustomer(Text1(0).Text, strName) Then
           SSTab1.Tab = 0
           Text1(0).SetFocus
           ChkInData = False
           Exit Function
        End If
   Else
        ShowMsg "客戶編號不可空白 !"
        SSTab1.Tab = 0
        Text1(0).SetFocus
        ChkInData = False
        Exit Function
   End If
  '******************************************************
   If Text1(1) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If Not objPublicData.GetNation(Text1(1), strName) Then
         If Not ClsPDGetNation(Text1(1), strName) Then
            Text1(1).SetFocus
            ChkInData = False
            Exit Function
        End If
   Else
        ShowMsg "申請國家不可空白 !"
        Text1(1).SetFocus
        ChkInData = False
        Exit Function
   End If
   '****************************************************
   If Text1(2) = "" Then
        ShowMsg "ID 號碼不可空白 !"
        Text1(2).SetFocus
        ChkInData = False
        Exit Function
   End If
    ChkInData = True
End Function
'Modified by Lydia 2015/04/10 改寫在form
Private Sub Ins_Data()
'Modified by Lydia 2015/04/10 同一國可以有多筆識別番號,改成先刪除再新增
' Dim i As Integer
'   AFID(0) = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
'   For i = 1 To 2
'      AFID(i) = Trim$(Text1(i).Text)
'   Next
'   'edit by nickc 2007/02/06 不用 dll 了
'   'Set obj0701 = CreateObject("PrjTaieDll.Class0703")
'   'nRet = obj0701.AddData0711(AFID)
'   'Set obj0701 = Nothing
'   nRet = Cls0703AddData0711(AFID)
Dim strM As String
    Call ReadCase(1)
    strM = "INSERT INTO APPLICANTFOREIGNID(AFID01,AFID02,AFID03) VALUES('" & AFID(0) & "','" & AFID(1) & "','" & AFID(2) & "')"

    cnnConnection.BeginTrans
      cnnConnection.Execute strM
    cnnConnection.CommitTrans
End Sub

Private Sub DeleteData()
'   AFID(0) = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
'   AFID(1) = Trim$(Text1(1).Text)
'   'edit by nickc 2007/02/06 不用 dll 了
'   'Set obj0701 = CreateObject("PrjTaieDll.Class0703")
'   'nRet = obj0701.EraseData0711(AFID)
'   'Set obj0701 = Nothing
'   nRet = Cls0703EraseData0711(AFID)
Dim strM As String
    Call ReadCase(1)
    strM = "DELETE FROM APPLICANTFOREIGNID WHERE AFID01='" & AFID(0) & "' AND AFID02='" & AFID(1) & "' AND AFID03='" & AFID(2) & "' "

    cnnConnection.BeginTrans
      cnnConnection.Execute strM
    cnnConnection.CommitTrans
End Sub

Private Sub UpdateData()
'Modified by Lydia 2015/04/10 同一國可以有多筆識別番號,改成先刪除再新增
' Dim i As Integer
'   AFID(0) = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
'   For i = 1 To 2
'      AFID(i) = Trim$(Text1(i).Text)
'   Next
'   'edit by nickc 2007/02/06 不用 dll 了
'   'Set obj0701 = CreateObject("PrjTaieDll.Class0703")
'   'nRet = obj0701.ModifyData0711(AFID)
'   'Set obj0701 = Nothing
'   nRet = Cls0703ModifyData0711(AFID)
Dim strM1 As String, StrM2 As String
    '先刪除修改前的記錄
    Call ReadCase(2) '讀tag
    strM1 = "DELETE FROM APPLICANTFOREIGNID WHERE AFID01='" & AFID(0) & "' AND AFID02='" & AFID(1) & "' AND AFID03='" & AFID(2) & "' "
    '新增修改前的記錄
    Call ReadCase(1)
    StrM2 = "INSERT INTO APPLICANTFOREIGNID(AFID01,AFID02,AFID03) VALUES('" & AFID(0) & "','" & AFID(1) & "','" & AFID(2) & "')"

    cnnConnection.BeginTrans
      cnnConnection.Execute strM1
      cnnConnection.Execute StrM2
    cnnConnection.CommitTrans

End Sub
'end 2015/04/10
Public Sub ShowDetail()
 Dim i As Integer, strTemp As String
   strTemp = IIf(IsNull(rs(0)), "", rs(0))
   If Mid(strTemp, 7, 2) = "00" Then
      Text1(0).Text = Mid$(strTemp, 1, 6)
   Else
      Text1(0).Text = strTemp
   End If
   For i = 1 To 2
      Text1(i).Text = IIf(IsNull(rs(i)), "", rs(i))
   Next
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetNation(Text1(1), strName) Then
   If ClsPDGetNation(Text1(1), strName) Then
      Lbld1(1).Caption = strName
   Else
      Lbld1(1).Caption = ""
   End If
   
   If Text1(0) <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCustomer(Text1(0).Text, strName) Then
      If ClsPDGetCustomer(Text1(0).Text, strName) Then
         Text1(0) = IIf(Right(Text1(0), 2) = "00", Left(Text1(0), 6), Text1(0))
         Me.Lbld1(0).Caption = strName
      Else
         Me.Lbld1(0).Caption = ""
      End If
   End If
End Sub

Public Sub OpenTable()

   strSql = "select * from applicantforeignid order by afid01,afid02"
   'edit by nickc 2007/02/02 不用 dll 了
   'Set rs = objPublicData.ReadRst(strSQL, True)
   Set rs = ClsPDReadRst(strSql, True)
End Sub
Private Sub OnOffTxt(OnOffValue As Boolean)
Dim i As Integer
    For i = 0 To 2
        Text1(i).Locked = Not OnOffValue
    Next
End Sub

Private Sub QueryData(StrCriteria As String, i As Integer)
 On Error Resume Next
   
   rs.Find StrCriteria
   i = i - 1
   If rs.EOF Then
      ShowMsg MsgText(9007)
      DelFlg = False
      ShowData
   Else
      If i <> 0 Then
        'Added by Lydia 2015/04/10 +afid03
        If i = 2 Then
           QueryData "afid03=" + CNULL(m_ID), i
        Else
            If DelFlg Then
               QueryData "afid02=" + CNULL(Fld4), i
            Else
               QueryData "afid02=" + CNULL(Fld2), i
            End If
        End If
      Else
        ShowDetail
        OnOff_Button tlbar, True
      End If
   End If
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        UseDatamaintain (vbKeyF2)
    Case 2
        UseDatamaintain (vbKeyF3)
    Case 3
        UseDatamaintain vbKeyF5
        UseDatamaintain vbKeyF9
    Case 4
        UseDatamaintain (vbKeyF4)
    Case 6
        UseDatamaintain (vbKeyHome)
    Case 7
        UseDatamaintain (vbKeyPageUp)
    Case 8
        UseDatamaintain (vbKeyPageDown)
    Case 9
        UseDatamaintain (vbKeyEnd)
    Case 11
        UseDatamaintain (vbKeyF9)
    Case 12
        UseDatamaintain (vbKeyF10)
    Case 14
        UseDatamaintain (vbKeyEscape)
    End Select
End Sub

Private Sub UseDatamaintain(j As Integer)
 Dim i As Integer, Rss As ADODB.Recordset
On Error GoTo HndErr
   If Data_Mission <> 2 And Not GetNowData Then
      Fld1 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
      Fld2 = Trim$(Text1(1).Text)
   End If
   If Data_Mission = 2 And GetNowData Then
      Fld1 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
      Fld2 = Trim$(Text1(1).Text)
      If RsCounts > 1 Then
          rs.MoveNext
          If rs.EOF Then rs.MoveFirst
          Fld3 = rs(0)
          Fld4 = rs(1)
      End If
   End If
   
    Command1.Enabled = True 'Modified by Lydia 2015/04/10
    
   Select Case j
      Case vbKeyF2 '新增
         If blnKeypreview Then
         Data_Mission = 1
         GetNowData = True
         ChkData = False
         Call Clear_AllTxtAry(Text1, 0, 2)
         Call StoreText 'Modified by Lydia 2015/04/10 tag初始化
         Command1.Enabled = False '防止查詢
         
         ChkData = True
         Text2.Text = ""
         Lbld1(0).Caption = ""
         Lbld1(1).Caption = ""
         Call OnOff_Button(tlbar, False)
         OnOffTxt True
         Text1(0).SetFocus
         blnKeypreview = False
         End If
      Case vbKeyF3 '修改
         If blnKeypreview Then
         Data_Mission = 3
         GetNowData = True
         Call StoreText 'Modified by Lydia 2015/04/10 保留修改前的記錄
         Command1.Enabled = False '防止查詢
         
         Call OnOff_Button(tlbar, False)
         OnOffTxt True
         Text1(0).Locked = True
         Text1(1).Locked = True
         Text1(2).SetFocus
         blnKeypreview = False
         End If
      Case vbKeyF5 '刪除
         If blnKeypreview Then
         Data_Mission = 2
         GetNowData = True
         Call OnOff_Button(tlbar, False)
         OnOffTxt False
         Text1(0).SetFocus
         blnKeypreview = False
         End If
      Case vbKeyF4 '查詢
         If blnKeypreview Then
         Data_Mission = 4
         Call OnOff_Button(tlbar, False)
         GetNowData = True
         ChkData = False
         Call Clear_AllTxtAry(Text1, 0, 2)
         ChkData = True
         Lbld1(0).Caption = ""
         Lbld1(1).Caption = ""
         OnOffTxt False
         Text1(0).Locked = False
         Text1(1).Locked = False
         Text1(0).SetFocus
         blnKeypreview = False
         End If
      Case vbKeyHome '第一筆
         If blnKeypreview Then
         rs.MoveFirst
         ShowDetail
         Text1(0).SetFocus
         End If
      Case vbKeyPageUp '上一筆
         If blnKeypreview Then
         rs.MovePrevious
         If rs.BOF Then
            rs.MoveFirst
           ShowMsg MsgText(9008)
         End If
         ShowDetail
         Text1(0).SetFocus
         End If
      Case vbKeyPageDown '下一筆
         If blnKeypreview Then
         rs.MoveNext
         If rs.EOF Then
            rs.MoveLast
            ShowMsg MsgText(9009)
         End If
         ShowDetail
         Text1(0).SetFocus
         End If
      Case vbKeyEnd '最後一筆
         If blnKeypreview Then
         rs.MoveLast
         ShowDetail
         Text1(0).SetFocus
         End If
      Case vbKeyF9 '確定
         If Not blnKeypreview Then
         Select Case Data_Mission
            Case 1 '新增->確定
               If ChkInData Then
                  'Add By Cheng 2002/05/22
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Sub
                   'Modified by Lydia 2015/04/10 改為模組
'                  strSql = "select * from applicantforeignid where afid01=" + CNULL(Text1(0)) + " and afid02=" + CNULL(Text1(1))
'                  'edit by nickc 2007/02/02 不用 dll 了
'                  'Set Rss = objPublicData.ReadRst(strSQL, True)
'                  Set Rss = ClsPDReadRst(strSql, True)
'                  If Not Rss.EOF Then
'                     MsgBox "編號:" + Trim$(Text1(0)) + "-" + Trim$(Text1(1)) + "的資料已存在.", vbCritical
'                     Rss.Close
'                     OnOffTxt True
'                     Text1(0).SetFocus
'                     Exit Sub
'                  End If
'                  Rss.Close
                  Ins_Data
                  DelFlg = False
               Else
                  Exit Sub
               End If
               Fld1 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
               Fld2 = Trim$(Text1(1).Text)
               ShowData
            Case 2 '刪除->確定
               If DelMsg Then
                  DeleteData
                  DelFlg = True
               End If
               ShowData
               DelFlg = False
            Case 3 '修改->確定
               If Not ChkInData Then Exit Sub
               'Add By Cheng 2002/05/22
               '重新檢查欄位有效性
               If TxtValidate = False Then Exit Sub
               
               UpdateData
               ShowData
               Text1(0).Locked = False
               Text1(1).Locked = False
               DelFlg = False
            Case 4  '查詢->確定
               If Text1(0) = "" Then
                  ShowMsg MsgText(9015)
                  SSTab1.Tab = 0
                  Text1(0).SetFocus
                  Exit Sub
               End If
  '******************************************************
               If Text1(1) = "" Then
                  ShowMsg MsgText(9015)
                  Text1(1).SetFocus
                  Exit Sub
               End If
               Fld3 = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
               Fld2 = Trim$(Text1(1).Text)
               rs.MoveFirst
               QueryData "afid01=" + CNULL(Fld3), 2
            End Select
         GetNowData = False
         OnOffTxt False
         Data_Mission = 0
         blnKeypreview = True
         Text1(0).SetFocus
         End If
      Case vbKeyF10 '取消
         If Not blnKeypreview Then
         If Data_Mission <> 4 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Sub
         End If
         OnOffTxt False
         'Call Clear_AllTxtAry(Text1, 0, 2)
         DelFlg = False
         Data_Mission = 0
         ShowData
         GetNowData = False
         blnKeypreview = True
         Text1(0).SetFocus
         End If
      Case vbKeyEscape
         Unload Me
   End Select
   Exit Sub
HndErr:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text1
   If objTxt.Enabled = True Then
      Cancel = False
      Text1_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

If Me.Text2.Enabled = True Then
   Cancel = False
   Text2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
'Added by Lydia 2015/04/10 模組判斷資料是否存在
Private Function CheckExistsAFID(ByVal F01 As String, Optional ByVal F02 As String, Optional ByVal F03 As String) As Boolean
Dim strTmp As String
    
    CheckExistsAFID = False
    If F01 = "" Or F02 = "" Or F03 = "" Then
       CheckExistsAFID = True
    Else
       strTmp = Text1(0)
       strTmp = strTmp + String(8 - Len(strTmp), "0")
       strExc(0) = "SELECT COUNT(*) FROM APPLICANTFOREIGNID WHERE AFID01='" & strTmp & "' AND AFID02='" & F02 & "' AND AFID03='" & F03 & "'"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If RsTemp.Fields(0) > 0 Then
          MsgBox "編號:" + Trim$(F01) + "-" + Trim$(F02) + "-" + Trim$(F03) + "的資料已存在.", vbCritical
       Else
          CheckExistsAFID = True
       End If
    End If
End Function
Private Sub ReadCase(ByVal aKind As Integer)
 Dim i As Integer
 
   If aKind = 1 Then '讀現在畫面資料
        AFID(0) = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
        For i = 1 To 2
           AFID(i) = Trim$(Text1(i).Text)
        Next
   
   Else '讀修改前資料
        AFID(0) = Text1(0).Tag & String(8 - Len(Text1(0).Tag), "0")
        For i = 1 To 2
           AFID(i) = Trim$(Text1(i).Tag)
        Next
   End If
End Sub
'保留修改前的記錄
Private Sub StoreText()
 Dim i As Integer
    For i = 0 To 2
       Text1(i).Tag = Trim$(Text1(i).Text)
    Next

End Sub
