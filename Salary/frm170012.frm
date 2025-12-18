VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170012 
   BorderStyle     =   1  '單線固定
   Caption         =   "互助會得標名單"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8190
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
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
            Picture         =   "frm170012.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170012.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   30
      TabIndex        =   6
      Top             =   650
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170012.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDsp(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblRemark(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textCUID"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCM(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCM(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCM(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCM(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCM(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCM(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170012.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt1(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txt1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txt1(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label16"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label12"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170012.frx":212C
         Height          =   3495
         Left            =   -74970
         TabIndex        =   7
         Top             =   840
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6165
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
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69930
         MaxLength       =   6
         TabIndex        =   19
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71040
         MaxLength       =   6
         TabIndex        =   18
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74280
         MaxLength       =   3
         TabIndex        =   16
         Top             =   405
         Width           =   735
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -73260
         MaxLength       =   3
         TabIndex        =   17
         Top             =   405
         Width           =   735
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   400
         Left            =   -68640
         TabIndex        =   20
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtCM 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1620
         MaxLength       =   7
         TabIndex        =   5
         Text            =   "9999999"
         Top             =   2110
         Width           =   915
      End
      Begin VB.TextBox txtCM 
         Height          =   285
         Index           =   2
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "01"
         Top             =   830
         Width           =   405
      End
      Begin VB.TextBox txtCM 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   5
         Left            =   1620
         MaxLength       =   7
         TabIndex        =   4
         Text            =   "9999999"
         Top             =   1790
         Width           =   915
      End
      Begin VB.TextBox txtCM 
         Height          =   270
         Index           =   1
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "001"
         Top             =   510
         Width           =   400
      End
      Begin VB.TextBox txtCM 
         Height          =   270
         Index           =   3
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "99999"
         Top             =   1150
         Width           =   915
      End
      Begin VB.TextBox txtCM 
         Height          =   270
         Index           =   4
         Left            =   1620
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "960501"
         Top             =   1470
         Width           =   915
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3960
         Width           =   5700
         VariousPropertyBits=   671105055
         Size            =   "7223;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "PS : 離職會員轉 60000 台一互助會"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   720
         TabIndex        =   24
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "員工代號：                      －"
         Height          =   180
         Left            =   -72000
         TabIndex        =   23
         Top             =   450
         Width           =   2070
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "會號：                   －"
         Height          =   180
         Left            =   -74880
         TabIndex        =   21
         Top             =   450
         Width           =   1575
      End
      Begin MSForms.Label lblDsp 
         Height          =   300
         Index           =   1
         Left            =   2685
         TabIndex        =   15
         Top             =   1200
         Width           =   810
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1429;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "會  款  金："
         Height          =   180
         Left            =   690
         TabIndex        =   13
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "編　　號：                       (電腦編號)"
         Height          =   180
         Left            =   690
         TabIndex        =   12
         Top             =   885
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "得  標  金："
         Height          =   180
         Left            =   690
         TabIndex        =   11
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "會　　號："
         Height          =   180
         Index           =   4
         Left            =   690
         TabIndex        =   10
         Top             =   555
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "會員員工代號："
         Height          =   180
         Index           =   3
         Left            =   330
         TabIndex        =   9
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "得  標  日："
         Height          =   180
         Index           =   2
         Left            =   690
         TabIndex        =   8
         Top             =   1515
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm170012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/22 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/23 add by sonia 隱藏會款金不顯示
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_CM As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean
Dim m_CO02 As Variant       '記錄會號之金額
Dim m_CO04 As Variant       '記錄會號之期間起
Dim m_CO05 As Variant       '記錄會號之期間迄
Dim m_CO07 As Variant       '記錄會號之低標


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
      If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
         Exit Sub
      End If
      If RunNick(txt1(2), txt1(3)) Then
         txt1(2).SetFocus
         Exit Sub
      End If
      GetData
   Else
      MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
      txt1(0).SetFocus
   End If
End Sub

Sub GetData()
Dim stCon As String
   
   stCon = ""
   If txt1(0) <> "" Then
      stCon = stCon & " and cm01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and cm01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
      stCon = stCon & " and cm03>='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
      stCon = stCon & " and cm03<='" & txt1(3) & "' "
   End If
   strExc(0) = "SELECT cm01 會號,cm02 編號,cm03 會員,ST02 姓名,sqldateT(cm04) 得標日,cm05 得標金 FROM CooperationMember,Cooperation,staff " & _
               " where cm01=co01(+) and cm03=st01(+) " & stCon & " order by cm01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
   End If
End Sub

Private Sub Form_Activate()
   If m_bActived = False Then
      SetInputEntry
      m_bActived = True
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170012 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from CooperationMember where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_CM = .Fields.Count
      ReDim m_FieldList(TF_CM) As FIELDITEM
      For Each oText In txtCM
         idx = oText.Index
         m_FieldList(idx).fiName = "CM" & Format(idx, "00")
         'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
         'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
            m_FieldList(idx).fiType = 0
         'Else
         '   m_FieldList(idx).fiType = 1
         'End If
         'end 2017/06/29
      Next
      End With
   End If
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stKey01 As String
Dim stKey02 As String
Dim adoRst As New ADODB.Recordset
   
   stKey01 = txtCM(1)
   stKey02 = txtCM(2)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM CooperationMember" & _
            " WHERE cm01 = '" & stKey01 & "' and cm02= '" & stKey02 & "'"
      Case -2
         strExc(0) = "SELECT * FROM CooperationMember order by 1 ASC"
      Case -1
         strExc(0) = "SELECT * FROM CooperationMember" & _
            " WHERE cm01||cm02 <'" & stKey01 & stKey02 & "' order by 1 DESC"
      Case 1
         strExc(0) = "SELECT * FROM CooperationMember" & _
            " WHERE cm01||cm02 >'" & stKey01 & stKey02 & "' order by 1 ASC"
      Case 2
         strExc(0) = "SELECT * FROM CooperationMember order by 1 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtCM(1).SetFocus
      txtCM_GotFocus 1
   End If
End Function

Private Sub GRD1_Click()
   Dim lCurRow As Long, i As Integer, j As Integer
   lCurRow = GRD1.row
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If GRD1.CellBackColor <> &HFFC0C0 Then
            GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
               GRD1.row = j
               If GRD1.CellBackColor <> QBColor(15) Then
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                  Next i
               End If
            Next j
            GRD1.row = lCurRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            GRD1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub GRD1_DblClick()
Dim lCurRow As Long
   
   lCurRow = GRD1.row
   '呼叫查詢
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
            If txtCM(1).Locked = False Then
               txtCM(1).Text = GRD1.TextMatrix(lCurRow, 0)
               txtCM(2).Text = GRD1.TextMatrix(lCurRow, 1)
               If TBar1.Buttons(11).Enabled = True Then
                  Call Tbar1_ButtonClick(TBar1.Buttons(11))
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 2 Then
      txt1(0).SetFocus
      TextInverse txt1(0)
   ElseIf SSTab1.Tab = 0 And PreviousTab = 2 Then
      GRD1_DblClick
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   CloseIme
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCM_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtCM(Index)
End Sub

Private Sub ClearField()
   For Each oText In txtCM
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_CM
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   m_bConfirmCheck = False
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtCM
         idx = oText.Index
         '日期轉民國
         If idx = 4 Then
            m_FieldList(idx).fiOldData = TransDate("" & .Fields(m_FieldList(idx).fiName), 1)
         Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End If
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
         oText.Tag = m_FieldList(idx).fiOldData
      Next
      
      If ClsPDGetStaff(txtCM(3), strExc(1)) Then
         lblDsp(1) = strExc(1)
      End If
      
      SetRefData (txtCM(1))
      
      CUID(1) = "" & .Fields("cm07")
      CUID(2) = "" & .Fields("cm08")
      CUID(3) = "" & .Fields("cm09")
      CUID(4) = "" & .Fields("cm10")
      CUID(5) = "" & .Fields("cm11")
      CUID(6) = "" & .Fields("cm12")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtCM(1).Tag = txtCM(1)
   txtCM(2).Tag = txtCM(2)
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtCM
      oText.Locked = bLocked
   Next
   txtCM(6).Locked = True      '會款金欄鎖住
   If m_EditMode = 4 Then
      txtCM(2).Locked = False  '查詢時編號欄打開
   Else
      txtCM(2).Locked = True   '非查詢時編號欄鎖住
   End If
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If

   End Select
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         SSTab1.Tab = 0
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF3 ' 修改
         SSTab1.Tab = 0
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         SSTab1.Tab = 0
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         SSTab1.Tab = 0
         m_EditMode = 4
         SetCtrlReadOnly True
         ClearField
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtCM(1) = txtCM(1).Tag
            txtCM(2) = txtCM(2).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And txtCM(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtCM(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtCM(1) <> "" Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         txtCM(1).Locked = False
         If Me.Visible = True Then
            txtCM(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtCM(1).Locked = True
         If Me.Visible = True Then
            txtCM(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 4
         txtCM(1).Locked = False
         If Me.Visible = True Then
            txtCM(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtCM(1).Locked = True
         If Me.Visible = True Then
            txtCM(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = True
   End Select
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtCM(1).SetFocus
               txtCM_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtCM
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtCM_Validate idx, bCancel
         If bCancel = True Then
            txtCM(idx).SetFocus
            txtCM_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtCM(1) = "" Then
         ShowMsg "請輸入會號 !"
         txtCM(1).SetFocus
         txtCM_GotFocus 1
         GoTo EscPoint
      End If
      If txtCM(2) = "" Then
         ShowMsg "請輸入編號 !"
         txtCM(2).SetFocus
         txtCM_GotFocus 2
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtCM(1) = "" And txtCM(1).Locked = False Then
         ShowMsg "請輸入會號 !"
         txtCM(1).SetFocus
         txtCM_GotFocus 1
         GoTo EscPoint
      End If
      If txtCM(3) = "" And txtCM(3).Locked = False Then
         ShowMsg "請輸入會員員工代號 !"
         txtCM(3).SetFocus
         txtCM_GotFocus 3
         GoTo EscPoint
      End If
      If txtCM(4) = "" And txtCM(4).Locked = False Then
         If txtCM(5) <> "" And txtCM(5).Locked = False Then
            ShowMsg "有得標金時請輸入得標日 !"
            txtCM(4).SetFocus
            txtCM_GotFocus 4
            GoTo EscPoint
         End If
      End If
      If txtCM(5) = "" And txtCM(5).Locked = False Then
         If txtCM(4) <> "" And txtCM(4).Locked = False Then
            ShowMsg "有得標日時請輸入得標金 !"
            txtCM(5).SetFocus
            txtCM_GotFocus 5
            GoTo EscPoint
         End If
      End If
   End If
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
Dim stCols As String, stValues As String, stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtCM
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            '日期轉西元
            If idx = 4 Then
               stValues = stValues & "," & CNULL(DBDATE(m_FieldList(idx).fiNewData), True)
            Else
               stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
            End If
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "declare intMax number;begin select max(CM02)+1 into intMax from CooperationMember where cm01='" & txtCM(1) & "';IF intMax IS NULL THEN intMax:=1; END IF;"
   stSQL = stSQL & "INSERT INTO CooperationMember (CM02," & stCols & ") Values (LTRIM(TO_CHAR(intMax,'00'))," & stValues & ");end;"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   stSQL = "select max(cm02) from CooperationMember where cm01='" & txtCM(1) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      txtCM(2) = RsTemp.Fields(0)
   End If
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE CooperationMember SET "
   stSet = ""
   For Each oText In txtCM
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where cm01='" & txtCM(1) & "' and cm02='" & txtCM(2) & "'; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

Private Sub UpdateFieldNewData()
   For Each oText In txtCM
      idx = oText.Index
      Select Case idx
         Case 4
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtCM_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 3     '2011/6/27 add by sonia 員工代號可輸英文字母
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtCM_Validate(Index As Integer, Cancel As Boolean)
Dim m_SDamount As Variant   '記錄個人薪資加項總額
Dim m_COamount As Variant   '記錄個人互助會總額
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 1
            If txtCM(Index) <> "" Then
               If SetRefData(txtCM(Index)) = False Then
                  Cancel = True
               End If
            End If
         Case 3
            If txtCM(Index) <> "" Then
               If ChkStaffID(txtCM(Index)) = True Then
                  Cancel = True
               End If
               If Cancel = False And ClsPDGetStaff(txtCM(Index), strExc(1)) = False Then
                  Cancel = True
               Else
                  lblDsp(1) = strExc(1)
               End If
               '新增時檢查該會號中之會員不可重覆，除台一以外
               If Cancel = False And m_EditMode = 1 And txtCM(3) <> "60000" Then
                  strExc(0) = "SELECT * FROM CooperationMember" & _
                              " WHERE cm01 = '" & txtCM(1) & "' and cm03= '" & txtCM(3) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     MsgBox "此會號此員工代號已輸入, 編號為第 " & RsTemp.Fields("CM02") & "號！"
                     Cancel = True
                  End If
               End If
               '檢查每人有效會款總額不可＞＝二家薪資加項總額的1/3
               If Cancel = False And txtCM(3) <> "60000" Then
                  m_SDamount = 0
                  'Modify By Sindy 2020/6/25 + 證照津貼 +NVL(SD52,0)
                  strExc(0) = "SELECT NVL(SD20,0)+NVL(SD21,0)+NVL(SD52,0)+NVL(SD22,0)+NVL(SD23,0)+NVL(SD24,0)+NVL(SD25,0)+NVL(SD26,0)+NVL(SD29,0)+NVL(SD30,0)+NVL(SD31,0)+NVL(SD32,0)+NVL(SD33,0)+NVL(SD34,0)+NVL(SD35,0) FROM SalaryData WHERE sd01 = '" & txtCM(3) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     m_SDamount = Val("" & RsTemp.Fields(0))
                  End If
                  strExc(0) = "SELECT sum(NVL(co02,0)) FROM CooperationMember,Cooperation " & _
                              "WHERE cm03= '" & txtCM(3) & "' and cm01=co01(+) and co04<=" & strSrvDate(1) & " and co05>=" & strSrvDate(1) & ""
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     m_COamount = Val("" & RsTemp.Fields(0))
                  End If
                  '新增或修改員工代號時要含畫面上此會號之金額
                  If txtCM(3).Tag <> txtCM(3) And txtCM(3) <> "" Then
                     m_COamount = m_COamount + m_CO02
                  End If
                  If m_COamount > m_SDamount / 3 Then
                     'Modified by Morgan 2015/11/27 改確認後可繼續 Ex.會號:315 會員:79017
                     'MsgBox "此員工之會款總額 " & m_COamount & " 已大於薪資加項總額 " & m_SDamount & " 的三分之一！"
                     If MsgBox("此員工之會款總額 " & m_COamount & " 已大於薪資加項總額 " & m_SDamount & " 的三分之一！是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                        Cancel = True
                     End If
                  End If
               End If
            End If
         Case 4
            If txtCM(Index) <> "" Then
               If ChkDate(txtCM(Index)) = False Then
                  Cancel = True
               End If
               If DBDATE(txtCM(Index)) < DBDATE(m_CO04) Or DBDATE(txtCM(Index)) > DBDATE(m_CO05) Then
                  MsgBox "得標日必須在互助會期間內 " & ChangeWStringToTDateString(DBDATE(m_CO04)) & "～" & ChangeWStringToTDateString(DBDATE(m_CO05)) & "！"
                  Cancel = True
               End If
            End If
         Case 5
            If txtCM(Index) <> "" Then
               If Val(txtCM(Index)) > Val(m_CO02) Then
                  MsgBox "得標金不可＞＝互助會的金額 " & m_CO02 & "！"
                  Cancel = True
               End If
               If Val(txtCM(Index)) < Val(m_CO07) Then
                  MsgBox "得標金不可＜互助會的低標 " & m_CO07 & "！"
                  Cancel = True
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtCM(Index)
      
      '若是按確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
         End Select
      End If
   End If
End Sub

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除
   stSQL = "delete from CooperationMember where cm01='" & txtCM(1) & "' and cm02='" & txtCM(2) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtCM(1).Tag = ""
   txtCM(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function SetRefData(stCode As String) As Boolean
   m_CO02 = 0: m_CO04 = 0: m_CO05 = 0: m_CO07 = 0
   strExc(0) = "select * from cooperation where co01='" & stCode & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         'lblDsp(1) = "" & .Fields("OC02")
         m_CO02 = "" & .Fields("CO02")
         m_CO04 = "" & .Fields("CO04")
         m_CO05 = "" & .Fields("CO05")
         m_CO07 = "" & .Fields("CO07")
      End With
      SetRefData = True
   Else
      MsgBox "無此互助會會號！"
   End If
End Function



