VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140503 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶預定收款日放寬月數上限"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8175
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
            Picture         =   "frm140503.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140503.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4350
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7673
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm140503.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCname"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblCU13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblCU13Nm"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textCL01"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textCL05"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm140503.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt1(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73950
         MaxLength       =   6
         TabIndex        =   5
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -70740
         MaxLength       =   8
         TabIndex        =   7
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -69600
         MaxLength       =   8
         TabIndex        =   9
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Default         =   -1  'True
         Height          =   315
         Left            =   -68310
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textCL05 
         Height          =   270
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1650
         Width           =   405
      End
      Begin VB.TextBox textCL01 
         Height          =   270
         Left            =   1980
         MaxLength       =   8
         TabIndex        =   2
         Top             =   900
         Width           =   1065
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm140503.frx":212C
         Height          =   3615
         Left            =   -74970
         TabIndex        =   11
         Top             =   690
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "客戶編號|客戶名稱|智權人員|預定收款日放寬月數"
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
         _Band(0).Cols   =   4
      End
      Begin MSForms.Label Label23 
         Height          =   300
         Left            =   60
         TabIndex        =   21
         Top             =   3990
         Visible         =   0   'False
         Width           =   6615
         VariousPropertyBits=   27
         Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LblCU13Nm 
         Height          =   300
         Left            =   2850
         TabIndex        =   20
         Top             =   1320
         Width           =   2265
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3995;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LblCU13 
         Height          =   300
         Left            =   1980
         TabIndex        =   19
         Top             =   1320
         Width           =   735
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCname 
         Height          =   300
         Left            =   3120
         TabIndex        =   18
         Top             =   900
         Width           =   4575
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label6 
         Height          =   300
         Left            =   -72990
         TabIndex        =   17
         Top             =   390
         Width           =   1035
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   16
         Top             =   390
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73260
         X2              =   -72660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71640
         TabIndex        =   15
         Top             =   390
         Width           =   540
      End
      Begin VB.Line Line3 
         X1              =   -70410
         X2              =   -69660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "預定收款日放寬月數："
         Height          =   180
         Index           =   2
         Left            =   750
         TabIndex        =   14
         Top             =   1710
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   0
         Left            =   750
         TabIndex        =   13
         Top             =   945
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   1
         Left            =   750
         TabIndex        =   12
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "智權人員："
         Height          =   225
         Left            =   -74880
         TabIndex        =   8
         Top             =   390
         Width           =   945
      End
      Begin VB.Label Label5 
         Caption         =   "客戶編號："
         Height          =   225
         Left            =   -71670
         TabIndex        =   6
         Top             =   390
         Width           =   945
      End
      Begin VB.Line Line2 
         X1              =   -69990
         X2              =   -69600
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label7 
         Caption         =   "PS：此設定不包含關係企業"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   540
         TabIndex        =   4
         Top             =   3540
         Width           =   5145
      End
   End
End
Attribute VB_Name = "frm140503"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 改成Form2.0 ; Label6、Label23、lblCname、LblCU13、LblCU13Nm
'Create By Lydia 2015/10/14 客戶預定收款日放寬月數上限
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料的Key值
Dim m_FirstKEY(1) As String
' 最後一筆資料的Key值
Dim m_LastKEY(1) As String
' 目前正在顯示的Key值
Dim m_CurrKEY(1) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_CL As Integer
Dim strText As String, arrKey As Variant
Private Const custSQL = "SELECT CU01 CL01,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CL02,CU13 CL03,ST02 CL04,CU143 CL05 FROM CUSTOMER,STAFF WHERE CU13=ST01(+) AND CU143>0 "

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open custSQL & " AND ROWNUM <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_CL = rsA.Fields.Count
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_Load()
   ReDim m_FieldList(tf_CL) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   InitialField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140503 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   If Me.textCL01.Enabled = True Then
      Cancel = False
      textCL01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_CL - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To tf_CL - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
   
   AddRecord = False
   
   ' 檢查記錄是否已存在
   If IsRecordExist(textCL01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '不顯示已存在記錄
      'UpdateCtrlData
      Exit Function
   End If
   
   For nIndex = 0 To tf_CL - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp = "CL05" Then
         strSql = "UPDATE CUSTOMER SET CU143=" & m_FieldList(nIndex).fiNewData & " WHERE CU01='" & textCL01 & "' "
      End If
   Next nIndex
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If (textCL01 < m_FirstKEY(0)) Or (textCL01 > m_LastKEY(0)) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord textCL01
   AddRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
       
   ModRecord = False
   
   For nIndex = 0 To tf_CL - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp = "CL05" Then
         strSql = "UPDATE CUSTOMER SET CU143=" & m_FieldList(nIndex).fiNewData & " WHERE CU01='" & textCL01 & "' AND CU143=" & m_FieldList(nIndex).fiOldData
      End If
   Next nIndex

On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans

   ShowCurrRecord m_CurrKEY(0)
      
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strSql = "UPDATE CUSTOMER SET CU143=NULL WHERE CU01='" & textCL01 & "' AND CU143=" & textCL05

   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If (m_CurrKEY(0) = m_LastKEY(0)) Or (m_CurrKEY(0) = m_FirstKEY(0)) Then
      RefreshRange
   End If
   ShowCurrRecord m_CurrKEY(0)
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryRecord = False
   
   If IsRecordExist(textCL01) = True Then
      m_CurrKEY(0) = textCL01
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse

   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If AddRecord = True Then
               RefreshRange
            Else
               Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textCL01 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            If textCL01 = "" Then
               MsgBox "請輸入客戶編號！", vbInformation
            End If
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textCL01.SetFocus
      Case 2: If Me.Visible = True Then textCL05.SetFocus
      Case 4: If Me.Visible = True Then textCL01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
    strSql = custSQL & " AND CU01='" & Left(Trim(strKEY01) & "00000000", 8) & "'"
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   strSql = custSQL & " AND ROWNUM < 2 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 0 Then
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
       strSql = custSQL & " AND CU01='" & Left(Trim(m_CurrKEY(0)) & "00000000", 8) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CL01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT MIN(CL01) FROM (" & custSQL & ")"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CurrKEY(0) = "" & rsTmp.Fields(0)
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If

   strSql = "SELECT MAX(CL01) FROM (" & custSQL & " AND CU01<'" & Left(Trim(m_CurrKEY(0)) & "00000000", 8) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_CurrKEY(0) = "" & rsTmp.Fields(0)
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_LastKEY(0) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT MIN(CL01) FROM (" & custSQL & " AND CU01>'" & Left(Trim(m_CurrKEY(0)) & "00000000", 8) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_CurrKEY(0) = "" & rsTmp.Fields(0)
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
               UpdateToolbarState
            Else
               Exit Sub
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   strSql = custSQL & " AND ROWNUM < 2"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 0 Then
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   strSql = "SELECT MIN(CL01) FROM (" & custSQL & ") "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_FirstKEY(0) = "" & rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   strSql = "SELECT MAX(CL01) FROM (" & custSQL & ") "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_LastKEY(0) = "" & rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer

   strSql = custSQL & " AND CU01='" & Left(Trim(m_CurrKEY(0)) & "00000000", 8) & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ClearField
   If rsTmp.RecordCount > 0 Then
      textCL01 = "" & rsTmp.Fields("CL01")
      lblCname = "" & rsTmp.Fields("CL02")
      LblCU13 = "" & rsTmp.Fields("CL03")
      LblCU13Nm = "" & rsTmp.Fields("CL04")
      textCL05 = "" & rsTmp.Fields("CL05")
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
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
         ' 新增
      Case 1, 2, 3, 4:
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

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTmp  As String
Dim strTit As String
Dim strMsg As String
   
   CheckDataValid = False
   
   If textCL01.Text = "" Then
       MsgBox "客戶編號不可空白！", vbExclamation
       textCL01.SetFocus
       Exit Function
   End If
   
   If Val(textCL05.Text) = 0 Then
       MsgBox "預定收款日放寬月數不可空白！", vbExclamation
       textCL05.SetFocus
       Exit Function
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCL01.Locked = bEnable
   If bEnable Then textCL01.BackColor = &H8000000F Else textCL01.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCL01.Locked = bEnable
   If bEnable Then textCL01.BackColor = &H8000000F Else textCL01.BackColor = &H80000005
   textCL05.Locked = bEnable
   '編輯資料時,關閉切換頁籤
   Me.SSTab1.TabEnabled(1) = bEnable
   If bEnable = False Then Me.SSTab1.Tab = 0
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textCL01 = Empty
   lblCname = Empty
   textCL05 = Empty
   Label23 = Empty
   LblCU13 = Empty
   LblCU13Nm = Empty
   
   For nIndex = 0 To tf_CL - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "CL01", textCL01
   End If
   SetFieldNewData "CL05", textCL05
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_CL
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CL" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex

   For nIndex = 0 To 2
      txt1(nIndex).Text = ""
   Next nIndex
   Label6.Caption = ""
   SetGrd
End Sub

Private Sub textCL01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textCL01
   End If
End Sub

Private Sub textCL01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCL01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   lblCname = Empty
   If IsEmptyText(textCL01) = False Then
      '2015/10/21 MODIFY BY SONIA 補齊8碼,否則存檔會讀不到
      'lblCname = GetPrjPeople1(Left(textCL01 & "00000000", 8), "1")
      textCL01 = Left(textCL01 & "00000000", 8)
      lblCname = GetPrjPeople1(textCL01, "1")
      '2015/10/21 END
      Select Case m_EditMode
         Case 1, 4:
            If Left(textCL01, 1) <> "X" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "必須輸入客戶編號"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCL01_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(lblCname) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "此客戶編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCL01_GotFocus
            Else
               QueryCustData
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textCL05_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textCL05
      CloseIme
   End If
End Sub

Private Sub textCL05_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
  If txt1(Index).Text <> "" Then
    Select Case Index
        Case 0
            If ClsPDGetStaff(txt1(0).Text, strExc(1), strExc(2)) Then
               Label6.Caption = strExc(1)
            Else
               txt1(0).SetFocus
               Cancel = True
            End If
        Case 1, 2
            If Index = 1 Then txt1(1).Text = Left(txt1(1).Text & "00000000", 8)
            If Index = 2 Then txt1(2).Text = Left(txt1(2).Text & "ZZZZZZZZ", 8)
            If Left(txt1(Index), 1) <> "X" Then
               MsgBox "客戶編號請輸入X編號!", vbCritical
               txt1(Index).SetFocus
               Cancel = True
            ElseIf txt1(1).Text <> "" And Index = 2 And txt1(1).Text > txt1(2).Text Then
               MsgBox "客戶編號起不可大於客戶編號止!", vbCritical
               txt1(Index).SetFocus
               Cancel = True
            End If
    End Select
  Else
    If Index = 0 Then Label6.Caption = ""
  End If
End Sub

Private Sub cmdOK_Click()
Dim Cancel As Boolean
Dim rsRead As New ADODB.Recordset

   For intI = 0 To 2
       txt1_Validate intI, Cancel
       If Cancel = True Then
          Exit Sub
       End If
   Next intI
   strExc(1) = "": strSql = ""
   If txt1(0).Text <> "" Then strExc(1) = strExc(1) & " AND CU13=" & CNULL(txt1(0).Text)
   If txt1(1).Text <> "" Then strExc(1) = strExc(1) & " AND CU01>=" & CNULL(txt1(1).Text)
   If txt1(2).Text <> "" Then strExc(1) = strExc(1) & " AND CU01<=" & CNULL(txt1(2).Text)
   
   strSql = custSQL & strExc(1) & " ORDER BY 1"
   intI = 0
   Set rsRead = ClsLawReadRstMsg(intI, strSql)
   Set GRD1.Recordset = rsRead
   GRD1.FixedCols = 0
   SetGrd (IIf(rsRead.RecordCount = 0, 2, rsRead.RecordCount + 1))

End Sub
Private Sub SetGrd(Optional ByVal iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   arrGridHeadText = Array("客戶編號", "客戶名稱", "CL03", "智權人員", "預定收款日放寬月數")
   arrGridHeadWidth = Array(1200, 2500, 0, 1000, 2500)
   GRD1.Visible = False
   GRD1.Rows = iR
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iCol = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iCol
      GRD1.Text = arrGridHeadText(iCol)
      GRD1.ColWidth(iCol) = arrGridHeadWidth(iCol)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub
Private Sub GRD1_DblClick()
   If InStr(GRD1.TextMatrix(GRD1.row, 0), textCL01.Text) > 0 Then
      Me.SSTab1.Tab = 0
   End If
End Sub
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub
Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j
GRD1.Visible = False
tmpMouseRow = GRD1.row
GRD1.Visible = True
If tmpMouseRow <> 0 Then
    GRD1.row = tmpMouseRow
    GRD1.col = 0
    If GRD1.CellBackColor <> &HFFC0C0 Then
                  GRD1.Visible = False
         For j = 1 To GRD1.Rows - 1
             GRD1.row = j
             For i = 0 To GRD1.Cols - 1
                  GRD1.col = i
                  GRD1.CellBackColor = QBColor(15)
             Next i
        Next j
        GRD1.row = tmpMouseRow
         For i = 0 To GRD1.Cols - 1
             GRD1.col = i
             GRD1.CellBackColor = &HFFC0C0
         Next i
         textCL01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         QueryRecord
         GRD1.Visible = True
    End If
End If
End Sub

Private Sub QueryCustData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   strSql = "SELECT cu13,st02 FROM Customer,staff WHERE CU01='" & Left(Trim(textCL01) & "00000000", 8) & "' and CU02='0' and cu13=st01(+)"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("cu13")) = False Then: LblCU13 = rsTmp.Fields("cu13")
      If IsNull(rsTmp.Fields("st02")) = False Then: LblCU13Nm = rsTmp.Fields("st02")
   Else
      LblCU13 = Empty
      LblCU13Nm = Empty
   End If

   rsTmp.Close

EXITSUB:
   Set rsTmp = Nothing
End Sub

