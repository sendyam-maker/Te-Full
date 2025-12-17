VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160011 
   BorderStyle     =   1  '單線固定
   Caption         =   "職稱代號資料"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8160
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
            Picture         =   "frm160011.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160011.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
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
      Height          =   4380
      Left            =   30
      TabIndex        =   3
      Top             =   660
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160011.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textAC03"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textAC02"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160011.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox textAC02 
         Height          =   285
         Left            =   1050
         TabIndex        =   0
         Top             =   390
         Width           =   585
      End
      Begin VB.TextBox textAC03 
         Height          =   270
         Left            =   1050
         MaxLength       =   20
         TabIndex        =   1
         Top             =   690
         Width           =   4035
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Bindings        =   "frm160011.frx":212C
         Height          =   4005
         Left            =   -74970
         TabIndex        =   4
         Top             =   330
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   7064
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
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   3990
         Width           =   7785
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13732;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "歸類：01~09 研究所　10~19 大學　20~29 專科　30~39 高中　40~ 國中以下"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1380
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "代號"
         Height          =   180
         Left            =   630
         TabIndex        =   6
         Top             =   450
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "中文說明"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   720
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm160011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by nickc 2006/11/01 copy from frm140401
Option Explicit
Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
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
' 第一筆資料的本所案號
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_ac As Integer
Dim MyKind As String


Private Sub Form_Initialize()
Set rsA = New ADODB.Recordset
If rsA.State = 1 Then rsA.Close
rsA.CursorLocation = adUseClient
rsA.Open "select * from allcode where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
tf_ac = rsA.Fields.Count
Select Case ProSysState
Case "A"
        MyKind = "01"
        Me.Caption = "職稱代號資料維護"
Case "B"
        MyKind = "02"
        Me.Caption = "職位代號資料維護"
Case "C"
        MyKind = "03"
        Me.Caption = "學歷代號資料維護"
        'Add By Sindy 2015/12/9
        Label1(2).Visible = True
        '2015/12/9 END
Case "D"
        MyKind = "04"
        Me.Caption = "假別代號資料維護"
Case "E"
        MyKind = "05"
        Me.Caption = "異動原因代號資料維護"
Case "F"
        MyKind = "06"
        Me.Caption = "出生地代號資料維護"
Case "G"
        MyKind = "07"
        Me.Caption = "執行智權人員業別代號資料維護"
Case "H"
        MyKind = "08"
        Me.Caption = "獎懲代號資料維護"
Case Else
        MyKind = "99"
        Me.Caption = "非法使用，請關閉程式"
End Select

SetGrd
End Sub

' 按下按鍵
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
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   ReDim m_FieldList(tf_ac) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name & Chr(Val(MyKind) + 64), strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name & Chr(Val(MyKind) + 64), strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name & Chr(Val(MyKind) + 64), strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name & Chr(Val(MyKind) + 64), strFind, False)
   
   textAC02.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160011 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

'Sub getGrdColRow(ByRef oObj As MSHFlexGrid, ByVal x As Single, ByVal y As Single, ByRef col As Long, ByRef row As Long)
'Dim nIndex As Integer
'col = 0: row = 0
'For nIndex = 0 To oObj.Rows - 1
'    If y > oObj.RowHeight(nIndex) Then
'        row = row + 1
'        y = y - oObj.RowHeight(nIndex)
'    ElseIf y > 0 Then
'        row = row + 1
'        Exit For
'    End If
'Next nIndex
'For nIndex = 0 To oObj.Cols - 1
'    If x > oObj.ColWidth(nIndex) Then
'        col = col + 1
'        x = x - oObj.ColWidth(nIndex)
'    ElseIf x > 0 Then
'        col = col + 1
'        Exit For
'    End If
'Next nIndex
'col = col - 1 + oObj.LeftCol
'row = row - 1 + oObj.TopRow
'End Sub



Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j
GRD1.Visible = False
tmpMouseRow = GRD1.row
GRD1.Visible = True
If tmpMouseRow <> 0 Then
    GRD1.row = tmpMouseRow
    GRD1.col = 0
    If GRD1.CellBackColor = QBColor(15) Then
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
         textAC02.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textAC03.Text = GRD1.TextMatrix(tmpMouseRow, 1)
         m_CurrKEY(1) = textAC02   '2008/12/12 ADD BY SONIA
         GRD1.Visible = True
    End If
End If
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

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   If IsNull(rsSrcTmp.Fields("ac04")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ac04")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("ac04"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ac05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ac05")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("ac05"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ac06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ac06")) = False Then
         strTemp = rsSrcTmp.Fields("ac06")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ac07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ac07")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("ac07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ac08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ac08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("ac08"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("ac09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("ac09")) = False Then
         strTemp = rsSrcTmp.Fields("ac09")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Function TxtValidate() As Boolean

Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textAC02.Enabled = True Then
   Cancel = False
   textac02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textAC02.Text = "" Then
    MsgBox "代碼不可以空白！", vbExclamation
    textAC02.SetFocus
    Exit Function
End If
If Me.textAC03.Enabled = True Then
   Cancel = False
   textac03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
TxtValidate = True
End Function

'add by nickc 2006/10/24
' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To tf_ac - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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
   
   For nIndex = 0 To tf_ac - 1
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
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strAC01 As String
   Dim strAC02 As String
   
   AddRecord = False
   
   strAC01 = MyKind
   strAC02 = textAC02

   ' 檢查記錄是否已存在
   If IsRecordExist(strAC01, strAC02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO allcode ("
   For nIndex = 0 To tf_ac - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To tf_ac - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If ((strAC01 & strAC02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strAC01 & strAC02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strAC01, strAC02
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
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strAC01 As String
   Dim strAC02 As String
       
   ModRecord = False
   
   strAC01 = m_CurrKEY(0)
   strAC02 = m_CurrKEY(1)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE allcode SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_ac - 1
      strTmp = Empty
      'If nIndex < 3 Or nIndex > 8 Then
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
                  End If
               Else
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
                  End If
               End If
            End If
            If strTmp <> Empty Then
               bDifference = True
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
        'End If
   Next nIndex

   strSql = strSql & " " & _
                  "WHERE ac01 = '" & strAC01 & "' and ac02='" & strAC02 & "' ; end; "
On Error GoTo ErrHand
      cnnConnection.BeginTrans
        If bDifference = True Then
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
        End If
        cnnConnection.CommitTrans

      ShowCurrRecord strAC01, strAC02
      
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strAC01 As String
   Dim strAC02 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strAC01 = m_CurrKEY(0)
   strAC02 = m_CurrKEY(1)

   strSql = "DELETE FROM allcode " & _
            "WHERE ac01 = '" & strAC01 & "'  and ac02='" & strAC02 & "' "
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql

   If (strAC01 = m_LastKEY(0) And strAC02 = m_LastKEY(1)) Or (strAC01 = m_FirstKEY(0) And strAC02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strAC01, strAC02
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strAC01 As String
   Dim strAC02 As String
   
   QueryRecord = False
   strAC01 = MyKind
   strAC02 = textAC02
   If IsRecordExist(strAC01, strAC02) = True Then
      m_CurrKEY(0) = strAC01
      m_CurrKEY(1) = strAC02
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
            'Add By Cheng 2002/05/22
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
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textAC02 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
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
      Case 1: If Me.Visible = True Then textAC02.SetFocus
      Case 2: If Me.Visible = True Then textAC03.SetFocus
      Case 4: If Me.Visible = True Then textAC02.SetFocus
   End Select
End Sub
' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM allcode " & _
            "WHERE ac01 = '" & strKEY01 & "'  and ac02='" & strKEY02 & "'  "
                  
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT ac01,ac02 FROM allcode " & _
               "WHERE ac01 = '" & m_CurrKEY(0) & "' and ac02='" & m_CurrKEY(1) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("ac01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ac01")
         If IsNull(rsTmp.Fields("ac02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ac02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT ac01,ac02 FROM allcode " & _
               "WHERE ac02 = (SELECT MIN(ac02) FROM allcode where ac01='" & MyKind & "' ) and ac01='" & MyKind & "' "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("ac01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ac01")
         If IsNull(rsTmp.Fields("ac02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ac02")
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
   m_CurrKEY(1) = m_FirstKEY(1)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT ac01,ac02 FROM allcode " & _
            "WHERE ac02 = (SELECT MAX(ac02) FROM allcode " & _
                          "WHERE ac02 < '" & m_CurrKEY(1) & "' and ac01='" & m_CurrKEY(0) & "' ) and ac01='" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ac01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ac01")
      If IsNull(rsTmp.Fields("ac02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ac02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ac01,ac02 FROM allcode " & _
            "WHERE ac02 = (SELECT Min(ac02) FROM allcode where ac01='" & m_CurrKEY(0) & "') and ac01='" & m_CurrKEY(0) & "' "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ac01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ac01")
      If IsNull(rsTmp.Fields("ac02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ac02")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT ac01,ac02 FROM allcode " & _
            "WHERE ac02 = (SELECT MIN(ac02) FROM allcode " & _
                          "WHERE ac02  > '" & m_CurrKEY(1) & "' and ac01='" & m_CurrKEY(0) & "' ) and ac01='" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ac01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ac01")
      If IsNull(rsTmp.Fields("ac02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ac02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ac01,ac02 FROM allcode " & _
            "WHERE ac02 = (SELECT max(ac02) FROM allcode where ac01='" & m_CurrKEY(0) & "' ) and ac01='" & m_CurrKEY(0) & "'  "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ac01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ac01")
      If IsNull(rsTmp.Fields("ac02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ac02")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   m_SubMode = 0
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
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT ac01,ac02 FROM allcode " & _
            "WHERE ac02 = (SELECT MIN(ac02) FROM allcode where ac01='" & MyKind & "') and ac01='" & MyKind & "'  "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ac01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("ac01")
      If IsNull(rsTmp.Fields("ac02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("ac02")
   End If
   rsTmp.Close

   strSql = "SELECT ac01,ac02 FROM allcode " & _
            "WHERE ac02 = (SELECT MAX(ac02) FROM allcode where ac01='" & MyKind & "') and ac01='" & MyKind & "'  "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ac01")) = False Then: m_LastKEY(0) = rsTmp.Fields("ac01")
      If IsNull(rsTmp.Fields("ac02")) = False Then: m_LastKEY(1) = rsTmp.Fields("ac02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM allcode " & _
            "WHERE ac01='" & m_CurrKEY(0) & "' and ac02 = '" & m_CurrKEY(1) & "'   "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("ac02")) = False Then: textAC02 = rsTmp.Fields("ac02")
      If IsNull(rsTmp.Fields("ac03")) = False Then: textAC03 = rsTmp.Fields("ac03")

      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

        '抓取所有資料
        strSql = "SELECT ac02,ac03  FROM allcode " & _
                 "WHERE ac01 = '" & m_CurrKEY(0) & "'  order by ac02 "
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        Set GRD1.Recordset = rsTmp
   End If
   SetGrd
   If textAC02 <> "" Then
        GRD1.Visible = False
         For j = 1 To GRD1.Rows - 1
             GRD1.row = j
             If textAC02 = GRD1.TextMatrix(j, 0) Then
                For i = 0 To GRD1.Cols - 1
                    GRD1.col = i
                    GRD1.CellBackColor = &HFFC0C0
                Next i
            Else
                For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                Next i
            End If
        Next j
        GRD1.Visible = True
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
         SSTab1.TabEnabled(1) = True 'Add by Morgan 2011/11/9
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
         'Add by Morgan 2011/11/9
         SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         'end 2011/11/9
   End Select
   
End Sub
Private Function CheckDataValid() As Boolean
   Dim nResponse As Boolean
   Dim strTmp  As String
   CheckDataValid = False
   nResponse = False
   textac02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textac03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   CheckDataValid = True
EXITSUB:
End Function
' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textAC02.Locked = bEnable
   If bEnable Then textAC02.BackColor = &H8000000F Else textAC02.BackColor = &H80000005
End Sub
' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   textAC02.Locked = bEnable
   If bEnable Then textAC02.BackColor = &H8000000F Else textAC02.BackColor = &H80000005
End Sub
Private Sub ClearField()
Dim nIndex As Integer
   
   textAC02 = Empty
   textAC03 = Empty
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_ac - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
    
   SetFieldNewData "AC01", MyKind
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "AC02", textAC02
   End If
   SetFieldNewData "AC03", ChgSQL(textAC03)

End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_ac
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "AC" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
SetGrd
End Sub

Private Sub textac02_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textAC02
End If
End Sub

Private Sub textac02_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textac02_Validate(Cancel As Boolean)
If m_EditMode = 1 Then
    If IsRecordExist(MyKind, textAC02) = True And textAC02.Enabled = True And textAC02.Locked = False Then
        MsgBox "此代號已經存在，請確認！", vbInformation
        Cancel = True
    End If
End If
End Sub

Private Sub textac03_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textAC03
    OpenIme
End If
End Sub

Private Sub textac03_Validate(Cancel As Boolean)
If m_EditMode <> 0 Then
    If CheckLengthIsOK(textAC03, textAC03.MaxLength) = False Then
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("代碼", "中文說明")
   arrGridHeadWidth = Array(600, 4000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub


