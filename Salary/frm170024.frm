VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170024 
   BorderStyle     =   1  '單線固定
   Caption         =   "基準月數輸入"
   ClientHeight    =   5070
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8200
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   8115
      _ExtentX        =   14323
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170024.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDsp(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "textCUID"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtYBM(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtYBM(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtYBM(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170024.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(1)=   "GRD1"
      Tab(1).Control(2)=   "cmdok"
      Tab(1).Control(3)=   "txt1(1)"
      Tab(1).Control(4)=   "txt1(0)"
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtYBM 
         Height          =   270
         Index           =   2
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "1"
         Top             =   900
         Width           =   405
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73800
         MaxLength       =   3
         TabIndex        =   8
         Top             =   405
         Width           =   500
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72990
         MaxLength       =   3
         TabIndex        =   9
         Top             =   405
         Width           =   500
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   400
         Left            =   -68280
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtYBM 
         Height          =   270
         Index           =   3
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "99.99"
         Top             =   1240
         Width           =   550
      End
      Begin VB.TextBox txtYBM 
         Height          =   270
         Index           =   1
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "999"
         Top             =   520
         Width           =   500
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170024.frx":0038
         Height          =   3495
         Left            =   -74990
         TabIndex        =   11
         Top             =   795
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6174
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
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   120
         TabIndex        =   13
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "        今日在職員工的公司別有："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   420
         TabIndex        =   16
         Top             =   2520
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PS：未輸入的公司別都以 2 公司的獎金月數計算年終獎金"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   15
         Top             =   2160
         Width           =   4500
      End
      Begin MSForms.Label lblDsp 
         Height          =   300
         Index           =   2
         Left            =   1830
         TabIndex        =   14
         Top             =   945
         Width           =   2500
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "獎金年度：              －"
         Height          =   180
         Left            =   -74760
         TabIndex        =   12
         Top             =   450
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "獎金月數： "
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   7
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別："
         Height          =   180
         Left            =   420
         TabIndex        =   5
         Top             =   910
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "獎金年度："
         Height          =   180
         Index           =   4
         Left            =   420
         TabIndex        =   4
         Top             =   555
         Width           =   900
      End
   End
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
            Picture         =   "frm170024.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170024.frx":1E25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   470
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8200
      _ExtentX        =   14464
      _ExtentY        =   829
      ButtonWidth     =   1076
      ButtonHeight    =   794
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
Attribute VB_Name = "frm170024"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/23 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/29 Modify by sindy
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_YBM As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean

Private Sub cmdok_Click()
   If txt1(0) & txt1(1) <> "" Then
      If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
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
      stCon = stCon & " and ybm01>='" & Val(txt1(0)) + 1911 & "' "
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and ybm01<='" & Val(txt1(1)) + 1911 & "' "
   End If
   strExc(0) = "SELECT ybm01-1911,a0801||' '||a0802,ybm03 FROM YearBonusMonth,acc080 " & _
                     " where ybm02=a0801(+) " & stCon & " order by ybm01,ybm02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
      SetGrd
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
'2015/1/21 ADD BY SONIA
Dim rs As New ADODB.Recordset
Dim CompList As String
'2015/1/21 END

   
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
   
   '2015/1/21 ADD BY SONIA
   CompList = ""
   rs.CursorLocation = adUseClient
   rs.Open "SELECT DISTINCT SD19 FROM SALARYDATA,STAFF WHERE SD01=ST01(+) AND ST04='1'", cnnConnection, adOpenStatic, adLockReadOnly
   While Not rs.EOF
      CompList = CompList & rs.Fields(0).Value & ", "
      rs.MoveNext
   Wend
   If rs.State <> adStateClosed Then rs.Close
   Set rs = Nothing
   CompList = Left(CompList, Len(CompList) - 2)
   Label1(2) = "        今日在職員工的公司別有：" & CompList
   '2015/1/21 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170024 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from YearBonusMonth where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_YBM = .Fields.Count
      ReDim m_FieldList(TF_YBM) As FIELDITEM
      For Each oText In txtYBM
         idx = oText.Index
         m_FieldList(idx).fiName = "YBM" & Format(idx, "00")
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
   
   stKey01 = Val(txtYBM(1)) + 1911
   stKey02 = txtYBM(2)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM YearBonusMonth" & _
            " WHERE ybm01 = '" & stKey01 & "' and ybm02= '" & stKey02 & "' "
      Case -2
         strExc(0) = "SELECT * FROM YearBonusMonth order by 1 ASC,2 ASC"
      Case -1
         strExc(0) = "SELECT * FROM YearBonusMonth" & _
            " WHERE ybm01||ybm02 <'" & stKey01 & stKey02 & "' order by 1 DESC,2 DESC"
      Case 1
         strExc(0) = "SELECT * FROM YearBonusMonth" & _
            " WHERE ybm01||ybm02 >'" & stKey01 & stKey02 & "' order by 1 ASC,2 ASC"
      Case 2
         strExc(0) = "SELECT * FROM YearBonusMonth order by 1 DESC,2 DESC"
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
      txtYBM(1).SetFocus
      txtYBM_GotFocus 1
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
            If txtYBM(1).Locked = False Then
               txtYBM(1).Text = GRD1.TextMatrix(lCurRow, 0)
               txtYBM(2).Text = Trim(Left(Trim(GRD1.TextMatrix(lCurRow, 1)), 2))
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
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 3
         KeyAscii = Pub_NumAscii(KeyAscii, True)
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtYBM_GotFocus(Index As Integer)
   TextInverse txtYBM(Index)
   CloseIme
End Sub

Private Sub ClearField()
   For Each oText In txtYBM
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_YBM
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
      For Each oText In txtYBM
         idx = oText.Index
         '所得年度轉民國年
         If idx = 1 Then
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName) - 1911
         ElseIf idx = 7 Then
            m_FieldList(idx).fiOldData = TransDate("" & .Fields(m_FieldList(idx).fiName), 1)
         Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End If
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
      Next
            
      lblDsp(2) = CompNameQuery(txtYBM(2))
      
      CUID(1) = "" & .Fields("ybm04")
      CUID(2) = "" & .Fields("ybm05")
      CUID(3) = "" & .Fields("ybm06")
      CUID(4) = "" & .Fields("ybm07")
      CUID(5) = "" & .Fields("ybm08")
      CUID(6) = "" & .Fields("ybm09")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtYBM(1).Tag = txtYBM(1)
   txtYBM(2).Tag = txtYBM(2)
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtYBM
      oText.Locked = bLocked
   Next
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
            txtYBM(1) = txtYBM(1).Tag
            txtYBM(2) = txtYBM(2).Tag
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
         If m_bUpdate And txtYBM(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtYBM(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtYBM(1) <> "" Then
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
         txtYBM(1).Locked = False
         If Me.Visible = True Then
            txtYBM(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtYBM(1).Locked = True
         If Me.Visible = True Then
            txtYBM(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 4
         txtYBM(1).Locked = False
         txtYBM(2).Locked = False
         If Me.Visible = True Then
            txtYBM(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtYBM(1).Locked = True
         If Me.Visible = True Then
            txtYBM(1).SetFocus
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
               txtYBM(1).SetFocus
               txtYBM_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtYBM
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtYBM_Validate idx, bCancel
         If bCancel = True Then
            txtYBM(idx).SetFocus
            txtYBM_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtYBM(1) = "" Then
         ShowMsg "請輸入獎金年度 !"
         txtYBM(1).SetFocus
         txtYBM_GotFocus 1
         GoTo EscPoint
      End If
      If txtYBM(2) = "" Then
         ShowMsg "請輸入公司別 !"
         txtYBM(2).SetFocus
         txtYBM_GotFocus 2
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtYBM(1) = "" And txtYBM(1).Locked = False Then
         ShowMsg "請輸入獎金年度 !"
         txtYBM(1).SetFocus
         txtYBM_GotFocus 1
         GoTo EscPoint
      End If
      If txtYBM(2) = "" And txtYBM(2).Locked = False Then
         ShowMsg "請輸入公司別 !"
         txtYBM(2).SetFocus
         txtYBM_GotFocus 2
         GoTo EscPoint
      End If
      If txtYBM(3) = "" And txtYBM(3).Locked = False Then
         ShowMsg "請輸入獎金月數 !"
         txtYBM(3).SetFocus
         txtYBM_GotFocus 3
         GoTo EscPoint
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
   For Each oText In txtYBM
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO YearBonusMonth (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
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
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE YearBonusMonth SET "
   stSet = ""
   For Each oText In txtYBM
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
      stSQL = stSQL & stSet & " where ybm01=" & Val(txtYBM(1)) + 1911 & " and ybm02='" & txtYBM(2) & "' ; end; "
      
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
   For Each oText In txtYBM
      idx = oText.Index
      Select Case idx
         Case 1
            m_FieldList(idx).fiNewData = Val(oText.Text) + 1911
         Case 7
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtYBM_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 3
         KeyAscii = Pub_NumAscii(KeyAscii, True)
      Case Else
'2008/12/31 cancel by sonia
'         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
'            KeyAscii = 0
'            Beep
'         End If
'2008/12/31 end
   End Select
End Sub

Private Sub txtYBM_Validate(Index As Integer, Cancel As Boolean)
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 1
            If txtYBM(Index) <> "" Then
               If ChkDate(txtYBM(Index) & "0101") = False Then
                  Cancel = True
               End If
            End If
         Case 2
            If txtYBM(Index) <> "" Then
               lblDsp(2) = CompNameQuery(txtYBM(Index))
               If lblDsp(2) = "" Then
                  ShowMsg "公司別錯誤 !"
                  Cancel = True
               End If
            Else
               lblDsp(2) = ""
            End If
         Case 3
            '2008/12/31 add by sonia
            If txtYBM(Index) <> "" Then
               If Val(txtYBM(Index)) > 99.9 Then
                  ShowMsg "獎金月數不可大於 99.9 !"
                  Cancel = True
               End If
            Else
               lblDsp(2) = ""
            End If
            '2008/12/31 end
      End Select
      
      If Cancel = True Then TextInverse txtYBM(Index)
      
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
   stSQL = "delete from YearBonusMonth where ybm01=" & Val(txtYBM(1)) + 1911 & " and ybm02='" & txtYBM(2) & "' "
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtYBM(1).Tag = ""
   txtYBM(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
End Function

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iCol As Integer
   
   '格式顯示中文,代號隱藏
   arrGridHeadText = Array("年度", "公司別", "獎金月數")
   arrGridHeadWidth = Array(900, 3600, 1200)
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
