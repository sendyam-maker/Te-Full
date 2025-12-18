VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1104 
   BorderStyle     =   1  '單線固定
   Caption         =   "多國案卷號關係建立"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   996
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdOK 
      Caption         =   "清除關聯"
      Height          =   400
      Index           =   2
      Left            =   3645
      TabIndex        =   30
      Top             =   120
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm1104.frx":0000
      Height          =   1020
      Left            =   90
      TabIndex        =   29
      Top             =   4620
      Width           =   9105
      _ExtentX        =   16066
      _ExtentY        =   1799
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "C00"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "C01"
         Caption         =   "專利種類"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "C02"
         Caption         =   "案件名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "C03"
         Caption         =   "申請國家"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "C04"
         Caption         =   "申請人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1428.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2448
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2928.189
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdMove2 
      Caption         =   "刪除"
      Height          =   320
      Index           =   5
      Left            =   4410
      TabIndex        =   26
      Top             =   3645
      Width           =   800
   End
   Begin VB.CommandButton cmdMove2 
      Caption         =   "新增"
      Height          =   320
      Index           =   4
      Left            =   3600
      TabIndex        =   25
      Top             =   3645
      Width           =   800
   End
   Begin VB.TextBox txtInCase 
      Height          =   264
      Index           =   0
      Left            =   675
      MaxLength       =   3
      TabIndex        =   21
      Text            =   "P"
      Top             =   3675
      Width           =   492
   End
   Begin VB.TextBox txtInCase 
      Height          =   264
      Index           =   1
      Left            =   1155
      MaxLength       =   6
      TabIndex        =   22
      Top             =   3675
      Width           =   852
   End
   Begin VB.TextBox txtInCase 
      Height          =   264
      Index           =   2
      Left            =   1995
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3675
      Width           =   252
   End
   Begin VB.TextBox txtInCase 
      Height          =   264
      Index           =   3
      Left            =   2235
      MaxLength       =   2
      TabIndex        =   24
      Top             =   3675
      Width           =   372
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "尋找(&F)"
      Height          =   320
      Index           =   3
      Left            =   8280
      TabIndex        =   19
      Top             =   630
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Height          =   1965
      Left            =   45
      TabIndex        =   18
      Top             =   1590
      Width           =   9165
      _ExtentX        =   16171
      _ExtentY        =   3471
      _Version        =   393216
      FixedCols       =   0
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
   Begin VB.CommandButton cmdMove 
      Caption         =   "清除(&C)"
      Height          =   320
      Index           =   2
      Left            =   6744
      TabIndex        =   7
      Top             =   630
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      Height          =   320
      Index           =   1
      Left            =   5916
      TabIndex        =   6
      Top             =   630
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "新增(&A)"
      Height          =   320
      Index           =   0
      Left            =   5088
      TabIndex        =   5
      Top             =   630
      Width           =   800
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2295
      TabIndex        =   11
      Top             =   630
      Width           =   2532
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   0
         Left            =   0
         MaxLength       =   6
         TabIndex        =   1
         Top             =   0
         Width           =   1212
      End
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   2
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   3
         Top             =   0
         Width           =   492
      End
   End
   Begin VB.TextBox txtSystem 
      Height          =   288
      Left            =   1536
      MaxLength       =   3
      TabIndex        =   0
      Top             =   630
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8280
      TabIndex        =   9
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7485
      TabIndex        =   8
      Top             =   120
      Width           =   800
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5490
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
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
      Caption         =   "Adodc2"
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
   Begin MSForms.ComboBox cboCaseName1 
      Height          =   300
      Left            =   1050
      TabIndex        =   31
      Top             =   3990
      Width           =   8070
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14235;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1050
      TabIndex        =   4
      Top             =   960
      Width           =   8070
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14235;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "＊代表已閉卷"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7920
      TabIndex        =   38
      Top             =   3570
      Width           =   1245
   End
   Begin VB.Label Label10 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   90
      TabIndex        =   37
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "申請人："
      Height          =   255
      Left            =   4140
      TabIndex        =   36
      Top             =   4320
      Width           =   855
   End
   Begin MSForms.Label lblCustomer1 
      Height          =   255
      Left            =   4980
      TabIndex        =   35
      Top             =   4320
      Width           =   4095
      VariousPropertyBits=   27
      Size            =   "7223;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNation1 
      Height          =   255
      Left            =   1110
      TabIndex        =   34
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label lblNationName1 
      Height          =   255
      Left            =   1650
      TabIndex        =   33
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   32
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label lblMainCase 
      Height          =   255
      Left            =   1260
      TabIndex        =   28
      Top             =   233
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "原多國案號："
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   27
      Top             =   233
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "與國內                                              案號相同"
      Height          =   180
      Index           =   14
      Left            =   90
      TabIndex        =   20
      Top             =   3720
      Width           =   3330
   End
   Begin VB.Label lblNationName 
      Height          =   255
      Left            =   1650
      TabIndex        =   17
      Top             =   1290
      Width           =   2295
   End
   Begin VB.Label lblNation 
      Height          =   255
      Left            =   1110
      TabIndex        =   16
      Top             =   1290
      Width           =   495
   End
   Begin MSForms.Label lblCustomer 
      Height          =   255
      Left            =   4980
      TabIndex        =   15
      Top             =   1290
      Width           =   4095
      VariousPropertyBits=   27
      Size            =   "7223;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "申請人："
      Height          =   255
      Left            =   4140
      TabIndex        =   14
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   90
      TabIndex        =   13
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   12
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "相關之本所案號："
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   630
      Width           =   1455
   End
End
Attribute VB_Name = "frm1104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/11 改成Form2.0 (MSFlexGrid1,DataGrid2,cboCaseName,cboCaseName1,lblCustomer,lblCustomer1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/8/18 日期欄已修改
Option Explicit

'intWhereComeFrom  1:frm050101_2/frm060101_1     2:Others
Public intWhereComeFrom As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, intLeaveKind As Integer
Public m_form As Form
'Add by Morgan 2006/7/3
Dim m_bolCombine As Boolean '是否有合併群組
Dim m_CombGroupCase() As String '合併群組案號資料
Dim m_CP(1 To 4) As String '本所號
Dim strSeqNo As String 'Add by Amy 2014/06/06 暫存TB序號
Dim m_Sys As String 'Added by Morgan 2021/2/25
Public m_CRL01 As String 'Add By Sindy 2022/11/23


'Modify by Morgan 2007/11/19 國內案不再限制 "P"
Private Sub cmdMove2_Click(Index As Integer)
   Dim varSaveCursor
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
         
   Select Case Index
      Case 4
         If CheckKeyIn3(3) = True Then
            Grid2Add 2
         End If
         
      Case 5
         Grid2Remove
         
   End Select
   
   Screen.MousePointer = varSaveCursor
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   InitialGrid
   Me.cmdMove(1).Enabled = False
   Me.cmdMove(2).Enabled = False
   InitGrid 'Add by Morgan 2006/9/12 連結資料表
End Sub

Private Sub Form_Activate()
   txtSystem.SetFocus
   m_bolCombine = False
End Sub
'Add by Morgan 2005/5/19 CFP分案讀取相關卷號資料
Public Sub GetRelation(Optional p_bMsg As Boolean = False)
   Me.MSFlexGrid1.Visible = False
   Me.DataGrid2.Visible = False
   Grid1Search
   Me.MSFlexGrid1.Visible = True
   Me.DataGrid2.Visible = True
   'Add By Sindy 2022/11/23
   If m_CRL01 <> MsgText(601) Then
      strExc(10) = Pub_GetCRLCaseMap(m_CRL01, "0", "CFP", m_CP(1), m_CP(2), m_CP(3), m_CP(4))
      If strExc(10) <> "" Then 'And m_CP(2) <> "" And InStr(strExc(10), m_CP(1)) = 0 And InStr(strExc(10), m_CP(2)) = 0
         txtSystem = SystemNumber(strExc(10), 1)
         txtCode(0) = SystemNumber(strExc(10), 2)
         txtCode(1) = SystemNumber(strExc(10), 3)
         txtCode(2) = SystemNumber(strExc(10), 4)
      End If
   End If
   '2022/11/23 END
End Sub

Private Sub cmdMove_Click(Index As Integer)
   Dim varSaveCursor
   
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
         
   Select Case Index
   
      Case 0 '新增
         Grid1Add 2

      Case 1 '刪除
         Grid1Remove
         
      Case 2 '清除
         Grid1Clear
                        
      Case 3 '尋找
         Grid1Search 2
         
   End Select
   
   Screen.MousePointer = varSaveCursor
   
End Sub
'新增國外案
'p_Mode:1=程式,2=按鈕
Private Sub Grid1Add(Optional ByVal p_Mode As Integer = 1)

   Dim bolRt As Boolean, i As Integer
   Dim strCaseName As String, strCaseCode As String
   Dim bolNoOtherGroup As Boolean, stPA57 As String
   
   bolRt = CheckKeyIn2(2)
   If bolRt Then
      strCaseCode = txtSystem + "-" + txtCode(0) + IIf(txtCode(1) = "", "-0", "-" + txtCode(1)) + IIf(txtCode(2) = "", "-00", "-" + txtCode(2))
      '檢查國外案號是否已存在
      For i = 1 To Me.MSFlexGrid1.Rows - 1
         If Left(MSFlexGrid1.TextMatrix(i, 1), 15) = strCaseCode Then
            Exit For
         End If
      Next
      If i = Me.MSFlexGrid1.Rows Then
         If p_Mode = 2 Then
            'Add by Morgan 2006/6/3 檢查該案號是否已存在於另一相關群組
            bolNoOtherGroup = ChkNoOtherGroup(strCaseCode)
         End If
         If p_Mode = 1 Or bolNoOtherGroup = True Then
            strCaseName = Mid(cboCaseName.List(cboCaseName.ListIndex), 3)
            'Modified by Morgan 2022/1/11 第1筆資料非空白才要加
            'Me.MSFlexGrid1.Rows = Me.MSFlexGrid1.Rows + 1
            If MSFlexGrid1.TextMatrix(1, 1) <> "" Then
               Me.MSFlexGrid1.Rows = Me.MSFlexGrid1.Rows + 1
            End If
            'end 2022/1/11
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 2) = GetPA08(strCaseCode, stPA57)
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 1) = strCaseCode & IIf(stPA57 = "Y", "＊", "")
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 3) = strCaseName
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 4) = lblNationName.Caption
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 5) = lblCustomer.Caption
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 6) = GetCP21(strCaseCode)
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 7) = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 6)
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 8) = lblNation.Caption
            strExc(0) = Right(String(3, " ") & Replace(strCaseCode, "-", ""), 12)
            strExc(1) = Trim(Left(strExc(0), 3))
            strExc(2) = Mid(strExc(0), 4, 6)
            strExc(3) = Mid(strExc(0), 10, 1)
            strExc(4) = Mid(strExc(0), 11, 2)
            SetInCase strExc
         End If
         txtSystem = ""
         txtCode(0) = ""
         txtCode(1) = ""
         txtCode(2) = ""
      Else
         ShowMsg MsgText(8005)
      End If
   End If
   
   'Modified by Morgan 2022/1/11
   'If Me.MSFlexGrid1.Rows <= 1 Then
   If MSFlexGrid1.TextMatrix(1, 1) = "" Then
   'end 2022/1/11
       Me.cmdMove(1).Enabled = False
       Me.cmdMove(2).Enabled = False
   Else
      Me.cmdMove(1).Enabled = True
      Me.cmdMove(2).Enabled = True
   End If
   
End Sub

Private Sub Grid1Remove()
   Dim ii As Integer
   
   If Me.MSFlexGrid1.Rows > 1 Then
       For ii = Me.MSFlexGrid1.Rows - 1 To 1 Step -1
           If Me.MSFlexGrid1.TextMatrix(ii, 0) = "V" Then
               If Me.MSFlexGrid1.Rows = 2 Then
                   'Modified by Morgan 2022/1/11
                   'Me.MSFlexGrid1.Rows = 1
                   InitialGrid
                   'end 2022/1/11
                   Exit For
               Else
                  '加判斷若案號合併自其他群組時相關案號也要一併刪除
                  If m_bolCombine = True Then
                     For intI = 0 To UBound(m_CombGroupCase, 2)
                        If m_CombGroupCase(0, intI) & m_CombGroupCase(1, intI) & m_CombGroupCase(2, intI) & m_CombGroupCase(3, intI) = Replace(Left(MSFlexGrid1.TextMatrix(ii, 1), 15), "-", "") Then
                           CheckAllGroupCase
                           m_bolCombine = False
                           ii = Me.MSFlexGrid1.Rows
                           Exit For
                        End If
                     Next
                     If m_bolCombine = True Then
                        Me.MSFlexGrid1.RemoveItem ii
                     End If
                  Else
                     Me.MSFlexGrid1.RemoveItem ii
                  End If
               End If
           End If
       Next ii
   End If
   'Modified by Morgan 2022/1/11
   'If Me.MSFlexGrid1.Rows <= 1 Then
   If MSFlexGrid1.TextMatrix(1, 1) = "" Then
   'end 2022/1/11
       Me.cmdMove(1).Enabled = False
       Me.cmdMove(2).Enabled = False
   Else
      Me.cmdMove(1).Enabled = True
      Me.cmdMove(2).Enabled = True

   End If
   
End Sub

Private Sub Grid1Clear()
   'Modified by Morgan 2022/1/11
   'MSFlexGrid1.Rows = 1
   InitialGrid
   'end 2022/1/11
   cmdMove(1).Enabled = False
   cmdMove(2).Enabled = False
   m_bolCombine = False
   Erase m_CP
   Erase m_CombGroupCase
   lblMainCase = ""
   InitGrid
End Sub
'p_Mode:1=程式,2=按鈕
Private Sub Grid1Search(Optional ByVal p_Mode As Integer = 1)
   Dim strRelation() As String
   Dim iRtn As Integer
   Dim i As Integer
   
   Grid1Clear
   
   If txtSystem = "" Then
      txtSystem.SetFocus
      Exit Sub
   End If
   If txtCode(0) = "" Then
      txtCode(0).SetFocus
      Exit Sub
   End If
   txtCode(1) = Right("0" & txtCode(1), 1)
   txtCode(2) = Right("00" & txtCode(2), 2)
   
   strExc(0) = txtSystem
   strExc(1) = txtCode(0)
   strExc(2) = txtCode(1)
   strExc(3) = txtCode(2)
   
   iRtn = oReadCaseRelationData(strExc(0), strExc(1), strExc(2), strExc(3), strRelation())
   If p_Mode = 1 Or iRtn = 1 Then
      '加入自己
      m_CP(1) = txtSystem
      m_CP(2) = txtCode(0)
      m_CP(3) = txtCode(1)
      m_CP(4) = txtCode(2)
      lblMainCase = m_CP(1) & "-" & m_CP(2) & "-" & m_CP(3) & "-" & m_CP(4)
      Grid1Add
      
   'Add by Morgan 2006/10/14
   '只有國內案
   Else
      strExc(0) = "select * from casemap where cm10='0' and cm01='" & txtSystem & "' and cm02='" & txtCode(0) & "' and cm03='" & txtCode(1) & "' and cm04='" & txtCode(2) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '加入自己
         m_CP(1) = txtSystem
         m_CP(2) = txtCode(0)
         m_CP(3) = txtCode(1)
         m_CP(4) = txtCode(2)
         lblMainCase = m_CP(1) & "-" & m_CP(2) & "-" & m_CP(3) & "-" & m_CP(4)
         Grid1Add
      End If
   End If
   If iRtn = 1 Then
      '加入其他多國案
      For i = 0 To UBound(strRelation, 2)
         txtSystem = strRelation(0, i)
         txtCode(0) = strRelation(1, i)
         txtCode(1) = strRelation(2, i)
         txtCode(2) = strRelation(3, i)
         Grid1Add
      Next
   End If
      
   'Add by Morgan 2006/10/14
   '依照本所號排序
   Me.MSFlexGrid1.col = 1
   Me.MSFlexGrid1.Sort = "1"
   
   'Modified by Morgan 2022/1/11
   'If Me.MSFlexGrid1.Rows <= 1 Then
   If MSFlexGrid1.TextMatrix(1, 1) = "" Then
   'end 2022/1/11
      Me.cmdMove(1).Enabled = False
      Me.cmdMove(2).Enabled = False
   Else
      Me.cmdMove(1).Enabled = True
      Me.cmdMove(2).Enabled = True
   End If
   
End Sub
'移除國內案號
Private Sub Grid2Remove()
   strExc(1) = txtInCase(0)
   strExc(2) = txtInCase(1)
   strExc(3) = Right("0" & txtInCase(2), 1)
   strExc(4) = Right("00" & txtInCase(3), 2)
   If strExc(2) <> "" Then
      If Adodc2.Recordset.RecordCount > 0 Then
         Adodc2.Recordset.MoveFirst
         Adodc2.Recordset.Find "C00='" & strExc(1) & "-" & strExc(2) & "-" & strExc(3) & "-" & strExc(4) & "'"
         If Not Adodc2.Recordset.EOF Then
            Adodc2.Recordset.Delete
            Adodc2.Recordset.UpdateBatch
            txtInCase(1) = ""
            txtInCase(2) = ""
            txtInCase(3) = ""
         End If
      End If
   End If
End Sub

Private Function GetCRList() As String
   Dim i As Integer, strCaseNo As String
   Dim strTemp
   Dim strCRList As String
   
   For i = 1 To Me.MSFlexGrid1.Rows - 1
      strTemp = Split(Left(Me.MSFlexGrid1.TextMatrix(i, 1), 15), "-")
      strCRList = strCRList & IIf(strCRList <> "", " UNION ", "") & " SELECT '" & strTemp(0) & "' CM1,'" & strTemp(1) & "' CM2,'" & strTemp(2) & "' CM3,'" & strTemp(3) & "' CM4 FROM DUAL"
   Next
   GetCRList = strCRList
   
End Function
'Add by Morgan 2006/7/4
Private Function SaveDataNew() As Boolean

   Dim bolNoUpdate As Boolean '是否不更新國外案新穎性期限
   Dim strTemp
   Dim i As Integer, j As Integer
   Dim strRelation() As String
   Dim strCRList As String
   Dim strCMList As String
   Dim strCNo(1 To 4) As String, strPrePA14 As String, strCP07 As String, strCP06 As String
   Dim strUpdCNo(1 To 4) As String, strUpdCNo2(1 To 4) As String, strUpdMsg As String, strTmpMsg As String
   Dim i101 As Integer '美國案index
   Dim strCaseNo(1 To 4) As String '美國案本所案號
   Dim str101PA12 As String '美國案公開日
   Dim str101PA46 As String '美國是否PCT案
   Dim str101PA08 As String '美國是否發明案 Added by Morgan 2021/2/22
   Dim strCP48 As String '承辦期限
   Dim strCP14 As String '承辦人
   Dim bolUpdAll As Boolean '已會稿完成皆可更新國外案齊備日
   'Modify by Morgan 2009/8/13 以公告日更新改呼叫共用函式
   'Dim strCP64 As String '期限更新備註 Add by Morgan 2009/5/11
   Dim strMsg As String
   Dim strCP29 As String, strST04 As String, strST06 As String 'Add by Morgan 2011/10/21
   Dim m_list() As String '相關案 Added by Morgan 2012/3/26
   Dim bolPublished As Boolean '是否已公開或公告 Added by Morgan 2015/9/22
   Dim strTwPA20 As String '臺灣案核准日'Added by Morgan 2015/12/18
   Dim bolJpUpd As Boolean 'Added by Morgan 2021/2/5 更新日本案齊備確認
   
   If Me.MSFlexGrid1.Rows > 1 Then
      ReDim strRelation(3, Me.MSFlexGrid1.Rows - 2)
   End If
   
   '國外案號
   strCRList = ""
   '美國案
   i101 = -1
   With Me.MSFlexGrid1
   For i = 1 To .Rows - 1
      strTemp = Split(Left(.TextMatrix(i, 1), 15), "-")
      For j = 0 To 3
        strRelation(j, i - 1) = strTemp(j)
      Next
      strCRList = strCRList & IIf(strCRList <> "", " UNION ", "") & " SELECT '" & strRelation(0, i - 1) & "' CM1,'" & strRelation(1, i - 1) & "' CM2,'" & strRelation(2, i - 1) & "' CM3,'" & strRelation(3, i - 1) & "' CM4 FROM DUAL"
      If .TextMatrix(i, 8) = "101" Then
         i101 = i - 1
      End If
   Next
   End With
   
   '國內案號
   strCMList = ""
   With Adodc2.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         strTemp = Split(.Fields(0), "-")
         strCMList = strCMList & IIf(strCMList <> "", " UNION ", "") & " SELECT '" & strTemp(0) & "' CM5,'" & strTemp(1) & "' CM6,'" & strTemp(2) & "' CM7,'" & strTemp(3) & "' CM8 FROM DUAL"
         .MoveNext
      Loop
   End If
   End With
   
   'Add by Morgan 2006/10/14
   '倘若P案公告後才收文要辦理CFP案件,則程序人員鍵關聯時,SHOW訊息提醒USER是否要UPDATE期限至CFP案內
   If strCMList <> "" Then
      bolNoUpdate = False
      bolPublished = False 'Added by Morgan 2015/9/22
      'Modified by Morgan 2012/8/23 所有相同案都要考慮 Ex.CFP-24471 --> CFP-25322
      'strExc(0) = "SELECT PA14,PA01,PA02,PA03,PA04 FROM PATENT WHERE (PA01,PA02,PA03,PA04) in (" & strCMList & ") AND PA14>0"
      'Modified by Morgan 2015/9/21 +公開
      'Modified by Morgan 2016/2/19 美國會輸入預定公開日,改判斷公告或公開日小於等於系統日
      strExc(0) = "SELECT PA14,PA01,PA02,PA03,PA04,'公告' Memo FROM PATENT WHERE (PA01,PA02,PA03,PA04) in (" & strCMList & " union " & strCRList & " union select cm01,cm02,cm03,cm04 from casemap where (cm05,cm06,cm07,cm08) in (" & strCMList & ") " & ") AND PA14>0 and pa14<=" & strSrvDate(1)
      strExc(0) = strExc(0) & " union all SELECT PA12,PA01,PA02,PA03,PA04,'公開' Memo FROM PATENT WHERE (PA01,PA02,PA03,PA04) in (" & strCMList & " union " & strCRList & " union select cm01,cm02,cm03,cm04 from casemap where (cm05,cm06,cm07,cm08) in (" & strCMList & ") " & ") AND PA12>0 and pa12<=" & strSrvDate(1)
      'end 2015/9/21
      strExc(0) = strExc(0) & " order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modified by Morgan 2012/8/23
         'strExc(0) = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS C1 WHERE (CP01,CP02,CP03,CP04) IN (" & strCRList & ")" & _
                  " AND CP27 IS NULL AND CP57 IS NULL AND CP10 IN (" & CaseMapOut & ")" & _
                  " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C2 WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP10='106' AND C2.CP57 IS NULL)"
         'Modified by Morgan 2015/9/21 +美國案例外(用公開日/公告日+1年更新)
         strExc(0) = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS C1 WHERE (CP01,CP02,CP03,CP04) IN (" & strCRList & " UNION " & strCMList & ")" & _
                  " AND CP27 IS NULL AND CP57 IS NULL AND CP10 IN (" & SameCaseProperty4Update & ")" & _
                  " AND NOT EXISTS(SELECT * FROM patent WHERE PA01=C1.CP01 AND PA02=C1.CP02 AND PA03=C1.CP03 AND PA04=C1.CP04 AND (PA46='Y' or PA09='101'))" & _
                  " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C2 WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP10='106' AND C2.CP57 IS NULL)"
         intI = 1
         Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Morgan 2012/8/23
            'Modified by Morgan 2018/1/23 已公告/公開不更新期限,訊息改只提醒--郭
            'If MsgBox("相同案 " & RsTemp(1) & "-" & RsTemp(2) & "-" & RsTemp(3) & "-" & RsTemp(4) & " 已於 " & Format(RsTemp(0) - 19110000, "##/##/##") & " " & RsTemp("memo") & "，是否要更新期限至其他未發文新案？", vbYesNo + vbDefaultButton2) = vbNo Then
            '   bolNoUpdate = True
            If MsgBox("相同案 " & RsTemp(1) & "-" & RsTemp(2) & "-" & RsTemp(3) & "-" & RsTemp(4) & " 已於 " & Format(RsTemp(0) - 19110000, "##/##/##") & " " & RsTemp("memo") & "，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            'end 2018/1/23
            End If
            bolNoUpdate = True 'Added by Morgan 2022/6/20 不更新期限應該要設 True
            'end 2012/8/23
         End If
         
         'Modified by Morgan 2012/8/23
         strCP07 = RsTemp(0)
         strUpdCNo(1) = RsTemp(1)
         strUpdCNo(2) = RsTemp(2)
         strUpdCNo(3) = RsTemp(3)
         strUpdCNo(4) = RsTemp(4)
         strUpdMsg = "(" & RsTemp("memo") & "日)"
         'end 2012/8/23
         bolPublished = True 'Added by Morgan 2015/9/22
      End If
   End If
   
   'Add by Morgan 2007/8/21 美國案已公開提醒
   If i101 <> -1 Then
      strExc(0) = "SELECT PA12,PA46,PA08 FROM PATENT WHERE PA01='" & strRelation(0, i101) & "' AND PA02='" & strRelation(1, i101) & "' AND PA03='" & strRelation(2, i101) & "' AND PA04='" & strRelation(3, i101) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         str101PA08 = "" & RsTemp.Fields("pa08") 'Added by Morgan 2021/2/22
         'Add by Morgan 2008/7/15
         str101PA46 = "" & RsTemp.Fields("pa46")
         If Not IsNull(RsTemp.Fields("pa12")) Then
            str101PA12 = "" & RsTemp.Fields(0)
            If Val(RsTemp.Fields(0)) < Val(strSrvDate(1)) Then
               If MsgBox("美國案已於 " & Format(RsTemp(0) - 19110000, "##/##/##") & " 公開，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Function
               End If
            End If
         End If
      End If
   End If
   'end 2007/8/21
   
ReDo:

   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   '更新是否多國設定
   With Me.MSFlexGrid1
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 6) <> .TextMatrix(i, 7) Then
         strTemp = Split(Left(.TextMatrix(i, 1), 15), "-")
         strSql = "Update caseprogress set cp21='" & .TextMatrix(i, 6) & "' where " & ChgCaseprogress(Join(strTemp, "")) & " and cp10 in (" & CaseMapOut & ")"
         cnnConnection.Execute strSql, intI
         
         'Remove by Morgan 2011/5/4
         ''Modify by Morgan 2011/3/30
         ''100/4/1 以後多國案改草墨圖都要計件
         'If strSrvDate(1) < "20110401" Then
         '   '當[是否多國=Y]時,若繪圖尚未確認分案則草圖不計件上, 墨圖不計件
         '   If .TextMatrix(i, 6) = "Y" Then
         '      strSql = "Update EngineerProgress SET EP20='N',EP29='N' Where EP02 in (select cp09 from caseprogress where " & ChgCaseprogress(Join(strTemp, "")) & " and cp10 in (" & CaseMapOut & ") and cp107 is null)"
         '      cnnConnection.Execute strSql, intI
         '   '當[是否多國=NULL]時,若繪圖尚未確認分案則墨圖要計件
         '   Else
         '      strSql = "Update EngineerProgress SET EP29=NULL Where EP02 in (select cp09 from caseprogress where " & ChgCaseprogress(Join(strTemp, "")) & " and cp10 in (" & CaseMapOut & ") and cp107 is null)"
         '      cnnConnection.Execute strSql, intI
         '   End If
         'End If
         'end 2011/5/4
         
      End If
   Next
   End With
   
   '有舊群組
   If m_CP(1) <> "" Then
      '刪除舊群組的國內外關聯
      '被剔除的案號(不含原案號)
      strSql = "delete from CaseMap where CM10='0' AND CM01='CFP' AND (CM01,CM02,CM03,CM04) IN" & _
         " (SELECT '" & m_CP(1) & "' CM01,'" & m_CP(2) & "' CM02,'" & m_CP(3) & "' CM03,'" & m_CP(4) & "' CM04 FROM DUAL" & _
         " UNION SELECT CR05,CR06,CR07,CR08 FROM CASERELATION WHERE CR01='" & m_CP(1) & "'" & _
         " AND CR02='" & m_CP(2) & "' AND CR03='" & m_CP(3) & "' AND CR04='" & m_CP(4) & "')" & _
         ""
      
      If strCRList <> "" Then
         strSql = strSql & " AND NOT (CM01,CM02,CM03,CM04) IN (" & strCRList & ")"
      End If
      cnnConnection.Execute strSql, intI
      
      '刪除舊群組的多國關聯(修改前的)
      strSql = "delete from caserelation where (CR01,CR02,CR03,CR04) IN" + _
         " (SELECT '" & m_CP(1) & "' CR01,'" & m_CP(2) & "' CR02,'" & m_CP(3) & "' CR03,'" & m_CP(4) & "' CR04 FROM DUAL" & _
         " UNION SELECT CR05,CR06,CR07,CR08 FROM CASERELATION WHERE CR01='" & m_CP(1) & "'" & _
         " AND CR02='" & m_CP(2) & "' AND CR03='" & m_CP(3) & "' AND CR04='" & m_CP(4) & "')"
      cnnConnection.Execute strSql, intI
   End If
   
   If strCRList <> "" Then
      '刪除新群組的現有多國關聯
      strSql = "delete from caserelation where (cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) IN (SELECT X.*,Y.* FROM (" & strCRList & ") X,(" & strCRList & ") Y WHERE NOT (X.CM1=Y.CM1 AND X.CM2=Y.CM2 AND X.CM3=Y.CM3 AND X.CM4=Y.CM4))"
      cnnConnection.Execute strSql, intI
   
      strSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) SELECT X.*,Y.*" & _
         " FROM (" & strCRList & ") X,(" & strCRList & ") Y WHERE NOT (X.CM1=Y.CM1 AND X.CM2=Y.CM2 AND X.CM3=Y.CM3 AND X.CM4=Y.CM4)"
      cnnConnection.Execute strSql, intI
   End If
   
   '新群組國內外關聯
   If strCMList = "" Then
      If strCRList <> "" Then
         '刪除所有國內外關聯
         strSql = "delete from CaseMap where CM10='0' AND CM01='CFP' AND (CM01,CM02,CM03,CM04) IN (" & strCRList & ")"
         cnnConnection.Execute strSql, intI
         'Add by Morgan 2009/1/10 若CFP案無國內案且有案件已會稿完成則更新其他多國案的齊備日
         'Removed by Moragn 2016/5/5 移到下面(國內案會稿或會完改只更新多國主案故CFP主案若會完時不管有無國內案都要更新其他多國案的齊備日)
         'end 2009/1/10
      End If
   Else
      '刪除新群組的其他國內案的國內外關聯
      strSql = "delete from CaseMap where CM10='0' AND CM01='CFP' AND (CM01,CM02,CM03,CM04) IN (" & strCRList & ")" & _
         " AND NOT (CM05,CM06,CM07,CM08) IN (" & strCMList & ")"
      cnnConnection.Execute strSql, intI
   
      '刪除國內案與非新群組案號的國內外關聯
      strSql = "delete from CaseMap where CM10='0' AND CM01='CFP'" & _
         " AND (CM05,CM06,CM07,CM08) IN (" & strCMList & ")" & _
         " AND NOT (CM01,CM02,CM03,CM04) IN (" & strCRList & ")"
      cnnConnection.Execute strSql, intI
      
      '新增多國相關案件的國內外關聯
      strSql = "insert into casemap(cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,CM10,cm12,cm13,cm14)" & _
         " SELECT CM1,CM2,CM3,CM4,CM5,CM6,CM7,CM8,'0','" & strUserNum & "'" & _
         ",TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MI')" & _
         " FROM (" & strCRList & ") X, (" & strCMList & ") Y" & _
         " WHERE NOT EXISTS(SELECT * FROM CASEMAP WHERE CM01=CM1 AND CM02=CM2 AND CM03=CM3 AND CM04=CM4" & _
         " AND CM05=CM5 AND CM06=CM6 AND CM07=CM7 AND CM08=CM8 AND CM10='0')"
      cnnConnection.Execute strSql, intI
   
      If m_Sys = "CFP" Then 'Added by Morgan 2021/2/25 限定CFP案(FCP也會用)
   
'*****注意!!國內外及多國的控制要和承辦人系統的規則同步。*****
   
      'Add by Morgan 2006/9/14 檢查若所有國內案已發文則未齊備的國外案上齊備日
      'Modify by Morgan 2007/10/26 加判斷已會稿完成，96/11/1以後再加判斷已會稿
      'strExc(0) = "select count(*) from caseprogress where (cp01,cp02,cp03,cp04) in (" & strCMList & ") and cp27 is null and cp57 is null and cp10 in (" & CaseMapIn & ")"
      
      'Modify by Morgan 2009/1/10
      '若CFP案與P案的承辦人 [相同] 則以P案的會稿日為CFP案的齊備日
      '若CFP案與P案的承辦人 [不同] 則以P案的會稿完成日為CFP案的齊備日
      '若CFP案無國內案則以該案的會稿完成日更新其他多國案的齊備日 ->程式在上面
      'If Val(strSrvDate(1)) > 20071100 Then
      '   strExc(0) = "select count(*) from caseprogress,engineerprogress where (cp01,cp02,cp03,cp04) in (" & strCMList & ") and cp57 is null and cp10 in (" & CaseMapIn & ") and cp27 is null and ep02(+)=cp09 and ep08 is null and ep07 is null"
      'Else
      '   strExc(0) = "select count(*) from caseprogress,engineerprogress where (cp01,cp02,cp03,cp04) in (" & strCMList & ") and cp57 is null and cp10 in (" & CaseMapIn & ") and cp27 is null and ep02(+)=cp09 and ep08 is null"
      'End If
      'Modify by Morgan 2011/7/8 若國內案有國外P案已發文時CFP案也上齊備日　Ex.P-98504
      'strExc(0) = "select max(cp14) eng,min(sign(nvl(cp27,0)+nvl(ep08,0)+nvl(ep07,0))) sam,min(sign(nvl(cp27,0)+nvl(ep08,0))) dif  from caseprogress,engineerprogress where (cp01,cp02,cp03,cp04) in (" & strCMList & ") and cp57 is null and cp10 in (" & CaseMapIn & ") and ep02(+)=cp09 and cp14 is not null"
      strExc(0) = " select cp14,1,1,1 Srt from caseprogress where (cp01,cp02,cp03,cp04) in" & _
         " (select cm01,cm02,cm03,cm04 from casemap where cm01='P' and cm10='0' and (cm05,cm06,cm07,cm08) in (" & strCMList & "))" & _
         " and cp57 is null and cp10 in (" & CaseMapIn & ") and cp14 is not null and cp27>0" & _
         " union " & _
         " select max(cp14) eng,min(sign(nvl(cp27,0)+nvl(ep08,0)+nvl(ep07,0))) sam,min(sign(nvl(cp27,0)+nvl(ep08,0))) dif,2 Srt" & _
         " from caseprogress,engineerprogress where (cp01,cp02,cp03,cp04) in (" & strCMList & ")" & _
         " and cp57 is null and cp10 in (" & CaseMapIn & ") and ep02(+)=cp09 and cp14 is not null" & _
         " order by Srt"
      'end2011/7/8
      'end 2007/10/26
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modify by Morgan 2009/1/10
         'If RsTemp.Fields(0) = 0 Then
         '國內案承辦人
         strCP14 = "" & RsTemp.Fields(0)
         '沒有未會稿未會稿完成且未發文案件
         If strCP14 <> "" And RsTemp.Fields(1) > 0 Then
            '若皆已會稿完成時不必判斷承辦人
            If RsTemp.Fields(2) > 0 Then
               bolUpdAll = True
            Else
               bolUpdAll = False
            End If
            '未發文未齊備的國外案
            'Modified by Morgan 2016/5/5
            'Modified by Morgan 2016/5/5 +控制CFP只更新主案(判斷要計件者,因日本案可能非主案要計件但也要更新)--柄佑
            'Modified by Morgan 2016/5/17 因建關聯時大都還沒設不計件故改回判斷是否主案--柄佑 Ex.CFP-28655
            strExc(0) = "select ep02,cp06,cp14,cp21,cp26,pa09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNo,cp157 from caseprogress,Engineerprogress,patent where (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp27 is null and cp57 is null and  cp10 in (" & CaseMapOut & ") and ep02=cp09 and ep06 is null and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               adoRecordset.MoveFirst
               Do While Not adoRecordset.EOF
                  '已會稿完成,或相同承辦人(已會稿),若未分案時先視為相同承辦規則需再討論
                  'Modify by Morgan 2009/7/31 分案已加控制
                  'If bolUpdAll Or IsNull(adoRecordset.Fields("cp14")) Or (strCP14 = "" & adoRecordset.Fields("cp14")) Then
                  'Modified by Morgan 2016/5/17 +改判斷是主案或日本要計件案--柄佑
                  'Modified by Morgan 2021/2/23 日本案加判斷有北所分案日(因為分所案件可能會先上承辦人但是否要計件未上N)
                  If Not IsNull(adoRecordset.Fields("cp14")) And (bolUpdAll Or (strCP14 = adoRecordset.Fields("cp14"))) And (IsNull(adoRecordset.Fields("cp21")) Or (adoRecordset.Fields("pa09") = "011" And Not IsNull(adoRecordset.Fields("cp157")) And IsNull(adoRecordset.Fields("cp26")))) Then
                  
                     '更新齊備日(此處更新皆為事後才建關聯案件故用系統日)
                     strSql = "Update Engineerprogress set ep06=" & strSrvDate(1) & " where ep06 is null and ep02='" & adoRecordset.Fields(0).Value & "'"
                     cnnConnection.Execute strSql, intI
                     
                     If PUB_IfSetCP48() Then  'Add by Morgan 2010/10/6
                     
                        'Modify by Morgan 2007/10/12 承辦期限改呼叫共用函數計算
                        'strExc(0) = "Select NVL(CF04,0) From CaseProgress, Patent, Casefee Where CP09='" & adoRecordset.Fields(0).Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and PA01=cf01(+) and pa09=cf02(+) and cp10=cf03 "
                        strExc(0) = "Select cp01,pa09,cp10 From CaseProgress, Patent Where CP09='" & adoRecordset.Fields(0).Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
                        'End 2007/10/12
                        intI = 1
                        Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           'Modify by Morgan 2007/10/12 承辦期限改呼叫共用函數計算
                           'If AdoRecordSet3.Fields(0) > 0 Then
                           '   strExc(1) = CompWorkDay(AdoRecordSet3.Fields(0), strSrvDate(1), 0)
                           '   '若承辦期限大於本所期限
                           '   If Val(strExc(1)) > Val("" & adoRecordset.Fields(1)) And Val("" & adoRecordset.Fields(1)) > 0 Then
                           '      strExc(1) = adoRecordset.Fields(1)
                           '   End If
                           strCP48 = Pub_GetHandleDay(AdoRecordSet3("cp01"), AdoRecordSet3("pa09"), AdoRecordSet3("cp10"), , "" & adoRecordset.Fields(1), adoRecordset.Fields(0))
                           If strCP48 <> "" Then
                           'end 2007/10/12
                              'End 2007/10/11
                              '更新承辦期限
                              strSql = "Update caseprogress set cp48=" & strCP48 & " where cp09='" & adoRecordset.Fields(0) & "'"
                              cnnConnection.Execute strSql, intI
                           End If
                        End If
                        
                     End If '2010/10/6
                  
                  End If
                  adoRecordset.MoveNext
               Loop
            End If
         End If
         
      End If
      
      'Add by Morgan 2006/9/15 國內案領證已發文且未公告時需更新國外新案期限(排除有主張國際優先權的)
      If bolNoUpdate = False Then
         If bolPublished = False Then 'Added by Morgan 2015/9/22
         
            strExc(0) = "SELECT PA14,PA01,PA02,PA03,PA04,PA09,PA08,PA10,NA32,NA33,PA16,PA20 FROM PATENT,NATION WHERE (PA01,PA02,PA03,PA04) in (" & strCMList & ") AND PA10>0 AND NA01(+)=PA09"
            'Add by Morgan 2009/7/1
            '加國內案的國外P案
            strExc(0) = strExc(0) & " UNION SELECT p2.PA14,p2.PA01,p2.PA02,p2.PA03,p2.PA04,p2.PA09,p2.PA08,p2.PA10,NA32,NA33,p2.PA16,p2.PA20" & _
               " FROM PATENT p1,CASEMAP,PATENT p2,NATION" & _
               " WHERE (p1.PA01,p1.PA02,p1.PA03,p1.PA04) in (" & strCMList & ")" & _
               " and cm05(+)=p1.pa01 and cm06(+)=p1.pa02 and cm07(+)=p1.pa03 and cm08(+)=p1.pa04 AND CM10 IN ('0','3') and cm01='P'" & _
               " and p2.pa01(+)=cm01 and p2.pa02(+)=cm02 and p2.pa03(+)=cm03 and p2.pa04(+)=cm04" & _
               " AND p2.PA10>0 AND NA01(+)=p2.PA09"
               
            'Added by Morgan 2015/12/18
            '加國內案的國內案(一案兩請)
            strExc(0) = strExc(0) & " UNION SELECT p2.PA14,p2.PA01,p2.PA02,p2.PA03,p2.PA04,p2.PA09,p2.PA08,p2.PA10,NA32,NA33,p2.PA16,p2.PA20" & _
               " FROM PATENT p1,CASEMAP,PATENT p2,NATION" & _
               " WHERE (p1.PA01,p1.PA02,p1.PA03,p1.PA04) in (" & strCMList & ")" & _
               " and cm01(+)=p1.pa01 and cm02(+)=p1.pa02 and cm03(+)=p1.pa03 and cm04(+)=p1.pa04 AND CM10 IN ('0','3') and cm01='P'" & _
               " and p2.pa01(+)=cm05 and p2.pa02(+)=cm06 and p2.pa03(+)=cm07 and p2.pa04(+)=cm08" & _
               " AND p2.PA10>0 AND NA01(+)=p2.PA09"
            'end 2015/12/18
               
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Do While Not RsTemp.EOF
                  strPrePA14 = ""
                  strCNo(1) = RsTemp("PA01")
                  strCNo(2) = RsTemp("PA02")
                  strCNo(3) = RsTemp("PA03")
                  strCNo(4) = RsTemp("PA04")
                  
                  If Not IsNull(RsTemp("PA14")) Then
                     strPrePA14 = RsTemp("PA14")
                     strTmpMsg = "(公告日)"
                  ElseIf RsTemp("PA09") = "000" Then
                     strPrePA14 = PUB_GetPrePA14(strCNo)
                     strTmpMsg = "(預定公告日)"
                     'Added by Morgan 2015/12/18
                     'Modified by Morgan 2016/11/3 核准日--郭
                     'If strPrePA14 = "" And RsTemp("pa20") > 0 Then
                     If strPrePA14 = "" And RsTemp("pa16") = "1" And RsTemp("pa20") > 0 Then
                     'end 2016/11/3
                        If strTwPA20 = "" Or Val(strTwPA20) > RsTemp("pa20") Then
                           strUpdCNo2(1) = strCNo(1)
                           strUpdCNo2(2) = strCNo(2)
                           strUpdCNo2(3) = strCNo(3)
                           strUpdCNo2(4) = strCNo(4)
                           strTwPA20 = RsTemp("pa20")
                        End If
                     End If
                     'end 2015/12/18
                  'Add by Morgan 2009/7/1
                  '大陸核准日+5個月作為預定公告更新到多國案期限
                  ElseIf RsTemp("PA09") = "020" And RsTemp("PA16") = "1" And RsTemp("PA20") > 0 Then
                     strPrePA14 = CompDate(1, 5, RsTemp("PA20"))
                     strTmpMsg = "(預定公告日)"
                  End If
                  'Add by Morgan 2006/10/30
                  '沒有預估公告日時抓預估公開日
                  'Modify by Morgan 2007/1/30 都要抓用期限小的更新
                  'If strPrePA14 = "" Then
                     'Modify by Morgan 2008/7/15 要排除PCT
                     'If RsTemp("PA08") = "1" Then
                     If RsTemp("PA08") = "1" And RsTemp("PA09") <> "056" Then
                        strExc(0) = "SELECT PD05 FROM PRIDATE WHERE PD01='" & strCNo(1) & "' AND PD02='" & strCNo(2) & "' AND PD03='" & strCNo(3) & "' AND PD04='" & strCNo(4) & "' ORDER BY PD05 ASC"
                        intI = 1
                        Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           strExc(1) = "" & AdoRecordSet3("PD05")
                        Else
                           strExc(1) = "" & RsTemp("PA10")
                         End If
                        '法定期限=預估公開日=申請日(最早優先權日)+18個月
                        'Modify by Morgan 2007/1/30 都要抓用期限小的更新
                        'strPrePA14 = CompDate(1, 18, TransDate(strExc(1), 2))
                        'Modify by Morgan 2008/7/15 預估公開日改抓國家檔設定
                        'strExc(2) = CompDate(1, 18, TransDate(strExc(1), 2))
                        If "" & RsTemp("NA32") = "2" And Not IsNull(RsTemp("NA33")) Then
                           strExc(2) = CompDate(1, RsTemp("NA33"), TransDate(strExc(1), 2))
                        Else
                           strExc(2) = CompDate(1, 18, TransDate(strExc(1), 2))
                        End If
                        'END 2008/7/15
                        If strPrePA14 = "" Or Val(strExc(2)) < Val(strPrePA14) Then
                           strPrePA14 = strExc(2)
                           strTmpMsg = "(預定公開日)"
                        End If
                        'end 2007/1/30
                     End If
                  'End If
                  'end 2007/1/30
                  'end 2006/10/30
                  
                  If strPrePA14 <> "" Then
                     '法定期限=預估公告日
                     If strCP07 = "" Then
                        strCP07 = strPrePA14
                        'Add by Morgan 2009/5/11 +期限更新備註
                        'Modify by Morgan 2009/8/13 以公告日更新改呼叫共用函式
                        'strCP64 = "期限來源:" & Right("  " & strCNo(1), 3) & "-" & strCNo(2) & "-" & strCNo(3) & "-" & strCNo(4) & ";"
                        strUpdCNo(1) = strCNo(1)
                        strUpdCNo(2) = strCNo(2)
                        strUpdCNo(3) = strCNo(3)
                        strUpdCNo(4) = strCNo(4)
                        strUpdMsg = strTmpMsg
                     '取期限最小的
                     ElseIf Val(strCP07) > Val(strPrePA14) Then
                        strCP07 = strPrePA14
                        'Add by Morgan 2009/5/11 +期限更新備註
                        'Modify by Morgan 2009/8/13 以公告日更新改呼叫共用函式
                        'strCP64 = "期限來源:" & Right("  " & strCNo(1), 3) & "-" & strCNo(2) & "-" & strCNo(3) & "-" & strCNo(4) & ";"
                        strUpdCNo(1) = strCNo(1)
                        strUpdCNo(2) = strCNo(2)
                        strUpdCNo(3) = strCNo(3)
                        strUpdCNo(4) = strCNo(4)
                        strUpdMsg = strTmpMsg
                     End If
                  End If
                  RsTemp.MoveNext
               Loop
               
            End If
         End If 'Added by Morgan 2015/9/22
         
         If strCP07 <> "" Then
            'Modify by Morgan 2009/8/13 以公告日更新改呼叫共用函式
            PUB_UpdCP07byPA14 strUpdCNo, strCP07, strMsg, strUpdMsg
            'end 2009/8/13
         End If
         
         'Added by Morgan 2015/12/18
         '台灣案核准時,若多國案仍未發文,則請預設自核准日起算1個月為該多國案之本所期限(原來無所限或較晚時更新，若法限晚於核准日+7個月時一併清除)
         If strTwPA20 <> "" Then
            PUB_UpdCP06byTwPA20 strUpdCNo2, strTwPA20
         End If
         'end 2015/12/18
            
      End If
   
      'Added by Morgan 2015/9/22
      '相關案已公開/公告
      If bolPublished Then
         PUB_UpdateUSCase strUpdCNo, strCP07, strMsg, strUpdMsg
      End If
      'end 2015/9/22
      
      'Add by Morgan 2006/9/15 若國外案無繪圖人員時帶國內案繪圖人員
      'Modified by Morgan 2025/1/13 不論國內或國外案的案件性質條件都改用案件性質改用 NewCasePtyList，因原用 CaseMapIn 或 CaseMapOut 有201新案翻譯可能會與其他申請程序如101發明申請同時收文而導致抓到2筆資料而造成存檔錯誤 Ex:CFP-034894
      strExc(0) = "SELECT EP13,ST04,ST06 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF" & _
         " WHERE (cp01,cp02,cp03,cp04) in (" & strCMList & ")" & _
         " and cp57 is null and cp10 in (" & NewCasePtyList & ")" & _
         " AND EP02(+)=CP09 AND EP13 IS NOT NULL AND ST01(+)=EP13 ORDER BY 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strCP29 = RsTemp.Fields(0)
         strST04 = RsTemp.Fields("ST04")
         strST06 = RsTemp.Fields("ST06")
         
         'Added by Morgan 2015/3/6
         '若國內案不繪圖(99999)時再抓國外案有繪圖的
         If strCP29 = "99999" Then
            strExc(0) = "SELECT cp29,ST04,ST06 FROM CASEPROGRESS,STAFF" & _
               " where (cp01,cp02,cp03,cp04) in (" & strCRList & ")" & _
               " and cp57 is null and cp10 in (" & NewCasePtyList & ") and cp29<>'99999'" & _
               " AND ST01(+)=cp29 ORDER BY 1"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCP29 = RsTemp.Fields(0)
               strST04 = RsTemp.Fields("ST04")
               strST06 = RsTemp.Fields("ST06")
            End If
         End If
         'end 2015/3/6
         
         'Add by Morgan 2011/10/21 若繪圖人員離職改抓該所繪圖主管
         If strST04 = "2" Then
            'Modified by Morgan 2021/3/17 +在職
            'Modified by Morgan 2022/12/20 改新抓在職主管若也離職再抓北所主管並EMail通知重新分案
            'strExc(0) = "SELECT ST01 FROM STAFF WHERE ST06='" & strST06 & "' AND ST05='81' and st04='1'"
            strExc(0) = "SELECT ST01,decode(st06,'" & strST06 & "',0,st06) Srt FROM STAFF WHERE ST05='81' and st04='1' order by 2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCP29 = RsTemp.Fields(0)
            End If
            
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " select '" & strUserNum & "','" & strCP29 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'的國內案繪圖人員已離職，請重新分案！','如旨',cp14" & _
               " from caseprogress WHERE (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp27 is null and cp57 is null" & _
               " and  cp10 in (" & NewCasePtyList & ") AND CP29 IS NULL"
            cnnConnection.Execute strSql, intI
         End If
         'end 2011/10/21
         strSql = "UPDATE CASEPROGRESS SET CP29='" & strCP29 & "'" & _
            " WHERE (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp27 is null and cp57 is null" & _
            " and  cp10 in (" & NewCasePtyList & ") AND CP29 IS NULL"
         cnnConnection.Execute strSql, intI
      End If
      
      End If 'Added by Morgan 2021/2/25 限定CFP案(FCP也會用)
      
   End If

If m_Sys = "CFP" Then 'Added by Morgan 2021/2/25 限定CFP案(FCP也會用)

   'Add by Morgan 2016/5/5 規則同UpdateEp08,CFP主案會稿完成更新相同承辦的其他多國案的齊備日--柄佑
   If strCRList <> "" Then
      '檢查是否有多國案主案已發文或會稿完成
'Modified by Morgan 2021/2/5
'      strExc(0) = "select distinct cp14 from caseprogress,Engineerprogress where (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp57 is null and  cp10 in (" & CaseMapOut & ") and cp21 is null and ep02(+)=cp09 and nvl(cp27,0)+nvl(ep08,0)>0"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         strExc(1) = Replace("'" & RsTemp.GetString(, , , "','") & "'", ",''", "")
'         '未發文未齊備且相同承辦的國外案
'         strExc(0) = "select ep02,cp06 from caseprogress,Engineerprogress where (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp27 is null and cp57 is null and  cp10 in (" & CaseMapOut & ") and cp14 is not null and cp14 in (" & strExc(1) & ") and ep02(+)=cp09 and ep06 is null"
'         intI = 1
'         Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            adoRecordset.MoveFirst
'            Do While Not adoRecordset.EOF
'               '更新齊備日
'               strSql = "Update Engineerprogress set ep06=" & strSrvDate(1) & " where ep06 is null and ep02='" & adoRecordset.Fields(0).Value & "'"
'               cnnConnection.Execute strSql, intI
'               '承辦期限
'               If PUB_IfSetCP48() Then
'                  strExc(0) = "Select cp01,pa09,cp10 From CaseProgress, Patent Where CP09='" & adoRecordset.Fields(0).Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'                  intI = 1
'                  Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     strCP48 = Pub_GetHandleDay(AdoRecordSet3("cp01"), AdoRecordSet3("pa09"), AdoRecordSet3("cp10"), , "" & adoRecordset.Fields(1), adoRecordset.Fields(0))
'                     If strCP48 <> "" Then
'                        strSql = "Update caseprogress set cp48=" & strCP48 & " where cp09='" & adoRecordset.Fields(0) & "'"
'                        cnnConnection.Execute strSql, intI
'                     End If
'                  End If
'               End If
'               adoRecordset.MoveNext
'            Loop
'         End If
'      End If
      PUB_SetEP06ByCR strRelation(0, 0), strRelation(1, 0), strRelation(2, 0), strRelation(3, 0)
'end 2021/2/5
   End If
   'end 2016/5/5
   
   'Add by Morgan 2007/9/6
   If Me.MSFlexGrid1.Rows > 2 Then
   
      'Add by Morgan 2007/8/20 當有美國案時,若有主張優先權則以最早優先權日+18個月為預定公開日，否則為申請日+18個月更新其他多國案期限
      'Modify by Morgan 2008/7/16 加判斷非PCT案才做
      'Modified by Morgan 2021/2/22 +判斷發明案才做 Ex:CFP-031624(美),CFP-032236(日)
      If i101 <> -1 And str101PA46 = "" And str101PA08 = "1" Then
         strCaseNo(1) = strRelation(0, i101)
         strCaseNo(2) = strRelation(1, i101)
         strCaseNo(3) = strRelation(2, i101)
         strCaseNo(4) = strRelation(3, i101)
         strCP07 = ""
         '已有公開日
         If str101PA12 <> "" Then
            strCP07 = str101PA12
         Else
            strExc(0) = "SELECT PD05 FROM PRIDATE WHERE PD01='" & strCaseNo(1) & "' AND PD02='" & strCaseNo(2) & "' AND PD03='" & strCaseNo(3) & "' AND PD04='" & strCaseNo(4) & "' AND PD05>0 ORDER BY PD05"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCP07 = RsTemp.Fields(0)
            Else
               strExc(0) = "SELECT PA10 FROM PATENT WHERE PA01='" & strCaseNo(1) & "' AND PA02='" & strCaseNo(2) & "' AND PA03='" & strCaseNo(3) & "' AND PA04='" & strCaseNo(4) & "' AND PA10>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCP07 = RsTemp.Fields(0)
               End If
            End If
            strCP07 = CompDate(1, 18, strCP07)
         End If
         
         If strCP07 <> "" Then
            PUB_UpdCP07byPA12 strCaseNo, strCP07
         End If
      End If
      'end 2007/8/20
   
      strCaseNo(1) = strRelation(0, 0)
      strCaseNo(2) = strRelation(1, 0)
      strCaseNo(3) = strRelation(2, 0)
      strCaseNo(4) = strRelation(3, 0)
      PUB_UpdCP07byCP27 strCaseNo
   End If
   'end 2007/9/6
   
   'Added by Morgan 2012/3/26 主張新穎性優惠期要更新相關美國發明案提申期限
   If PUB_GetRefCaseList(m_CP(), m_list()) = True Then
      If UBound(m_list, 2) > 1 Then
         PUB_UpdateUsPatent m_list
      End If
   End If
   'end 2012/3/26
   
   'Added by Morgan 2023/3/17 未發文美國新案IDS檢查(相關案是否已有OA)
   If i101 <> -1 And str101PA08 = "1" Then
      PUB_NewUsCaseIdsChk strRelation(0, i101), strRelation(1, i101), strRelation(2, i101), strRelation(3, i101)
   End If
   'end 2023/3/17
   
   'Added by Morgan 2023/8/2 補建或修改關聯
   '承辦人是外翻人員,且已分案時,若P案已完稿,則請系統自動發MAIL通知承辦人(系統會透過ST14轉發給品薇)
   If strCMList <> "" Then
      strExc(0) = "SELECT cp01,cp02,cp03,cp04,cp14" & _
         " FROM caseprogress where (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp10 in (" & CaseMapOut & ") and cp14 like 'F5%' and cp158=0 and cp159=0 and cp157>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            CFPMail2F5xx .Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), .Fields("cp14")
            .MoveNext
         Loop
         End With
      End If
   End If
   'end 2023/8/2
   
End If 'Added by Morgan 2021/2/25
   
   cnnConnection.CommitTrans
   SaveDataNew = True
   If strMsg <> "" Then MsgBox strMsg 'Add by Morgan 2010/04/23
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function
'檢查國內案是否有與其他國外建關聯
Private Function CheckInnerCase() As Boolean
   Dim arrCaseNo, strOutterCase As String, strDropCaseList As String
   
   For intI = 1 To Me.MSFlexGrid1.Rows - 1
      strOutterCase = strOutterCase & "," & Left(MSFlexGrid1.TextMatrix(intI, 1), 15)
   Next
   With Adodc2.Recordset
      .MoveFirst
      Do While Not .EOF
         strDropCaseList = ""
         arrCaseNo = Split(.Fields(0), "-")
         strExc(0) = "select cm01||'-'||cm02||'-'||cm03||'-'||cm04 from casemap where cm01='CFP' and cm05='" & arrCaseNo(0) & "' and cm06='" & arrCaseNo(1) & "' and cm07='" & arrCaseNo(2) & "' and cm08='" & arrCaseNo(3) & "' and cm10='0'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
               .MoveFirst
               Do While Not .EOF
                  If InStr(strOutterCase, .Fields(0)) = 0 Then
                     strDropCaseList = strDropCaseList & vbCrLf & .Fields(0)
                  End If
                  .MoveNext
               Loop
            End With
         End If
         If strDropCaseList <> "" Then
            If MsgBox("國內案 " & .Fields(0) & " 尚與下列國外案" & vbCrLf & "有關連，是否確定要移除關連？" & vbCrLf & strDropCaseList, vbYesNo + vbDefaultButton2) = vbNo Then
               CheckInnerCase = False
               Exit Function
            End If
         End If
         .MoveNext
      Loop
   End With
   CheckInnerCase = True
End Function

Private Sub cmdok_Click(Index As Integer)

   Dim bolCancel As Boolean
   
   Select Case Index
      Case 0
         If Me.MSFlexGrid1.Rows < 2 Then
            MsgBox "請輸入國外案資料！"
            Exit Sub
         End If
         
         m_Sys = Left(MSFlexGrid1.TextMatrix(1, 1), 3) 'Added by Morgan 2021/2/25
            
         If Adodc2.Recordset.RecordCount = 0 Then
            If Me.MSFlexGrid1.Rows = 2 Then
               MsgBox "至少須有一筆國內案或兩筆國外案才可建關聯！"
               Exit Sub
            ElseIf MsgBox("是否確定沒有國內案號！", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Sub
            End If
         '檢查國內案號
         ElseIf CheckInnerCase = False Then
            Exit Sub
         End If
         
         '檢查多國主案設定
         If Me.MSFlexGrid1.Rows > 2 Then
            If MainCaseCheck = False Then
               Exit Sub
            End If
         End If
         
         'Added by Morgan 2016/6/6
         '多國案建國內關聯時必須先建主案，否則若國內已發文而子案又未正確設定則存檔時將會自動被上齊備日。(事後改設子案也不會還原，需請工程師主管手動清除)
         'Modified by Morgan 2021/2/25 限定CFP案(FCP也會用)
         If Me.MSFlexGrid1.Rows = 2 And m_Sys = "CFP" Then
            If MsgBox("CFP多國案與P案建關聯時，須先建主案！" & vbCrLf & vbCrLf & "是否確定要繼續？" & vbCrLf & vbCrLf & "(若國內案已發文多國主案會在建立關聯時自動齊備！)", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
               Exit Sub
            End If
         End If
         'end 2016/6/6
         
         If SaveDataNew Then
            If intWhereComeFrom = 1 Then
               intLeaveKind = 1
               bolLeave = True
               Unload Me
            Else
               MsgBox "存檔成功！"
               Grid1Clear
            End If
         End If
         
      Case 1
         intLeaveKind = 1
         bolLeave = False
         Unload Me
         
     'Add by Morgan 2006/9/25
     Case 2 '清除所有關聯
         If lblMainCase <> "" Then
            If MsgBox("是否要清除 " & lblMainCase & " 群組的所有關聯？", vbYesNo + vbDefaultButton2) = vbYes Then
               If DeleteAllRelation Then
                  If intWhereComeFrom = 1 Then
                     intLeaveKind = 1
                     bolLeave = True
                     Unload Me
                  Else
                     MsgBox "存檔成功！"
                     Grid1Clear
                  End If
               End If
               
            End If
         End If
   End Select
   
End Sub

Private Function DeleteAllRelation() As Boolean

   Dim strCRList As String
   
   strCRList = "select '" & m_CP(1) & "' CM1,'" & m_CP(2) & "' CM2,'" & m_CP(3) & "' CM3,'" & m_CP(4) & "' CM4 from dual union all select cr01,cr02,cr03,cr04 from caserelation where cr05='" & m_CP(1) & "' and cr06='" & m_CP(2) & "' and cr07='" & m_CP(3) & "' and cr08='" & m_CP(4) & "'"
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
   
   '若原為多國案且繪圖尚未確認分案則墨圖要計件
   strSql = "Update EngineerProgress SET EP29=NULL Where EP02 in (select cp09 from caseprogress where (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp10 in (" & CaseMapOut & ") and cp107 is null and cp21='Y')"
   cnnConnection.Execute strSql, intI
   
   '清除是否多國設定
   strSql = "Update caseprogress set cp21=NULL where (cp01,cp02,cp03,cp04) in (" & strCRList & ") and cp10 in (" & CaseMapOut & ") and cp107 is null and cp21='Y'"
   cnnConnection.Execute strSql, intI
   
   '清除國內外關聯
   strSql = "Delete from CaseMap where CM10='0' AND (CM01,CM02,CM03,CM04) IN (" & strCRList & ")"
   cnnConnection.Execute strSql, intI
   
   '清除多國關聯
   strSql = "delete from caserelation where (cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) IN (SELECT X.*,Y.* FROM (" & strCRList & ") X,(" & strCRList & ") Y WHERE NOT (X.CM1=Y.CM1 AND X.CM2=Y.CM2 AND X.CM3=Y.CM3 AND X.CM4=Y.CM4))"
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   DeleteAllRelation = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If

End Function

Private Sub DataGrid2_Click()
   Dim arrCaseNo
   If Adodc2.Recordset.RecordCount > 0 Then
      If Not (Adodc2.Recordset.EOF Or Adodc2.Recordset.BOF) Then
         arrCaseNo = Split(Adodc2.Recordset.Fields(0), "-")
         txtInCase(0) = arrCaseNo(0)
         txtInCase(1) = arrCaseNo(1)
         txtInCase(2) = arrCaseNo(2)
         txtInCase(3) = arrCaseNo(3)
      End If
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If Me.MSFlexGrid1.Rows - 1 <= 1 Then
      Exit Sub
   End If
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Morgan 2004/9/7
   If intWhereComeFrom = 1 Then
      Me.m_form.Show
   End If
   'Add By Cheng 2002/07/18
   Set frm1104 = Nothing
End Sub


Private Sub lblNation_Change()
   Dim strTemp As String
   
   If lblNation = "" Then
      lblNationName = ""
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNation(lblNation, strTemp) Then
      If ClsPDGetNation(lblNation, strTemp) Then
         lblNationName = strTemp
      End If
   End If
End Sub

Private Sub lblNation1_Change()
   Dim strTemp As String
   
   If lblNation1 = "" Then
      lblNationName1 = ""
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNation(lblNation1, strTemp) Then
      If ClsPDGetNation(lblNation1, strTemp) Then
         lblNationName1 = strTemp
      End If
   End If
End Sub

Private Sub MSFlexGrid1_Click()
   Dim ii As Integer
   
   With Me.MSFlexGrid1
      If .row < 1 Then Exit Sub
      If .TextMatrix(.row, 1) = "" Then Exit Sub
      For ii = 1 To .Rows - 1
         If ii <> .row Then .TextMatrix(ii, 0) = ""
      Next ii
      If .TextMatrix(.row, 0) = "" Then
         .TextMatrix(.row, 0) = "V"
      Else
         .TextMatrix(.row, 0) = ""
      End If
   End With
End Sub

Private Sub MSFlexGrid1_DblClick()

   With Me.MSFlexGrid1
      If .row < 1 Then Exit Sub
      If .col = 6 Then
         If .TextMatrix(.row, 6) = "" Then
            .TextMatrix(.row, 6) = "Y"
         Else
            .TextMatrix(.row, 6) = ""
         End If
      End If
   End With
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii = 13 Then
      cmdMove_Click 0
      txtSystem.SetFocus
   End If
End Sub

Private Sub txtInCase_Change(Index As Integer)
   cboCaseName1.Clear
   lblNation1 = ""
   lblCustomer1 = ""
End Sub

Private Sub txtInCase_GotFocus(Index As Integer)
   TextInverse txtInCase(Index)
End Sub

Private Sub txtInCase_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii = 13 Then
      cmdMove2_Click 4
   End If
End Sub

Private Function CheckKeyIn3(ByRef intIndex As Integer) As Boolean
   Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
   Dim strCustomer As String, strNation As String, i As Integer
   
   If intIndex > 0 And Len(txtInCase(intIndex)) > 0 And Len(txtInCase(intIndex)) < txtInCase(intIndex).MaxLength Then
      ShowMsg MsgText(10)
   ElseIf intIndex = 3 Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(txtInCase(0), txtInCase(1), _
         IIf(txtInCase(2) = "", "0", txtInCase(2)), IIf(txtInCase(3) = "", "00", txtInCase(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
      If ClsPDCheckCaseCodeIsExist(txtInCase(0), txtInCase(1), _
         IIf(txtInCase(2) = "", "0", txtInCase(2)), IIf(txtInCase(3) = "", "00", txtInCase(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
         SetNameToCombo cboCaseName1, strCaseName1, strCaseName2, strCaseName3
         lblNation1 = strNation
         lblCustomer1 = strCustomer
         CheckKeyIn3 = True
      End If
   Else
      CheckKeyIn3 = True
   End If
End Function

Private Sub txtInCase_Validate(Index As Integer, Cancel As Boolean)
   CheckKeyIn3 Index
End Sub

Private Sub txtSystem_Change()
   If cboCaseName.ListCount > 0 Then cboCaseName.Clear
   lblNation = ""
   lblCustomer = ""
End Sub

Private Sub txtSystem_GotFocus()
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem.Text)
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   If txtSystem <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
      If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
         ShowMsg MsgText(1056)
         Cancel = True
         txtSystem_GotFocus
      End If
   End If
End Sub

Private Sub txtCode_Change(Index As Integer)
   If cboCaseName.ListCount > 0 Then cboCaseName.Clear
   lblNation = ""
   lblCustomer = ""
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   CheckKeyIn2 (Index)
End Sub

Private Function CheckKeyIn2(ByRef intIndex As Integer) As Boolean
   Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
   Dim strCustomer As String, strNation As String, i As Integer
   
   If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
      ShowMsg MsgText(10)
   ElseIf intIndex = 2 Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
         IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
      If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
         IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
         SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
         lblNation = strNation
         lblCustomer = strCustomer
         CheckKeyIn2 = True
      End If
   Else
      CheckKeyIn2 = True
   End If
End Function

'Add By Cheng 2003/09/17
Private Sub InitialGrid()
With Me.MSFlexGrid1
    .Cols = 10
    'Modified by Morgan 2022/1/11 改用MSHFlexGrid後要設2，否則固定列會變白色
    '.Rows = 1
    .Clear
    .Rows = 2
    'end 2022/1/11
    .row = 0
    .col = 0: .Text = "V"
    .ColWidth(0) = 300: .ColAlignment(0) = flexAlignCenterCenter
    .col = 1: .Text = "本所案號"
    .ColWidth(1) = 1625: .ColAlignment(1) = flexAlignLeftCenter
    .col = 2: .Text = "專利種類"
    .ColWidth(2) = 800: .ColAlignment(2) = flexAlignLeftCenter
    .col = 3: .Text = "案件名稱"
    .ColWidth(3) = 2325: .ColAlignment(3) = flexAlignLeftCenter
    .col = 4: .Text = "申請國家"
    .ColWidth(4) = 950: .ColAlignment(4) = flexAlignLeftCenter
    .col = 5: .Text = "申請人"
    .ColWidth(5) = 2235: .ColAlignment(5) = flexAlignLeftCenter
    .col = 6: .Text = "是否多國"
    .ColWidth(6) = 800: .ColAlignment(6) = flexAlignLeftCenter
    .col = 7: .Text = "" '原是否多國設定
    .ColWidth(7) = 0
    .col = 8: .Text = "" '申請國家
    .ColWidth(8) = 0
    .col = 9: .Text = "" '多國主案順位
    .ColWidth(9) = 0
End With
End Sub
'Add by Morgan 2006/7/3
Private Function ChkNoOtherGroup(p_stAddCase As String) As Boolean
   
   Dim strTmp As Variant, cr(5 To 8) As String, j As Integer
   Dim strGroupCase() As String
   Dim bolQuest As Boolean
   
   strTmp = Split(p_stAddCase, "-")
   For j = 0 To 3
       cr(5 + j) = strTmp(j)
   Next
   '修改
   If m_CP(1) <> "" Then
      If m_CP(1) = cr(5) And m_CP(2) = cr(6) And m_CP(3) = cr(7) And m_CP(4) = cr(8) Then
         ChkNoOtherGroup = True
         Exit Function
      Else
         strExc(0) = "select * from caserelation where cr01='" & m_CP(1) & "' and cr02='" & m_CP(2) & "' and cr03='" & m_CP(3) & "' and cr04='" & m_CP(4) & "'" & _
            " and cr05='" & cr(5) & "' and cr06='" & cr(6) & "' and cr07='" & cr(7) & "' and cr08='" & cr(8) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '原本群組案件
            ChkNoOtherGroup = True
            Exit Function
         End If
      End If
   End If
      
   
      strExc(0) = "select * from caserelation where cr05='" & cr(5) & "' and cr06='" & cr(6) & "' and cr07='" & cr(7) & "' and cr08='" & cr(8) & "' ORDER BY 1,2,3,4"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         '非本群組案件且非其他群組案件
         ChkNoOtherGroup = True
      Else
         If m_bolCombine = True Then
            MsgBox "此案號已存在於另一相關群組，因本群組已有合併群組，故不可再行合併！"
         Else
            
            If MSFlexGrid1.Rows > 1 Then
               If MsgBox("此案號已存在於另一相關群組，是否要合併？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                  m_bolCombine = True
               Else
                  Exit Function
               End If
            End If
            
            Erase m_CombGroupCase
            ReDim m_CombGroupCase(3, RsTemp.RecordCount)
            RsTemp.MoveFirst
            m_CombGroupCase(0, 0) = RsTemp.Fields("cr05")
            m_CombGroupCase(1, 0) = RsTemp.Fields("cr06")
            m_CombGroupCase(2, 0) = RsTemp.Fields("cr07")
            m_CombGroupCase(3, 0) = RsTemp.Fields("cr08")
            For j = 1 To RsTemp.RecordCount
               m_CombGroupCase(0, j) = RsTemp.Fields("cr01")
               m_CombGroupCase(1, j) = RsTemp.Fields("cr02")
               m_CombGroupCase(2, j) = RsTemp.Fields("cr03")
               m_CombGroupCase(3, j) = RsTemp.Fields("cr04")
               RsTemp.MoveNext
            Next
            SetRelationToLisBox m_CombGroupCase, True
            If m_bolCombine = False Then
               m_CP(1) = m_CombGroupCase(0, 0)
               m_CP(2) = m_CombGroupCase(1, 0)
               m_CP(3) = m_CombGroupCase(2, 0)
               m_CP(4) = m_CombGroupCase(3, 0)
               lblMainCase = m_CP(1) & "-" & m_CP(2) & "-" & m_CP(3) & "-" & m_CP(4)
               SetInCase m_CP
            End If
            
         End If
      End If
   
End Function

Private Sub CheckAllGroupCase()
   Dim i As Integer, j As Integer
   For i = 1 To Me.MSFlexGrid1.Rows - 1
      If MSFlexGrid1.TextMatrix(i, 0) = "" Then
         For j = 0 To UBound(m_CombGroupCase, 2)
            If m_CombGroupCase(0, j) & m_CombGroupCase(1, j) & m_CombGroupCase(2, j) & m_CombGroupCase(3, j) = Replace(Left(MSFlexGrid1.TextMatrix(i, 1), 15), "-", "") Then
               MSFlexGrid1.TextMatrix(i, 0) = "V"
            End If
         Next
      End If
   Next
End Sub
'Add by Morgan 2006/7/3
Private Function SetRelationToLisBox(ByRef strRelation() As String, Optional ByVal p_bolAdd As Boolean = False) As Boolean

   Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
   Dim strCustomer As String, strNation As String, i As Integer, strCaseName As String
   Dim strCaseCode As String, strTemp As String
   Dim iPos As Integer, stPA57 As String
   
   If p_bolAdd = False Then
      Me.MSFlexGrid1.Rows = 1
   End If
   For i = 0 To UBound(strRelation, 2)
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(strRelation(0, i), strRelation(1, i), strRelation(2, i), strRelation(3, i), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
      If ClsPDCheckCaseCodeIsExist(strRelation(0, i), strRelation(1, i), strRelation(2, i), strRelation(3, i), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
         If strCaseName1 <> "" Then
            strCaseName = strCaseName1
         ElseIf strCaseName2 <> "" Then
            strCaseName = strCaseName2
         Else
            strCaseName = strCaseName3
         End If
         strCaseCode = strRelation(0, i) + "-" + strRelation(1, i) + IIf(strRelation(2, i) = "0", "-0", "-" + strRelation(2, i)) + IIf(strRelation(3, i) = "00", "-00", "-" + strRelation(3, i))
         
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(strNation, strTemp) Then
         If ClsPDGetNation(strNation, strTemp) Then
            'Modified by Morgan 2022/1/11
            'Me.MSFlexGrid1.AddItem Me.MSFlexGrid1.Rows
            If MSFlexGrid1.TextMatrix(1, 1) <> "" Then
               Me.MSFlexGrid1.AddItem Me.MSFlexGrid1.Rows
            End If
            'end 2022/1/11
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 0) = ""
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 2) = GetPA08(strCaseCode, stPA57)
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 1) = strCaseCode & IIf(stPA57 = "Y", "＊", "")
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 3) = strCaseName
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 4) = strTemp
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 5) = strCustomer
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 6) = GetCP21(strCaseCode)
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 7) = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 6)
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 8) = strNation
            iPos = InStr(MultiCountryPriority, strNation)
            If iPos = 0 Then iPos = 99
            '多國主案順位
            Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 9) = iPos
            
            strExc(1) = strRelation(0, i)
            strExc(2) = strRelation(1, i)
            strExc(3) = strRelation(2, i)
            strExc(4) = strRelation(3, i)
            SetInCase strExc
            SetRelationToLisBox = True
         Else
            SetRelationToLisBox = False
            Exit For
         End If
      Else
         SetRelationToLisBox = False
         Exit For
      End If
   Next
   MSFlexGrid1.row = 1 'Added by Morgan 2022/1/11 要指定列否則第1筆會反白
End Function
'Add by Morgan 2006/7/3
'讀取相關卷號檔
Private Function oReadCaseRelationData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String, ByRef strRelation() As String) As Integer
   Dim i As Integer, j As Integer
On Error GoTo ErrHand
   oReadCaseRelationData = 0
   strSql = "select cr05,cr06,cr07,cr08 from caserelation where cr01=" + CNULL(strCode1) + " and cr02=" + CNULL(strCode2) + " and cr03=" + CNULL(strCode3) + " and cr04=" + CNULL(strCode4) & "  ORDER BY 1,2,3,4"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         ReDim Preserve strRelation(3, j)
         For i = 0 To 3
            strRelation(i, j) = .Fields(i)
         Next
         .MoveNext
         j = j + 1
      Loop
      End With
      oReadCaseRelationData = 1
   End If
ErrHand:
   If Err.NUMBER <> 0 Then
      oReadCaseRelationData = -1
      MsgBox Err.Description
   End If
End Function
'新增國內案
'p_Mode:1=程式,2=按鈕
Private Sub Grid2Add(Optional ByVal p_Mode As Integer = 1)
   Dim bolAdd As Boolean, strCRList As String
   strExc(1) = txtInCase(0)
   strExc(2) = txtInCase(1)
   strExc(3) = Right("0" & txtInCase(2), 1)
   strExc(4) = Right("00" & txtInCase(3), 2)
   
   strSql = "select pa01||'-'||pa02||'-'||pa03||'-'||pa04 C00,ptm03 C01,nvl(pa05,pa06) C02,na03 C03,cu04 C04" & _
      " from patent, patenttrademarkmap , nation, customer" & _
      " where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "'" & _
      " and ptm01(+)='1' and ptm02(+)=pa08" & _
      " and na01(+)=pa09 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)"
   intI = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      With AdoRecordSet3
      bolAdd = False
      If Adodc2.Recordset.RecordCount > 0 Then
         Adodc2.Recordset.MoveFirst
         Adodc2.Recordset.Find "C00='" & .Fields("C00") & "'"
         If Adodc2.Recordset.EOF Then
            bolAdd = True
         End If
      Else
         bolAdd = True
      End If
      If bolAdd = True Then
         If p_Mode = 2 Then
            strCRList = GetCRList
            strSql = "SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE CM10='0' AND CM01='CFP' AND CM06='" & strExc(2) & "' AND CM07='" & strExc(3) & "' AND CM08='" & strExc(4) & "' AND CM05='" & strExc(1) & "'"
            If strCRList <> "" Then
               strSql = strSql & " AND NOT (CM01,CM02,CM03,CM04) IN (" & strCRList & ")"
            End If
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
            If intI = 1 Then
               If MsgBox("本國內案尚有其他國外案是否加入多國？", vbYesNo + vbDefaultButton2) = vbYes Then
                  txtSystem = "" & adoRecordset.Fields(0)
                  txtCode(0) = "" & adoRecordset.Fields(1)
                  txtCode(1) = "" & adoRecordset.Fields(2)
                  txtCode(2) = "" & adoRecordset.Fields(3)
                  cmdMove_Click 0
               End If
               Exit Sub
            End If
         End If
         Adodc2.Recordset.AddNew
         'Add by Amy 2014/06/06
         Adodc2.Recordset.Fields("FormName") = Me.Name
         Adodc2.Recordset.Fields("ID") = strUserNum
         Adodc2.Recordset.Fields("SeqNo") = strSeqNo
         Adodc2.Recordset.Fields("C00") = .Fields("C00")
         Adodc2.Recordset.Fields("C01") = .Fields("C01")
         Adodc2.Recordset.Fields("C02") = .Fields("C02")
         Adodc2.Recordset.Fields("C03") = .Fields("C03")
         Adodc2.Recordset.Fields("C04") = .Fields("C04")
         Adodc2.Recordset.UpdateBatch
         txtInCase(1) = ""
         txtInCase(2) = ""
         txtInCase(3) = ""
      End If
      End With
   Else
      MsgBox "查無此國內案號資料!!!", vbExclamation + vbOKOnly
   End If
End Sub
Private Sub SetInCase(p_CM() As String)
   strSql = "select cm05,cm06,cm07,cm08" & _
      " from casemap" & _
      " where cm10='0' and cm01='" & p_CM(1) & "' and cm02='" & p_CM(2) & "' and cm03='" & p_CM(3) & "' and cm04='" & p_CM(4) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      Do While Not RsTemp.EOF
         txtInCase(0) = "" & RsTemp.Fields(0)
         txtInCase(1) = "" & RsTemp.Fields(1)
         txtInCase(2) = "" & RsTemp.Fields(2)
         txtInCase(3) = "" & RsTemp.Fields(3)
         Grid2Add
         RsTemp.MoveNext
      Loop
   'Add By Sindy 2022/11/23
   ElseIf m_CRL01 <> MsgText(601) Then
      strExc(10) = Pub_GetCRLCaseMap(m_CRL01, "0", "P", p_CM(1), p_CM(2), p_CM(3), p_CM(4))
      If strExc(10) <> "" Then
         txtInCase(0) = SystemNumber(strExc(10), 1)
         txtInCase(1) = SystemNumber(strExc(10), 2)
         txtInCase(2) = SystemNumber(strExc(10), 3)
         txtInCase(3) = SystemNumber(strExc(10), 4)
      End If
      '2022/11/23 END
   End If
End Sub

Private Function GetCP21(ByVal p_CaseNo As String) As String
   p_CaseNo = Replace(p_CaseNo, "-", "")
   strSql = "select cp21 from caseprogress where " & ChgCaseprogress(p_CaseNo) & " and cp10 in (" & CaseMapOut & ")"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      GetCP21 = "" & RsTemp.Fields(0)
   End If
End Function
'Modify by Morgan 2008/1/9 加抓是否閉卷
Private Function GetPA08(ByVal p_CaseNo As String, Optional p_PA57 As String) As String
   p_CaseNo = Replace(p_CaseNo, "-", "")
   strSql = "select decode(pa09,'020',PTM04,PTM03),PA57 from patent,PatentTradeMarkMap  where " & ChgPatent(p_CaseNo) & " and ptm02(+)=pa08 and ptm01='1'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      GetPA08 = "" & RsTemp.Fields(0)
      p_PA57 = "" & RsTemp.Fields(1)
   End If
End Function

Private Sub InitGrid()
   'Modify by Amy 2014/06/06 +FormName 改暫存TB
   'Set Adodc2.Recordset = PUB_CreateRecordset(, 5)
   Set Adodc2.Recordset = PUB_CreateRecordset(, 5, , , Me.Name, strSeqNo)
   Set DataGrid2.DataSource = Adodc2
   txtInCase(1) = ""
   txtInCase(2) = ""
   txtInCase(3) = ""
End Sub

Private Function MainCaseCheck() As Boolean

   Dim i As Integer, j As Integer
   j = 0
   For i = 1 To Me.MSFlexGrid1.Rows - 1
      If Me.MSFlexGrid1.TextMatrix(i, 6) = "" Then
         j = j + 1
      End If
   Next
   If j = 0 Then
      If MsgBox("是否確定不設主案！", vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   ElseIf j > 1 Then
      If MsgBox("是否確認要設定 " & j & " 個主案？", vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   MainCaseCheck = True
   
End Function
