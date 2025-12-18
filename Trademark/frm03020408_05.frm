VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020408_05 
   BorderStyle     =   1  '單線固定
   Caption         =   "更改來函期限"
   ClientHeight    =   5736
   ClientLeft      =   216
   ClientTop       =   936
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9144
   Begin VB.TextBox textCP08 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2210
      Width           =   3255
   End
   Begin VB.TextBox textCP48 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2550
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2220
      Width           =   2532
   End
   Begin VB.TextBox textCP07 
      Height          =   285
      Left            =   5670
      MaxLength       =   7
      TabIndex        =   9
      Top             =   3690
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1548
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   570
      Width           =   2532
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3870
      TabIndex        =   37
      Top             =   3120
      Width           =   4215
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   7
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   3
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   5
         Top             =   150
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到          天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        月"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      日"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   6
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1200
      TabIndex        =   36
      Top             =   3120
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   144
         TabIndex        =   0
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   4620
      TabIndex        =   12
      Top             =   30
      Width           =   1212
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1558
      Width           =   2532
   End
   Begin VB.TextBox textCP06 
      Height          =   285
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   8
      Top             =   3690
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   570
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1889
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   15
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   13
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   14
      Top             =   30
      Width           =   1212
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   972
      Left            =   1200
      TabIndex        =   51
      Top             =   4032
      Width           =   7752
      _ExtentX        =   13674
      _ExtentY        =   1715
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
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   1200
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2541
      Width           =   2532
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   648
      Left            =   1200
      TabIndex        =   10
      Top             =   5040
      Width           =   7728
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13631;1143"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5670
      TabIndex        =   49
      Top             =   1879
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   48
      Top             =   886
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1217
      Width           =   7485
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13203;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblDate2 
      Caption         =   "lblDate2"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5640
      TabIndex        =   46
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblDate1 
      Caption         =   "lblDate1"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   45
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "修改前法定期限:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   44
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "修改前本所期限:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label32 
      Caption         =   "來函期限:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3780
      TabIndex        =   35
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   255
      Index           =   5
      Left            =   4740
      TabIndex        =   34
      Top             =   570
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   255
      Index           =   7
      Left            =   4740
      TabIndex        =   33
      Top             =   1563
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   1563
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   31
      Top             =   3705
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "本所期限 :"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3705
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   570
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   901
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1232
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   2225
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   1894
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4740
      TabIndex        =   24
      Top             =   1894
      Width           =   885
   End
   Begin VB.Label Label7 
      Caption         =   "本案期限 :"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4032
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   22
      Top             =   2225
      Width           =   885
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2556
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   20
      Top             =   2565
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
End
Attribute VB_Name = "frm03020408_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/20 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/09/13 改成Form2.0 ; cmbTM05、textTM23、textCP13、textCP14_2、textCP64、grdList改字型=新細明體-ExtB
'Create by Lydia 2015/05/05 更改來函期限
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 申請國家
Dim m_TM10 As String
' 來函收文日
Dim m_CP05 As String
' 機關文號
Dim m_CP08 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 業務區
Dim m_CP12 As String
' 智權人員
Dim m_CP13 As String
'Added by Lydia 2018/02/06 承辦人員
Dim m_CP14 As String
' 暫時存放 CF15
Dim m_CF15 As String
'
Dim m_CurrSel As Integer

Private Sub cmdCancel_Click()
   Unload Me
   frm03020408_03.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03020408_03
   Unload frm03020408_02
   Unload frm03020408_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass

      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm03020408_03
      Unload frm03020408_02
      Unload Me
      frm03020408_01.Show
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP08.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP48.BackColor = &H8000000F
  
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "下一程序/進度"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 1000
   grdList.col = 4
   grdList.Text = "機關文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "相關人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "備註"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "收文號"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "下一程序代號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "序號"
   grdList.ColWidth(9) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020408_04 = Nothing
End Sub



Private Sub grdList_Click()
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
         End If
      End If
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
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

' 取得商標基本檔
Private Sub QueryTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔資料
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         m_CP05 = rsTmp.Fields("CP05")
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 原-所限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         lblDate1.Caption = TAIWANDATE(rsTmp.Fields("CP06"))
      Else
         lblDate1.Caption = ""
      End If
      ' 原-法限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         lblDate2.Caption = TAIWANDATE(rsTmp.Fields("CP07"))
      Else
         lblDate2.Caption = ""
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If

      Call ChgType '帶入預設期限
      
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = "" & rsTmp.Fields("CP14") 'Added by Lydia 2018/02/06
         textCP14_2.Text = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         textCP08.Text = Trim(rsTmp.Fields("CP08"))
      End If
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64.Text = Trim(rsTmp.Fields("CP64"))
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   Dim strDay As String
   m_TM10 = Empty
   m_CP13 = Empty
      
   ' 本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04

   ' 讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 以來函性質來計算承辦期限
   strDay = Empty
      
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' and NP01 = '" & m_CP09 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 是否續辦欄位必須為空白
'         If IsNull(rsTmp.Fields("NP06")) = False Then
'            If IsEmptyText(rsTmp.Fields("NP06")) = False Then
'               GoTo NextRecord
'            End If
'         End If
         
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(grdList.row, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("NP15")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("NP22")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/20
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/20
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Sub


Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim nIndex As Integer
   Dim strSql As String, StrSQLa As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strNP06 As String
   Dim strNCP09 As String
'---------------
   
On Error GoTo CheckingErr
   
   '抓下一程序的處理狀況
   strExc(0) = " select * from nextprogress where np01='" & m_CP09 & "' and np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strNP06 = "" & RsTemp.Fields("NP06")
   End If
      '若已收文,抓相關收文號
   If strNP06 = "Y" Then
      strExc(0) = " select cp09 from caseprogress where cp43='" & m_CP09 & "' and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' "
         
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strNCP09 = "" & RsTemp.Fields("CP09")
      End If
   End If
   
cnnConnection.BeginTrans
 
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   'Modified by Lydia 2015/06/22 新增案件性質-更改來函期限
   'strCP10 = "1706"
   strCP10 = "1726"
   
    'cp13智權人員存最近收文A類接洽記錄單的智權人員
  'Modified by Lydia 2015/06/22 更正期限改為-更改來函期限
  'Modified by Lydia 2018/02/06 一般來函輸入-更改來函期限(1726)之承辦人及智權人員與原來函相同;
  ' strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP27,CP06,CP07,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & strSrvDate(1) & "," & CNULL(DBDATE(textCP06), True) & "," & CNULL(DBDATE(textCP07), True) & _
                    ",'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    CNULL(strUserNum) & ",'N','N','N','" & m_CP09 & "','更正來函期限') "
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP27,CP06,CP07,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & strSrvDate(1) & "," & CNULL(DBDATE(textCP06), True) & "," & CNULL(DBDATE(textCP07), True) & _
                    ",'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "'," & CNULL(m_CP12) & "," & CNULL(m_CP13) & " ," & _
                    CNULL(m_CP14) & ",'N','N','N','" & m_CP09 & "','更正來函期限') "
   'end 2018/02/06
   cnnConnection.Execute strSql
   '官方期限月數
   If Trim(Text11) <> "" Then
     strSql = "UPDATE CaseProgress SET CP134=" & Text11 & " WHERE CP09='" & strCP09 & "' "
     cnnConnection.Execute strSql
     StrSQLa = ",CP134=" & Text11
   End If

   '更新被修改的進度之所限、法限和進度備註
   strSql = " update CaseProgress set CP06=" & CNULL(DBDATE(textCP06), True) & ",CP07=" & CNULL(DBDATE(textCP07), True) & _
            ",cp64='" & ChgSQL(textCP64) & "' " & StrSQLa & _
            " where cp09=" & CNULL(m_CP09)
   cnnConnection.Execute strSql
   '更新下一程序的所限、法限和進度備註
   strSql = " update NextProgress set NP08=" & CNULL(DBDATE(textCP06), True) & ",NP09=" & CNULL(DBDATE(textCP07), True) & ",NP15='" & ChgSQL(textCP64) & "'" & _
            " where  np01='" & m_CP09 & "' and np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "'"
   cnnConnection.Execute strSql
   '更新下一程序已收文的進度檔的所限和法限
   If strNP06 = "Y" And Len(strNCP09) > 0 Then
       strSql = " update CaseProgress set CP06=" & CNULL(DBDATE(textCP06), True) & ",CP07=" & CNULL(DBDATE(textCP07), True) & _
                 StrSQLa & " where cp09=" & CNULL(strNCP09)
       cnnConnection.Execute strSql
   End If
  
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     'edit by nick 2004/11/03
     OnSaveData = False
End Function

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "進度備註資料內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查日期的格式
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/09
      End If

   End If
EXITSUB:
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      ' 檢查日期的格式
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim Cancel As Boolean
   
   CheckDataValid = False
   If Me.textCP06.Text <> "" Then
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If

    If IsEmptyText(textCP06) = True Then
       strTit = "資料檢核"
       strMsg = "本所期限不可為空白"
       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       textCP06.SetFocus
       GoTo EXITSUB
    End If

    If IsEmptyText(textCP07) = True Then
       strTit = "資料檢核"
       strMsg = "法定期限不可為空白"
       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       textCP07.SetFocus
       GoTo EXITSUB
    End If
    ' 本所期限不可大於法定期限
    If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
       If Val(textCP06) > Val(textCP07) Then
          strTit = "資料檢核"
          strMsg = "本所期限的日期不可超過法定期限的日期"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textCP06.SetFocus
          GoTo EXITSUB
       End If
    End If
    
    If Me.Option4(0).Value = True Then
       If Text10.Text = "" Then
          MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
          Me.Text10.SetFocus
          GoTo EXITSUB
       End If
    ElseIf Me.Option4(1).Value = True Then
       If Text11.Text = "" Then
          MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
          Me.Text11.SetFocus
          GoTo EXITSUB
       End If
    ElseIf Me.Option4(2).Value = True Then
       If Text12.Text = "" Then
          MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
          Me.Text12.SetFocus
          GoTo EXITSUB
       End If
    End If
    '以防修改期限天數或月數,重新計算期限
    If Me.Text10.Enabled = True Then
       Cancel = False
       Text10_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    If Me.Text11.Enabled = True Then
       Cancel = False
       Text11_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    'Added by Lydia 2021/09/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         GoTo EXITSUB
    End If
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub Option1_Click(Index As Integer)
   If Me.Option4(0).Value Then
      Text10_Validate False
   ElseIf Me.Option4(1).Value Then
      Text11_Validate False
   ElseIf Me.Option4(2).Value Then
      Text12_Validate False
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_LostFocus()
   '非台灣"天"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   CloseIme
End Sub

Private Sub Text11_LostFocus()
   '非台灣"月"跳離時到"本所期限"欄位
   'If m_TM10 <> 台灣國家代號 Then
   '   If textCP06.Enabled = True Then textCP06.SetFocus
   'End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   '非台灣"日"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         If m_TM10 = 台灣國家代號 Then
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               textCP07 = Text12
               If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                  textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
               Else
                  textCP06 = TransDate(CompDate(2, -2, TransDate(textCP07, 2)), 1)
               End If
               textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            End If
         End If
        If grdList.row >= 1 And grdList.col >= 9 Then
           If textCP06.Text <> "" And textCP06.Text <> textCP06.Tag Then
              Call ChangeGridShow(textCP06)
           End If
           If textCP07.Text <> "" And textCP07.Text <> textCP07.Tag Then
              Call ChangeGridShow(textCP07)
           End If
        End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '期限起算日
   
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   
  '  strFromDate = DBDATE(frm03020408_01.textCP05)
     strFromDate = DBDATE(m_CP05)
   If m_TM10 = 台灣國家代號 Then
      '文到天數
      If Option4(0).Value = True Then
         textCP07 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '文到月數
      ElseIf Option4(1).Value = True Then
         textCP07 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      If textCP07 <> "" Then
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
         Else
            textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
         End If
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
   If grdList.row >= 1 And grdList.col >= 9 Then
      If textCP06.Text <> "" And textCP06.Text <> textCP06.Tag Then
         Call ChangeGridShow(textCP06)
      End If
      If textCP07.Text <> "" And textCP07.Text <> textCP07.Tag Then
         Call ChangeGridShow(textCP07)
      End If
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '期限起算日
   
  
  ' strFromDate = DBDATE(frm03020408_01.textCP05)
   strFromDate = DBDATE(m_CP05)
   
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If

   If ClsPDGetCaseProperty(m_TM01, m_CP10, strTempName, bolTmp) Then
      textCP06 = ""
      textCP07 = ""
      
      If m_TM10 = 台灣國家代號 Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & m_CP10 & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(1)) Then
                     '文到天數
                     Option4(0).Value = True
                     Text10 = .Fields(1)
                     textCP07 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                  ElseIf Not IsNull(.Fields(2)) Then
                     '文到月數
                     Option4(1).Value = True
                     Text11 = .Fields(2)
                     textCP07 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                  Else
                     '文到天數
                     Option4(0).Value = True
                     Text10 = ""
                     Text11 = ""
                  End If
                  If textCP07 <> "" And Not IsNull(.Fields(0)) Then
                     '文到當日
                     If .Fields(0) = "1" Then
                        Option1(0).Value = True
                        textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
                     '文到次日
                     Else
                        Option1(1).Value = True
                     End If
                  End If
                  '文到天數
                  If Text10 <> "" Then
                     If Val(Text10) >= 60 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  '文到月數
                  ElseIf Not IsNull(.Fields(2)) Then
                     If Val(.Fields(2)) >= 2 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  End If
                  If textCP07 <> "" Then
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                        textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
                     Else
                        textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
                     End If
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function

Private Sub ChangeGridShow(pList As TextBox)
     If pList.Name = "textCP06" Then
        ' 本所期限
        If grdList.Rows > 0 And Len(grdList.TextMatrix(1, 1)) > 0 Then
           For intI = 1 To grdList.Rows - 1
             grdList.TextMatrix(intI, 2) = pList.Text
           Next intI
        End If
     Else
        ' 法定期限
        If grdList.Rows > 0 And Len(grdList.TextMatrix(1, 1)) > 0 Then
           For intI = 1 To grdList.Rows - 1
             grdList.TextMatrix(intI, 3) = pList.Text
           Next intI
        End If
     End If
   pList.Tag = pList.Text
End Sub
