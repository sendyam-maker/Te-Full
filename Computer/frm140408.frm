VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140408 
   BorderStyle     =   1  '單線固定
   Caption         =   "往來記錄資料查詢"
   ClientHeight    =   6012
   ClientLeft      =   3780
   ClientTop       =   3696
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6012
   ScaleWidth      =   9312
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   1095
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   420
      Width           =   300
   End
   Begin VB.Frame Frame3 
      Height          =   2250
      Left            =   1080
      TabIndex        =   19
      Top             =   780
      Width           =   5430
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   2565
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1845
         Width           =   852
      End
      Begin VB.OptionButton Option2 
         Caption         =   "編號："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   210
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1530
         Width           =   852
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   7
         Top             =   1845
         Width           =   852
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1380
         MaxLength       =   9
         TabIndex        =   3
         Top             =   120
         Width           =   1965
      End
      Begin VB.OptionButton Option2 
         Caption         =   "名稱："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   945
         Width           =   1000
      End
      Begin VB.Frame Frame1 
         Height          =   350
         Left            =   90
         TabIndex        =   24
         Top             =   480
         Width           =   2436
         Begin VB.OptionButton Option1 
            Caption         =   "中文"
            Height          =   180
            Index           =   0
            Left            =   72
            TabIndex        =   27
            Top             =   135
            Value           =   -1  'True
            Width           =   732
         End
         Begin VB.OptionButton Option1 
            Caption         =   "英文"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   26
            Top             =   135
            Width           =   732
         End
         Begin VB.OptionButton Option1 
            Caption         =   "日文"
            Height          =   180
            Index           =   2
            Left            =   1656
            TabIndex        =   25
            Top             =   135
            Width           =   732
         End
      End
      Begin VB.Frame Frame2 
         Height          =   350
         Left            =   2580
         TabIndex        =   21
         Top             =   480
         Width           =   2595
         Begin VB.OptionButton Option3 
            Caption         =   "字首比對"
            Height          =   180
            Index           =   0
            Left            =   72
            TabIndex        =   23
            Top             =   144
            Width           =   1125
         End
         Begin VB.OptionButton Option3 
            Caption         =   "模糊比對"
            Height          =   180
            Index           =   1
            Left            =   1260
            TabIndex        =   22
            Top             =   144
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "E-Mail："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1275
         Width           =   1000
      End
      Begin VB.TextBox Text9 
         Height          =   300
         Left            =   1380
         TabIndex        =   5
         Top             =   1215
         Width           =   1935
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   1553
         Width           =   1155
         Caption         =   "lblFM2"
         Size            =   "2037;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   330
         Left            =   1380
         TabIndex        =   4
         Top             =   870
         Width           =   3825
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         Size            =   "6747;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "開發日期："
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   1905
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   1590
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2310
         X2              =   2430
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         Height          =   180
         Left            =   3420
         TabIndex        =   31
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label10 
         Height          =   270
         Left            =   2280
         TabIndex        =   30
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(民國)"
         Height          =   180
         Left            =   3600
         TabIndex        =   29
         Top             =   1905
         Width           =   480
      End
   End
   Begin VB.TextBox Text10 
      Height          =   264
      Left            =   120
      MaxLength       =   200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton cmdAddSort 
      Caption         =   "新增↑"
      Height          =   285
      Left            =   7020
      TabIndex        =   10
      Top             =   3390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdRemSort 
      Caption         =   "移除↓"
      Height          =   285
      Left            =   7020
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3090
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "frm140408.frx":0000
      Left            =   1125
      List            =   "frm140408.frx":0002
      TabIndex        =   9
      Text            =   "cboSort"
      Top             =   3120
      Width           =   5550
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "往來記錄(&N)"
      Height          =   345
      Index           =   0
      Left            =   6630
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   15
      Width           =   1450
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2505
      Left            =   30
      TabIndex        =   16
      Top             =   3465
      Width           =   9285
      _ExtentX        =   16383
      _ExtentY        =   4424
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   8
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1095
      MaxLength       =   7
      TabIndex        =   0
      Top             =   90
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   2295
      MaxLength       =   7
      TabIndex        =   1
      Top             =   90
      Width           =   852
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5640
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   15
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   8085
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   15
      Width           =   1200
   End
   Begin VB.ListBox lstSort 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      ItemData        =   "frm140408.frx":0004
      Left            =   1440
      List            =   "frm140408.frx":000B
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   5550
   End
   Begin VB.Label Label9 
      Caption         =   "筆數："
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7920
      TabIndex        =   37
      Top             =   3240
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "往來對象：          ( 1.申請人 2.代理人 3.潛在客戶)"
      Height          =   180
      Left            =   90
      TabIndex        =   33
      Top             =   480
      Width           =   3810
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2070
      X2              =   2190
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "往來類別："
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "往來日期：                                                    (民國)"
      Height          =   180
      Left            =   90
      TabIndex        =   15
      Top             =   120
      Width           =   3720
   End
End
Attribute VB_Name = "frm140408"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、Label1、Text2
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'2008/11/27 ADD BY TONI
Option Explicit

Dim i As Long, j As Long
Dim StrTag As String, StrToGrid As String
Dim strSql As String, lngCounter As Long, lngCounterI As Long
Public cmdState As Integer
Dim m_dbl_LeftMargin As Double
Dim m_dbl_TopMargin  As Double
Dim SeekPrintL As Integer
Dim SeekPrint As Integer
'Dim m_bolPrintRight As Boolean   '2008/12/9 CANCEL BY SONIA
Private Const CB_SHOWDROPDOWN = &H14F
Dim StrCR05 As Variant              '往來類別用
Dim m_strCRexcept As String 'Added by Lydia 2025/08/08

Private Sub SetDataListWidth()
Dim ii As Integer

   GrdDataList.row = 0
   GrdDataList.col = 0: GrdDataList.Text = "V"
   GrdDataList.ColWidth(0) = 200
   
   GrdDataList.CellAlignment = flexAlignCenterCenter
   GrdDataList.col = 1: GrdDataList.Text = "記錄編號"
   GrdDataList.ColWidth(1) = 975
   
   GrdDataList.CellAlignment = flexAlignCenterCenter
   GrdDataList.col = 2: GrdDataList.Text = "日期"
   GrdDataList.ColWidth(2) = 825
   
   GrdDataList.CellAlignment = flexAlignCenterCenter
   GrdDataList.col = 3: GrdDataList.Text = "對象編號"
   GrdDataList.ColWidth(3) = 1000
   
   GrdDataList.CellAlignment = flexAlignLeftCenter 'flexAlignCenterCenter
   GrdDataList.col = 4: GrdDataList.Text = "名稱"
   GrdDataList.ColWidth(4) = 2000
   
   'Add By Sindy 2020/11/13
   GrdDataList.CellAlignment = flexAlignCenterCenter
   GrdDataList.col = 5: GrdDataList.Text = "建檔人員"
   GrdDataList.ColWidth(5) = 800
   '2020/11/13 END
   
   'Add By Sindy 2019/4/17
   GrdDataList.CellAlignment = flexAlignCenterCenter
   GrdDataList.col = 6: GrdDataList.Text = "往來類別"
   GrdDataList.ColWidth(6) = 1000
   '2019/4/17 END
   
   GrdDataList.CellAlignment = flexAlignLeftCenter 'flexAlignCenterCenter
   GrdDataList.col = 7: GrdDataList.Text = "主旨"
   GrdDataList.ColWidth(7) = 2500
   
   'Add By Sindy 2019/3/13
   For ii = 8 To GrdDataList.Cols - 1
      GrdDataList.ColWidth(ii) = 0
   Next ii
End Sub

Public Sub PubShowNextData()
Dim blnPrintAdd As Boolean
Dim ii As Integer
Dim j As Integer
Dim strTmp As String

   Select Case cmdState
      Case 0 '往來記錄
         Me.Enabled = False
         For i = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = i
         
         If Trim(GrdDataList.Text) = "V" Then
             GrdDataList.col = 0
             GrdDataList.Text = ""
             
             For j = 0 To GrdDataList.Cols - 1
                If j <> 1 Then
                    GrdDataList.col = j
                    GrdDataList.CellBackColor = QBColor(15)
                End If
             Next j
             
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            
            GrdDataList.col = 1
            Screen.MousePointer = vbHourglass
            
            strSql = "select * from contactrecord1 where cor01='" & Pub_RplStr(GrdDataList.Text) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               frm100101_19.Show
               frm100101_19.Tag = Pub_RplStr(GrdDataList.Text)
               frm100101_19.StrMenu
            Else
               frm100101_16.Show
               frm100101_16.Tag = Pub_RplStr(GrdDataList.Text)
               frm100101_16.StrMenu
            End If
            
            Screen.MousePointer = vbDefault
            GrdDataList.col = 0
            GrdDataList.Text = ""
            
            For j = 0 To GrdDataList.Cols - 1
               If j <> 1 Then
                   GrdDataList.col = j
                   GrdDataList.CellBackColor = QBColor(15)
               End If
            Next j
            
            Me.Enabled = True
            Exit Sub
         End If
         Next i
         Me.Enabled = True
            
      Case 1 '結束
         Unload frm140408
         
      Case Else
   End Select
End Sub

Private Sub cboSort_Click()
Dim iPos As Integer
   
   iPos = InStr(cboSort.Text, Chr(1))
   If iPos > 0 Then
      cboSort.Text = Left(cboSort.Text, iPos - 1)
   End If
End Sub

Private Sub cboSort_GotFocus()
   If cboSort.Locked = False Then
      CloseIme
      SendMessage cboSort.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cmdAddSort_Click()
   If AddList(lstSort, cboSort) = True Then
      Text10 = ComposeList(lstSort)
      cboSort = ""
   End If
   cboSort.SetFocus
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub cmdRemSort_Click()
   If RemoveList(lstSort) = True Then
      Text10 = ComposeList(lstSort)
      cboSort.SetFocus
   End If
End Sub

Private Sub cmdSearch_Click()
   Search_Process
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
      
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   'cmdOK(0).Enabled = True    '2008/12/9 CANCEL BY SONIA
   'cmdOK(1).Enabled = False   '2008/12/9 CANCEL BY SONIA
   
   Option1(0).Enabled = False
   Option1(1).Enabled = False
   Option1(2).Enabled = False
   Option3(0).Enabled = False
   Option3(1).Enabled = False
   
   bolToEndByNick = False
   'm_bolPrintRight = IsUserHasRightOfFunction("frm140408", strPrint, False)  '2008/12/9 CANCEL BY SONIA
   'Me.cmdOK(1).Enabled = m_bolPrintRight  '2008/12/9 CANCEL BY SONIA
   cmdState = -1
   
   AddCombo cboSort
   lstSort.Clear
   SetcboSort 'Add By Sindy 2019/3/11
   Label1.Caption = ""  'Added by Lydia 2022/01/26
   m_strCRexcept = Pub_GetCRExceptNo(Me.Name) 'Added by Lydia 2025/08/08
End Sub

'Add By Sindy 2019/3/11 往來類別
Private Sub SetcboSort()
   cboSort.Clear
   strSql = "select ac02,ac03 from allcode where ac01='11'" & _
            " order by ac02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   cboSort.AddItem ""
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         cboSort.AddItem RsTemp.Fields("ac02") & " " & RsTemp.Fields("ac03")
         RsTemp.MoveNext
      Loop
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140408 = Nothing
End Sub

Private Sub GrdDataList_Click()
   GrdDataList.Visible = False
   GrdDataList.row = GrdDataList.MouseRow
   GrdDataList.col = 0
   If GrdDataList.row <> 0 Then
      If GrdDataList.Text = "V" Then
         GrdDataList.Text = ""
         For i = 0 To GrdDataList.Cols - 1
            If i <> 1 Then
               GrdDataList.col = i
               GrdDataList.CellBackColor = QBColor(15)
            End If
         Next i
      Else
         GrdDataList.Text = "V"
         For i = 0 To GrdDataList.Cols - 1
            If i <> 1 Then
               GrdDataList.col = i
               GrdDataList.CellBackColor = &HFFC0C0
            End If
         Next i
      End If
   End If
   GrdDataList.Visible = True

End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 1 Then
       CloseIme
       Text2.SetFocus
   Else
       OpenIme
       Text2.SetFocus
   End If
End Sub

Private Sub Option2_Click(Index As Integer)
   Select Case Index
      '潛在客戶編號
      Case 0
         If Option2(0).Value = True Then
            Option2(1).Value = False
            Option2(2).Value = False
            
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
            
            Option3(0).Enabled = False
            Option3(1).Enabled = False
            
            Text1.SetFocus 'Add By Sindy 2012/4/11
         End If
      '潛在客戶/聯絡人名稱
      Case 1
         If Option2(1).Value = True Then
            Option1(0).Enabled = True
            Option1(0).Value = True
            Option1(1).Enabled = True
            Option1(2).Enabled = True
            
            Option2(0).Value = False
            Option2(2).Value = False
            
            Option3(0).Enabled = True
            Option3(1).Enabled = True
            Option3(1).Value = True     '2012/3/28 ADD BY SONIA
            
            Text2.SetFocus 'Add By Sindy 2012/4/11
         End If
      'E-Mail
      Case 2
         If Option2(2).Value = True Then
            Option2(0).Value = False
            Option2(1).Value = False
            
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
            
            Option3(0).Enabled = False
            Option3(1).Enabled = False
            
            Text9.SetFocus 'Add By Sindy 2012/4/11
         End If
      Case Else
   End Select
End Sub
Private Sub Text1_GotFocus()
   Me.Option2(0).Value = True
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
   CloseIme
End Sub

'Add By Sindy 2019/3/12
Private Sub Text1_LostFocus()
   If Text1 <> "" Then
      '1.申請人 2.代理人 3.潛在客戶
      If Left(Trim(Text1), 1) = "X" Then
         Text8 = "1"
      ElseIf Left(Trim(Text1), 1) = "Y" Then
         Text8 = "2"
      Else
         Text8 = "3"
      End If
   End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(0).Value = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   Me.Option2(1).Value = True
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
   '英文
   If Option1(1).Value = True Then
      CloseIme
   Else
      OpenIme
   End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(1).Value = True
End Sub
Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   '英文
   If Option1(1).Value = True Then
      CloseIme
   Else
      OpenIme
   End If
End Sub

Private Sub Text3_Click()
   If Text3 <> "" And GetStaffName(Text3) = "" Then
      Label1.Caption = ""
      MsgBox "開發人員不存在,請查核", vbCritical
      TextInverse Text3
      Exit Sub
   Else
      Label1.Caption = GetStaffName(Text3)
   End If
End Sub

Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
Cancel = False
   If Text3 <> "" And GetStaffName(Text3) = "" Then
      Cancel = True
      Label1.Caption = ""
      MsgBox "開發人員不存在,請查核", vbCritical
      TextInverse Text3
      Exit Sub
   Else
      Label1.Caption = GetStaffName(Text3)
   End If
End Sub

Private Sub Text4_GotFocus()
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
   If PUB_CheckKeyInDate(Me.Text4) = -1 Then
      Me.Text4.SetFocus
      Text4_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
   If PUB_CheckKeyInDate(Me.Text5) = -1 Then
      Me.Text5.SetFocus
      TextInverse Text5
      Exit Sub
   End If
   
   If Not nickChgRan(Text4, Text5, "開發日期") Then
      Text4.SetFocus
      TextInverse Text4
      Exit Sub
   End If
End Sub

Private Sub Text6_GotFocus()
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_LostFocus()
   If PUB_CheckKeyInDate(Me.Text6) = -1 Then
      Me.Text6.SetFocus
      TextInverse Text6
      Exit Sub
   End If
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_LostFocus()
   If PUB_CheckKeyInDate(Me.Text7) = -1 Then
      Me.Text7.SetFocus
      TextInverse Text7
      Exit Sub
   End If
   
   If Not nickChgRan(Text6, Text7, "往來日期") Then
      Text6.SetFocus
      TextInverse Text6
      Exit Sub
   End If
End Sub

Private Sub Text8_GotFocus()
   Text8.SelStart = 0
   Text8.SelLength = Len(Text8)
   CloseIme
End Sub

'Add By Sindy 2019/3/12
Private Sub Text8_LostFocus()
   '1.申請人 2.代理人 3.潛在客戶
   If Text8 = "1" And Left(Trim(Text1), 1) <> "X" Then
      Text1 = ""
   ElseIf Text8 = "2" And Left(Trim(Text1), 1) <> "Y" Then
      Text1 = ""
   ElseIf Text8 = "3" And Left(Trim(Text1), 1) <> "R" Then
      Text1 = ""
   End If
End Sub

Private Sub Text9_GotFocus()
   Me.Option2(2).Value = True
   Text9.SelStart = 0
   Text9.SelLength = Len(Text9)
   CloseIme
End Sub

Private Sub Text9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(2).Value = True
End Sub

Private Function ChgPotCustomer(ByVal strTemp As String) As String

On Error GoTo ErrHand
   If strTemp = "" Then GoTo ErrHand
   
   If Len(strTemp) = 9 Then
      ChgPotCustomer = "PCU01='" & Left(strTemp, 8) & "' AND PCU02='" & Right(strTemp, 1) & "'"
   Else
      ChgPotCustomer = "PCU01='" & strTemp & String(8 - Len(strTemp), "0") & "' AND PCU02='0'"
   End If
   Exit Function

ErrHand:
   ChgPotCustomer = "PCU01 IS NULL AND PCU02 IS NULL"
End Function

Private Function ComposeList(oList As ListBox, Optional p_iOpt As Integer = 0) As String
Dim iPos As Integer, stItem As String
   
   strExc(1) = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         If p_iOpt = 0 Then
            iPos = InStr(oList.List(intI), Chr(1))
            If iPos > 0 Then
               stItem = Left(oList.List(intI), iPos - 1)
            Else
               stItem = oList.List(intI)
            End If
         Else
            stItem = Format(oList.ItemData(intI), "00")
         End If
         If intI = 0 Then
            strExc(1) = stItem
         Else
            strExc(1) = strExc(1) & "," & stItem
         End If
      Next
   End If
   ComposeList = strExc(1)
End Function
Private Function AddList(oList As ListBox, oCombo As ComboBox, Optional p_iOpt As Integer = 0) As Boolean
Dim idx As Integer, bFound As Boolean, stNewItem As String, iNewItemData As Integer
Dim stSort As String, iPos As Integer
   
   If oCombo.Text = "" Then
      Exit Function
   End If
   
   '若有控制字元時後面為說明文字不抓
   iPos = InStr(oCombo, Chr(1))
   If iPos > 0 Then
      stNewItem = Left(oCombo, iPos - 1)
   Else
      stNewItem = oCombo
   End If
      
   If InStr(stNewItem, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      oCombo.SetFocus
      Exit Function
   End If

   If stNewItem <> "" Then
      If bFound = False Then
         oList.AddItem stNewItem, 0
         If p_iOpt <> 0 Then
            oList.ItemData(0) = oCombo.ItemData(oCombo.ListIndex)
         End If
         AddList = True
      End If
   End If
End Function

Private Function RemoveList(oList As ListBox) As Boolean
Dim ii As Integer
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList = True
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function

Private Sub AddCombo(oCombo As ComboBox)
   With oCombo
      .Clear
      'Modified by Morgan 2018/3/23 調整內容及順序 --David
      '.AddItem "IP諮詢"
      '.AddItem "非IP法律諮詢"
      '.AddItem "詢價"
      '.AddItem "申請所需文件"
      '.AddItem "利益衝突"
      '.AddItem "互惠"
      '.AddItem "詢問IP侵害"
      '.AddItem "訪談" & Chr(1) & "(包括來所訪問、出國拜訪、國際會議)"
      '.AddItem "客戶特別指示" & Chr(1) & "(譬如:不要寄confirmation copy, 聯絡方式只限fax or e-mail, 付款方式只限credit card or cheque…)"
      '.AddItem "報價"   'add by sonia 2018/1/12
      '.AddItem "慰問"   'add by sonia 2018/2/9
      .AddItem "總指示"
      .AddItem "來函詢價/申請文件"
      .AddItem "利益衝突調查"
      .AddItem "合作契約"
      .AddItem "催款"
      .AddItem "投標/固定價"
      .AddItem "網路平台/使用費"
      .AddItem "問卷"
      .AddItem "互惠"
      .AddItem "IP諮詢"
      .AddItem "訪談" & Chr(1) & "(包括來所訪問、出國拜訪、國際會議)"
      .AddItem "代理人提供報價"
      .AddItem "慰問"
      'end 2018/3/23
   End With
End Sub

Private Function Search_Process()
Dim strCon1 As String
Dim strCon2 As String
Dim strCon3 As String
Dim strCon4 As String
Dim strCW As String 'Add By Sindy 2021/4/7
Dim strCheckWay As String
Dim StrCR04 As String
Dim k As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim StrSQLa As String
Dim s As Integer, ii As Integer
Dim strKey As String
Dim strCR As String, strCOR As String
Dim varTmp As Variant
   
   'Modify By Sindy 2019/3/11
   'If Text6 = "" And Text7 = "" And Text8 = "" And Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" And Text9 = "" And lstSort = "" Then
   If Text6 = "" And Text7 = "" And Text8 = "" And Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" And Text9 = "" And cboSort.Text = "" Then
      MsgBox ("條件不可空白,請至少輸入一項")
      Text6.SetFocus
      Exit Function
   End If
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/24 清除查詢印表記錄檔欄位
   Screen.MousePointer = vbHourglass
   GrdDataList.Clear
   GrdDataList.Rows = 2
   'SetDataListWidth
   StrSQLa = ""
   strCon1 = "" '客戶檔
   strCon2 = "" '國外代理人檔
   strCon3 = "" '潛在客戶
   strCon4 = "" '國內潛在客戶
   strCR = "" '國外往來記錄
   strCOR = "" '國內往來記錄
   strCW = "" '客戶平台
   
   '若為國內智權人員或國內工程師, 不可查代理人資料
   'Modify By Sindy 2011/01/04 取消
'   If bolFNation = False Then
'      StrSQLa = " And FA01<'Y' "
'   End If
   
   '往來日期
   If Len(Text6) <> 0 And Len(Text7) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Left(Label5, 5) & Text6 & "-" & Text7 'Add By Sindy 2010/12/24
      strCR = strCR & "and CR02>='" & ChangeTStringToWString(Text6) & "' AND CR02<='" & ChangeTStringToWString(Text7) & "' " '國外往來記錄
      strCOR = strCOR & "and COR02>='" & ChangeTStringToWString(Text6) & "' AND COR02<='" & ChangeTStringToWString(Text7) & "' " '國內往來記錄
   End If
   
   '編號
   If Option2(0).Value = True And Text1 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Text1 'Add By Sindy 2010/12/24
      strCon1 = strCon1 & "AND CU01='" & Left(GetNewFagent(Text1), 8) & "' "
      strCon2 = strCon2 & "AND FA01='" & Left(GetNewFagent(Text1), 8) & "' "
      strCon3 = strCon3 & "AND PCU01='" & Left(GetNewFagent(Text1), 8) & "' "
      strCon4 = strCon4 & "AND POC01='" & Left(GetNewFagent(Text1), 8) & "' "
      
   '名稱
   ElseIf Option2(1).Value = True And Text2 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Text2 'Add By Sindy 2010/12/24
      '以編號或名稱
      '模糊比對
      If Option3(0).Value = False Then
         pub_QL05 = pub_QL05 & ";" & Option3(1).Caption 'Add By Sindy 2010/12/24
         strCheckWay = ">0"
      '字首比對
      Else
         pub_QL05 = pub_QL05 & ";" & Option3(0).Caption 'Add By Sindy 2010/12/24
         strCheckWay = "=1"
      End If
      
      '中文
      If Option1(0).Value = True Then
         pub_QL05 = pub_QL05 & ";" & "以" & Option1(0).Caption & "查詢" 'Add By Sindy 2010/12/24
         '申請人
         strCon1 = strCon1 & "AND (instr(CU04,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
         '代理人
         strCon2 = strCon2 & "AND (instr(fa04,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
         '查潛在客戶
         strCon3 = strCon3 & "AND (instr(pcu08,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
         '國內潛在客戶
         strCon4 = strCon4 & "AND (instr(poc03,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
         
      '英文
      ElseIf Option1(1).Value = True Then
         pub_QL05 = pub_QL05 & ";" & "以" & Option1(1).Caption & "查詢" 'Add By Sindy 2010/12/24
         '申請人
         strCon1 = strCon1 & "AND (instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Text2)) & "')" & strCheckWay & ") "
         '代理人
         strCon2 = strCon2 & "AND (instr(upper(fa05||' '||fa63||' '||fa64||' '||fa65),'" & UCase(ChgSQL(Text2)) & "')" & strCheckWay & ") "
         '查潛在客戶
         strCon3 = strCon3 & "AND (instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Text2)) & "')" & strCheckWay & ") "
         '國內潛在客戶
         strCon4 = strCon4 & "AND (instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Text2)) & "')" & strCheckWay & ") "
         
      '日文
      ElseIf Option1(2).Value = True Then
         pub_QL05 = pub_QL05 & ";" & "以" & Option1(2).Caption & "查詢" 'Add By Sindy 2010/12/24
         '申請人
         strCon1 = strCon1 & "AND (instr(CU06,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
         '代理人
         strCon2 = strCon2 & "AND (instr(fa06,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
         '查潛在客戶
         strCon3 = strCon3 & "AND (instr(PCU07,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
         '國內潛在客戶
         strCon4 = strCon4 & "AND (instr(poc27,'" & ChgSQL(Text2) & "')" & strCheckWay & ") "
      End If
      'Add By Sindy 2021/4/7
      '客戶平台
      strCW = strCW & "AND (instr(upper(cw12),'" & UCase(ChgSQL(Text2)) & "')" & strCheckWay & ") "
      '2021/4/7 END
      
   'E-Mail
   ElseIf Option2(2).Value = True And Text9 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & Text9 'Add By Sindy 2010/12/24
      '申請人
      'Modified by Lydia 2024/09/18 +財務副本信箱CU200
      strCon1 = strCon1 & "AND (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Text9)) & "')>0 or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Text9)) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Text9)) & "')>0  or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Text9)) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Text9)) & "') > 0 or instr(NLS_Upper(CU200),'" & UCase(ChgSQL(Text9)) & "')>0) "
      '代理人
       'Modified by Lydia 2018/07/20 +FA105 財務信箱(CF)
       'Modified by Lydia 2024/09/18 +財務副本信箱FA134
      strCon2 = strCon2 & "AND (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Text9)) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Text9)) & "')> 0 or instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Text9)) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Text9)) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Text9)) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Text9)) & "') > 0 or instr(NLS_Upper(FA134),'" & UCase(ChgSQL(Text9)) & "')> 0) "
      '潛在客戶
      strCon3 = strCon3 & "AND (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Text9)) & "') >0) "
      '國內潛在客戶
      strCon4 = strCon4 & "AND (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Text9)) & "') >0) "
   End If
   
   '開發者
   If Len(Text3) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label4 & Text3 & Label1 'Add By Sindy 2010/12/24
      '申請人
      strCon1 = strCon1 & "AND instr(CU129,'" & ChgSQL(Text3) & "')>0 "
      '代理人
      strCon2 = strCon2 & "AND instr(fa94,'" & ChgSQL(Text3) & "')> 0 "
      '潛在客戶
      strCon3 = strCon3 & "AND instr(PCU38,'" & ChgSQL(Text3) & "')>0 "
      '國內潛在客戶
      strCon4 = strCon4 & "AND instr(POC13,'" & ChgSQL(Text3) & "')>0 "
   End If
   
   '往來對象
   Call Text1_LostFocus 'Add By Sindy 2020/11/13
   If Len(Text8) <> 0 Then
      Select Case Text8
         '申請人
         Case "1"
            pub_QL05 = pub_QL05 & ";" & Left(Label3, 5) & "1.申請人" 'Add By Sindy 2010/12/24
            strCR = strCR & "AND SUBSTR(CR03,1,1)='X' " '國外往來記錄
            strCOR = strCOR & "AND SUBSTR(COR03,1,1)='X' " '國內往來記錄
         '代理人
         Case "2"
            pub_QL05 = pub_QL05 & ";" & Left(Label3, 5) & "2.代理人" 'Add By Sindy 2010/12/24
            strCR = strCR & "AND SUBSTR(CR03,1,1)='Y' " '國外往來記錄
            strCOR = strCOR & "AND SUBSTR(COR03,1,1)='Y' " '國內往來記錄
         '潛在客戶
         Case "3"
            pub_QL05 = pub_QL05 & ";" & Left(Label3, 5) & "3.潛在客戶" 'Add By Sindy 2010/12/24
            strCR = strCR & "AND SUBSTR(CR03,1,1)='R' " '國外往來記錄
            strCOR = strCOR & "AND SUBSTR(COR03,1,1)='R' " '國內往來記錄
      End Select
   End If
   
   '開發日期
   If Text4 <> "" And Text5 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label6 & Text4 & "-" & Text5 'Add By Sindy 2010/12/24
      '申請人
      strCon1 = strCon1 & "and cu14 >='" & ChangeTStringToWString(Text4) & "' and cu14 <='" & ChangeTStringToWString(Text5) & "' "
      '代理人
      strCon2 = strCon2 & "and fa11 >='" & ChangeTStringToWString(Text4) & "' and fa11 <='" & ChangeTStringToWString(Text5) & "' "
      '潛在客戶
      strCon3 = strCon3 & "and PCU37 >='" & ChangeTStringToWString(Text4) & "' and PCU37 <='" & ChangeTStringToWString(Text5) & "' "
      '國內潛在客戶
      strCon4 = strCon4 & "and POC12 >='" & ChangeTStringToWString(Text4) & "' and POC12 <='" & ChangeTStringToWString(Text5) & "' "
   End If
   
   '往來類別
'   StrCR05 = Split(Text10, ",")
'   If Trim(Text10) <> "" Then
'      pub_QL05 = pub_QL05 & ";" & Label7 & Trim(Text10) 'Add By Sindy 2010/12/24
'   End If
   'Modify By Sindy 2019/3/11
   If Trim(cboSort.Text) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label7 & Trim(cboSort.Text)
      varTmp = Split(cboSort.Text, " ")
      strCR = strCR & "and CR05='" & Trim(varTmp(0)) & "' " '國外往來記錄
      'Add By Sindy 2020/11/13 有輸入"往來類別"查詢,不顯示"國內往來記錄"
      strCOR = strCOR & "AND COR01='XXX' "
   End If
   '2019/3/11 END
   
   '2008/12/8 MODIFY BY SONIA
   'strSQL = "SELECT ' ' as V,CR01 AS 記錄編號,CR02 AS 日期,CU01||DECODE(CR04,'','','-'||CR04) AS 對象編號,CR04  AS 名稱,CR06 AS 主旨,CR05 FROM ContactRecord,Customer Where  CU01=substr(cr03,1,8)  " & strCon1
   'strSQL = strSQL & " union all SELECT ' ' AS V,CR01 AS 記錄編號,CR02 AS 日期,FA01||DECODE(CR04,'','','-'||CR04) AS 對象編號,CR04  AS 名稱,CR06 AS 主旨,CR05 FROM ContactRecord,fagent Where  fa01=substr(cr03,1,8)" & StrSQLa & strCon2
   'strSQL = strSQL & " union all SELECT ' ' AS V,CR01 AS 記錄編號,CR02 AS 日期,PCU01||DECODE(CR04,'','','-'||CR04) AS 對象編號,CR04  AS 名稱,CR06 AS 主旨,CR05 FROM ContactRecord,PotCustomer  Where pcu01=substr(cr03,1,8)" & strCon3
                       strSql = "SELECT ' ' as V,CR01 AS 記錄編號,sqldatet(CR02) AS 日期,CU01||DECODE(CR04,'','','-'||CR04) AS 對象編號,CR04 AS 名稱,st02 as 建檔人員,AC03 AS 往來類別,CR06 AS 主旨,CR05,DECODE(CU64,'1',nvl(CU04,nvl(CU05,CU06)),'3',nvl(CU06,nvl(CU04,CU06)),nvl(CU05,nvl(CU04,CU06))),'CR' FROM ContactRecord,Customer,allcode,staff Where ac01='11' and cr05=ac02(+) and CU01=substr(cr03,1,8) and CU02='0' and cr12=st01(+) " & strCon1 & strCR
   strSql = strSql & " union all SELECT ' ' AS V,CR01 AS 記錄編號,sqldatet(CR02) AS 日期,FA01||DECODE(CR04,'','','-'||CR04) AS 對象編號,CR04 AS 名稱,st02 as 建檔人員,AC03 AS 往來類別,CR06 AS 主旨,CR05,DECODE(FA31,'1',nvl(FA04,nvl(FA05,FA06)),'3',nvl(FA06,nvl(FA04,FA06)),nvl(FA05,nvl(FA04,FA06))),'CR' FROM ContactRecord,fagent,allcode,staff Where ac01='11' and cr05=ac02(+) and fa01=substr(cr03,1,8) and fa02='0' and cr12=st01(+) " & StrSQLa & strCon2 & strCR
   strSql = strSql & " union all SELECT ' ' AS V,CR01 AS 記錄編號,sqldatet(CR02) AS 日期,PCU01||DECODE(CR04,'','','-'||CR04) AS 對象編號,CR04 AS 名稱,st02 as 建檔人員,AC03 AS 往來類別,CR06 AS 主旨,CR05,DECODE(PCU36,'1',nvl(PCU08,nvl(PCU03,PCU07)),'3',nvl(PCU07,nvl(PCU08,PCU03)),nvl(PCU03,nvl(PCU08,PCU07))),'CR' FROM ContactRecord,PotCustomer,allcode,staff Where ac01='11' and cr05=ac02(+) and pcu01=substr(cr03,1,8) and pcu02='0' and cr12=st01(+) " & strCon3 & strCR
   strSql = strSql & " union all SELECT ' ' AS V,CR01 AS 記錄編號,sqldatet(CR02) AS 日期,POC01||DECODE(CR04,'','','-'||CR04) AS 對象編號,CR04 AS 名稱,st02 as 建檔人員,AC03 AS 往來類別,CR06 AS 主旨,CR05,POC03,'CR' FROM ContactRecord,PotCustomer1,allcode,staff Where ac01='11' and cr05=ac02(+) and poc01=substr(cr03,1,8) and poc02='0' and cr12=st01(+) " & strCon4 & strCR
   
   'Add By Sindy 2021/4/7
   strSql = strSql & " union all SELECT ' ' AS V,CR01 AS 記錄編號,sqldatet(CR02) AS 日期,CW01 AS 對象編號,CW12 AS 名稱,st02 as 建檔人員,AC03 AS 往來類別,CR06 AS 主旨,CR05,CW12,'CR' FROM ContactRecord,CustWeb,allcode,staff Where ac01='11' and cr05=ac02(+) and cw01=cr03 and cr12=st01(+) " & strCW & strCR
   
   strSql = strSql & " union all SELECT ' ' as V,COR01 AS 記錄編號,sqldatet(COR02) AS 日期,COR03 AS 對象編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,st02 as 建檔人員,'' AS 往來類別,COR04 AS 主旨,'',DECODE(CU64,'1',nvl(CU04,nvl(CU05,CU06)),'3',nvl(CU06,nvl(CU04,CU06)),nvl(CU05,nvl(CU04,CU06))),'COR' FROM ContactRecord1,Customer,staff Where CU01=substr(cor03,1,8) and CU02='0' and cor06=st01(+) " & strCon1 & strCOR
   strSql = strSql & " union all SELECT ' ' AS V,COR01 AS 記錄編號,sqldatet(COR02) AS 日期,COR03 AS 對象編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) AS 名稱,st02 as 建檔人員,'' AS 往來類別,COR04 AS 主旨,'',DECODE(FA31,'1',nvl(FA04,nvl(FA05,FA06)),'3',nvl(FA06,nvl(FA04,FA06)),nvl(FA05,nvl(FA04,FA06))),'COR' FROM ContactRecord1,fagent,staff Where fa01=substr(cor03,1,8) and fa02='0' and cor06=st01(+) " & StrSQLa & strCon2 & strCOR
   strSql = strSql & " union all SELECT ' ' AS V,COR01 AS 記錄編號,sqldatet(COR02) AS 日期,COR03 AS 對象編號,NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,st02 as 建檔人員,'' AS 往來類別,COR04 AS 主旨,'',DECODE(PCU36,'1',nvl(PCU08,nvl(PCU03,PCU07)),'3',nvl(PCU07,nvl(PCU08,PCU03)),nvl(PCU03,nvl(PCU08,PCU07))),'COR' FROM ContactRecord1,PotCustomer,staff Where pcu01=substr(cor03,1,8) and pcu02='0' and cor06=st01(+) " & strCon3 & strCOR
   strSql = strSql & " union all SELECT ' ' AS V,COR01 AS 記錄編號,sqldatet(COR02) AS 日期,COR03 AS 對象編號,NVL(POC03,DECODE(poc23,null,poc27,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26))) AS 名稱,st02 as 建檔人員,'' AS 往來類別,COR04 AS 主旨,'',POC03,'COR' FROM ContactRecord1,PotCustomer1,staff Where poc01=substr(cor03,1,8) and poc02='0' and cor06=st01(+) " & strCon4 & strCOR
   strSql = strSql & " ORDER BY 記錄編號 desc"
   '2008/12/8 END

   '--------
   intI = 1
   SetDataListWidth
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
      
   Label9.Caption = "筆數：" 'Add By Sindy 2019/7/12
   If intI = 1 Then
      'Added by Lydia 2025/08/08 國外往來記錄的維護及查詢限制
      If m_strCRexcept <> "" Then
          strExc(1) = "select * from (" & Mid(UCase(strSql), 1, InStr(UCase(strSql), "ORDER") - 1) & ") where 記錄編號 not in (" & GetAddStr(m_strCRexcept) & ")" & _
                      " ORDER BY 記錄編號 desc"
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
          If intI = 1 Then
             If RsTemp.RecordCount <> adoRecordset.RecordCount Then
                 MsgBox "限閱往來記錄" & adoRecordset.RecordCount - RsTemp.RecordCount & "筆！", vbInformation
                 Set adoRecordset = RsTemp
             End If
          Else
              Screen.MousePointer = vbDefault
              Me.Enabled = True
              MsgBox "限閱往來記錄" & adoRecordset.RecordCount - RsTemp.RecordCount & "筆！", vbInformation
              Exit Function
          End If
      End If
      'end 2025/08/08
      '2008/11/17 MODIFY BY SONIA 逐筆抓聯絡人名稱,逐筆檢查往來類別
      Set GrdDataList.Recordset = adoRecordset
      Label9.Caption = "筆數：" & adoRecordset.RecordCount 'Add By Sindy 2019/7/12
      SetDataListWidth
      'adoRecordset.MoveFirst
      Me.Enabled = False
      'Do While Not adoRecordset.EOF
'         If adoRecordset.Fields(8) = "CR" And Trim(Text10) <> "" Then
'            '檢查往來類別
'            k = 0
'            For i = 0 To UBound(StrCR05)
'               If InStr(UCase(adoRecordset.Fields(6)), UCase(StrCR05(i))) > 0 Then
'                  k = k + 1
'               End If
'            Next i
'            If k < UBound(StrCR05) + 1 Then
'               GoTo NextRecord
'            End If
'         End If
'         GrdDataList.Rows = GrdDataList.Rows + 1
'         GrdDataList.row = GrdDataList.Rows - 2
'         '記錄編號
'         If Not IsNull(adoRecordset.Fields(1)) Then
'            GrdDataList.TextMatrix(GrdDataList.row, 1) = adoRecordset.Fields(1)
'         End If
'         '日期
'         If Not IsNull(adoRecordset.Fields(2)) Then
'            GrdDataList.TextMatrix(GrdDataList.row, 2) = Format(Left(adoRecordset.Fields(2), 4) - 1911) + "/" + Mid(adoRecordset.Fields(2), 5, 2) + "/" + Right(adoRecordset.Fields(2), 2)
'         End If
'         '主旨
'         If Not IsNull(adoRecordset.Fields(5)) Then
'            GrdDataList.TextMatrix(GrdDataList.row, 5) = adoRecordset.Fields(5)
'         End If
      For ii = 1 To GrdDataList.Rows - 1
         '名稱
         If Not IsNull(GrdDataList.TextMatrix(ii, 4)) Then
            StrCR04 = GrdDataList.TextMatrix(ii, 4)
         '2008/12/8 ADD BY SONIA
         Else
            StrCR04 = ""
         '2008/12/8 END
         End If
         '往來對象
         If Not IsNull(GrdDataList.TextMatrix(ii, 3)) Then
'            GrdDataList.TextMatrix(GrdDataList.row, 3) = adoRecordset.Fields(3)
            strKey = Mid(GrdDataList.TextMatrix(ii, 3), 1, 8)
         '2008/12/8 ADD BY SONIA
         Else
            strKey = ""
         '2008/12/8 END
         End If
         If GrdDataList.TextMatrix(ii, 10) = "COR" Then '國內往來
            '對象名稱
            If Not IsNull(GrdDataList.TextMatrix(ii, 9)) Then
               GrdDataList.TextMatrix(ii, 4) = GrdDataList.TextMatrix(ii, 9)
            Else
               GrdDataList.TextMatrix(ii, 4) = "" & GrdDataList.TextMatrix(ii, 4)
            End If
         ElseIf GrdDataList.TextMatrix(ii, 10) = "CR" Then '國外往來
            '2008/12/8 ADD BY SONIA  先抓往來對象名稱
            If Not IsNull(GrdDataList.TextMatrix(ii, 9)) Then
               GrdDataList.TextMatrix(ii, 4) = GrdDataList.TextMatrix(ii, 9)
               'Modify By Sindy 2021/4/7 + And Len(grdDataList.TextMatrix(ii, 3)) <> 4
               '客戶平台名稱不要加-
               If StrCR04 <> "" And Len(GrdDataList.TextMatrix(ii, 3)) <> 4 Then
                  GrdDataList.TextMatrix(ii, 4) = GrdDataList.TextMatrix(ii, 4) & "-"
               End If
            End If
            '2008/12/8 END
            '抓聯絡人名稱
            If StrCR04 <> "" Then
               strSql = "SELECT nvl(pcc05,nvl(pcc03,pcc04)) NM FROM PotCustCont " & _
                        "WHERE PCC01 = '" & strKey & "'" & " AND PCC02 IN ('" & StrCR04 & "') ORDER BY PCC02"
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  Do While Not rsTmp.EOF
                     For i = 1 To rsTmp.RecordCount
                        GrdDataList.TextMatrix(ii, 4) = GrdDataList.TextMatrix(ii, 4) & rsTmp.Fields("NM") & ";"
                        Exit For
                     Next i
                     rsTmp.MoveNext
                  Loop
               End If
               rsTmp.Close
            End If
         End If
      Next ii
         
'NextRecord:
'         adoRecordset.MoveNext
'      Loop
'      adoRecordset.Close
      Me.Enabled = True
      
'      If GrdDataList.Rows > 2 Then
         InsertQueryLog (GrdDataList.Rows - 1) 'Add By Sindy 2010/12/24
         'GrdDataList.Rows = GrdDataList.Rows - 1
'      '2008/12/8 ADD BY SONIA
'      Else
'         InsertQueryLog (0) 'Add By Sindy 2010/12/24
'         ShowNoData
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'         'tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         GrdDataList.Clear
'         SetDataListWidth
'         Exit Function
'      '2008/12/8 END
'      End If
      
      Screen.MousePointer = vbDefault
      
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/24
      ShowNoData
      Screen.MousePointer = vbDefault
      Me.Enabled = True
      'tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      grdDataList.Clear
'      SetDataListWidth
      Exit Function
   End If
   '--------
   
   With Me.GrdDataList
      For i = 0 To .Rows - 1
         .row = i
         .col = 1
         If Right(.Text, 1) = "$" Then
           .CellBackColor = &HFF&
         End If
      Next i
   End With
   
   '若只有一筆資料, 則直接設定為點選此筆資料
   With Me.GrdDataList
      If .Rows = 2 Then
         .row = 1
         .col = 1
         If .Text <> "" Then
           .Visible = False
           .row = 1
           .col = 0
           .Text = "V"
           For i = 0 To .Cols - 1
               If i <> 1 Then
                   .col = i
                   .CellBackColor = &HFFC0C0
               End If
           Next i
           .Visible = True
         End If
      End If
   End With

End Function
