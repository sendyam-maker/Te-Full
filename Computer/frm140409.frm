VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140409 
   BorderStyle     =   1  '單線固定
   Caption         =   "潛在客戶名條列印"
   ClientHeight    =   6084
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8532
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   8532
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   90
      TabIndex        =   46
      Top             =   4140
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "frm140409.frx":0000
      Left            =   1050
      List            =   "frm140409.frx":0002
      TabIndex        =   14
      Text            =   "cboSort"
      Top             =   4410
      Width           =   6500
   End
   Begin VB.CommandButton cmdRemSort 
      Caption         =   "移除 ↓"
      Height          =   285
      Left            =   7680
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4110
      Width           =   735
   End
   Begin VB.CommandButton cmdAddSort 
      Caption         =   "新增↑"
      Height          =   285
      Left            =   7680
      TabIndex        =   15
      Top             =   4410
      Width           =   735
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
      ItemData        =   "frm140409.frx":0004
      Left            =   1050
      List            =   "frm140409.frx":000B
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3810
      Width           =   6500
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   4305
      TabIndex        =   41
      Top             =   2640
      Width           =   4155
      Begin VB.TextBox txtPCC 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   810
         MaxLength       =   70
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.CommandButton cmdAddTit 
         Caption         =   "<- 新增"
         Height          =   255
         Left            =   45
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdRemoveTit 
         Caption         =   "移除 ->"
         Height          =   255
         Left            =   45
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   390
         Width           =   735
      End
      Begin VB.ComboBox cboTitle 
         Height          =   300
         ItemData        =   "frm140409.frx":0018
         Left            =   810
         List            =   "frm140409.frx":001A
         TabIndex        =   8
         Text            =   "cboTitle"
         Top             =   120
         Width           =   3300
      End
   End
   Begin VB.ListBox lstTitle 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      ItemData        =   "frm140409.frx":001C
      Left            =   1050
      List            =   "frm140409.frx":0023
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2670
      Width           =   3180
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   4290
      TabIndex        =   39
      Top             =   1740
      Width           =   4155
      Begin VB.TextBox txtPCC 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   810
         MaxLength       =   70
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.CommandButton cmdRemoveDept 
         Caption         =   "移除 ->"
         Height          =   255
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton cmdAddDept 
         Caption         =   "<- 新增"
         Height          =   255
         Left            =   45
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         ItemData        =   "frm140409.frx":0031
         Left            =   810
         List            =   "frm140409.frx":0033
         TabIndex        =   5
         Text            =   "cboDept"
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.ListBox lstDept 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      ItemData        =   "frm140409.frx":0035
      Left            =   1050
      List            =   "frm140409.frx":003C
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3180
   End
   Begin VB.ComboBox cboCity 
      Height          =   300
      Left            =   1050
      TabIndex        =   4
      Text            =   "cboCity"
      Top             =   1380
      Width           =   5415
   End
   Begin VB.TextBox txtPCU06 
      Height          =   285
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   13
      Top             =   3480
      Width           =   705
   End
   Begin VB.TextBox txtPCU05 
      Height          =   285
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   12
      Top             =   3480
      Width           =   705
   End
   Begin VB.TextBox txtPCU 
      Height          =   285
      Index           =   4
      Left            =   1050
      MaxLength       =   6
      TabIndex        =   11
      Top             =   3120
      Width           =   705
   End
   Begin VB.TextBox txtPCU 
      Height          =   285
      Index           =   2
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   2
      Top             =   750
      Width           =   705
   End
   Begin VB.TextBox txtPCU 
      Height          =   285
      Index           =   1
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   1
      Top             =   750
      Width           =   705
   End
   Begin VB.TextBox txtPCU 
      Height          =   270
      Index           =   0
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   0
      Top             =   420
      Width           =   250
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6720
      TabIndex        =   20
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   7545
      TabIndex        =   21
      Top             =   30
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   660
      Left            =   90
      TabIndex        =   22
      Top             =   4800
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   43
         Top             =   240
         Width           =   4400
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   23
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1440
      TabIndex        =   44
      Top             =   5460
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1440
      TabIndex        =   45
      Top             =   5760
      Width           =   705
   End
   Begin MSForms.ComboBox cboPCU11 
      Height          =   285
      Left            =   1050
      TabIndex        =   3
      Top             =   1050
      Width           =   1710
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3016;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "PS:有部門或職稱條件時針對聯絡人過濾資料, 名條會印聯絡人           無部門及職稱條件時針對潛在客戶過濾資料, 名條不印聯絡人"
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   3360
      TabIndex        =   47
      Top             =   5520
      Width           =   5040
   End
   Begin VB.Label Label9 
      Height          =   270
      Left            =   1920
      TabIndex        =   38
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "(1亞洲 2.美洲 3.歐洲 4.非洲 5.大洋洲)"
      Height          =   270
      Left            =   1350
      TabIndex        =   37
      Top             =   420
      Width           =   3000
   End
   Begin VB.Label Label4 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1770
      TabIndex        =   36
      Top             =   3540
      Width           =   105
   End
   Begin VB.Label Label4 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   35
      Top             =   810
      Width           =   105
   End
   Begin VB.Label Label1 
      Caption         =   "往來類別："
      Height          =   270
      Index           =   8
      Left            =   90
      TabIndex        =   34
      Top             =   3810
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "往來日期："
      Height          =   270
      Index           =   7
      Left            =   90
      TabIndex        =   33
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "開發人員："
      Height          =   270
      Index           =   6
      Left            =   90
      TabIndex        =   32
      Top             =   3150
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "職　　稱："
      Height          =   270
      Index           =   5
      Left            =   90
      TabIndex        =   31
      Top             =   2730
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "部　　門："
      Height          =   270
      Index           =   4
      Left            =   90
      TabIndex        =   30
      Top             =   1830
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "城　　市："
      Height          =   270
      Index           =   3
      Left            =   90
      TabIndex        =   29
      Top             =   1410
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "類　　別："
      Height          =   270
      Index           =   2
      Left            =   90
      TabIndex        =   28
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "國　　籍："
      Height          =   270
      Index           =   1
      Left            =   90
      TabIndex        =   27
      Top             =   750
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "洲　　別："
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   26
      Top             =   420
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   25
      Top             =   5520
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   24
      Top             =   5820
      Width           =   3240
   End
End
Attribute VB_Name = "frm140409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'2008/11/11 ADD BY TONI
Option Explicit

Dim idx As Integer
Dim rsNation As ADODB.Record '國家檔 2008/11/11
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Dim m_dbl_LeftMargin  As Double '橫軸偏移值
Dim m_dbl_TopMargin  As Double '縱軸偏移值
Dim m_PrinterName As String
Dim Prn As Printer
Dim strSql As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
Dim strSQL1 As String
Dim strTemp As String, stSQLa As String
Dim strCity As String


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
         AddList = True
      End If
   End If
End Function

Private Sub cboCity_Change()
   If cboCity <> "" Then
      For intI = 0 To cboCity.ListCount - 1
         If cboCity = cboCity.List(intI) Then
            cboCity.ListIndex = intI
            cboCity.SelStart = Len(cboCity)
         End If
      Next
   End If
End Sub

' 設定列印的印表機
Public Sub SetPrinter(ByVal strPrinterName As String)
   m_PrinterName = strPrinterName
End Sub

Private Sub cboCity_GotFocus()
   CloseIme
   If cboCity.Locked = False Then
      SendMessage cboCity.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cboCity_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboSort_Click()
Dim iPos As Integer
   
   iPos = InStr(cboSort.Text, Chr(1))
   If iPos > 0 Then
      cboSort.Text = Left(cboSort.Text, iPos - 1)
   End If
End Sub

'新增部門
Private Sub cmdAddDept_Click()
   If InStr(cboDept, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      cboDept.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo cboDept, lstDept
   txtPCC(6) = ComposeList(lstDept)
   cboDept.SetFocus
End Sub

'新增類別
Private Sub cmdAddSort_Click()
   If AddList(lstSort, cboSort) = True Then
      Text2 = ComposeList(lstSort)
      cboSort = ""
   End If
   cboSort.SetFocus
End Sub

'新增職稱
Private Sub cmdAddTit_Click()
   If InStr(cboTitle, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      cboTitle.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo cboTitle, lstTitle
   txtPCC(7) = ComposeList(lstTitle)
   cboTitle.SetFocus
End Sub

Private Sub cmdBack_Click()
   bolToEndByNick = True
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim strTempName As String
Dim Index As Integer

   '若整批列印地址條
   'Modify By Sindy 2021/6/28
   'If txtPCU(0) = "" And txtPCU(1) = "" And txtPCU(2) = "" And txtPCU(3) = "" And txtPCU(4) = "" And txtPCU05 = "" And txtPCU06 = "" And cboCity = "" And lstDept = "" And lstTitle = "" And lstSort = "" Then
   If txtPCU(0) = "" And txtPCU(1) = "" And txtPCU(2) = "" And Len(Trim(cboPCU11)) = 0 And txtPCU(4) = "" And txtPCU05 = "" And txtPCU06 = "" And cboCity = "" And lstDept = "" And lstTitle = "" And lstSort = "" Then
   '2021/6/28 END
      MsgBox "請輸入任一項條件"
      txtPCU(0).SetFocus
      Exit Sub
   End If
   
   '***************  90.11.14  NICKC
   ' 設定印表機
   If IsEmptyText(m_PrinterName) Then
       j = Combo1.ListIndex
       Set Printer = Printers(j)
   Else
       For Each Prn In Printers
           If Prn.DeviceName = m_PrinterName Then
               Set Printer = Prn
               Exit For
           End If
       Next
   End If
   Printer.Orientation = 1
   '****************************************
   DoEvents
   Screen.MousePointer = vbHourglass
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/24 清除查詢印表記錄檔欄位
   Process
    
   '印完後預設回預設印表機
   Set Printer = Printers(SeekPrint)
   Printer.Orientation = SeekPrintL
   bolToEndByNick = True
   Screen.MousePointer = vbDefault
   txtPCU(0).SetFocus
End Sub

'移除部門
Private Sub cmdRemoveDept_Click()
   RemoveList lstDept
   txtPCC(6) = ComposeList(lstDept)
End Sub

'移除職稱
Private Sub cmdRemoveTit_Click()
   RemoveList lstTitle
   txtPCC(7) = ComposeList(lstTitle)
End Sub

'移除往來類別
Private Sub cmdRemSort_Click()
   If RemoveList1(lstSort) = True Then
      Text2 = ComposeList(lstSort)
      cboSort.SetFocus
   End If
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
     
   '改不排除預設印表機
   PUB_SetPrinter Me.Name, Me.Combo1, , , SeekPrint, Me.Text1(7), Me.Text1(8)
   SeekPrintL = Printer.Orientation

   FormClear
   'Add by Sindy 2021/6/28
   Call PUB_SetComboPCU11(cboPCU11, "", True) '設定國外潛在客戶類別選項
   
   AddCombo 1
   AddCombo 2
   PUB_AddCombo cboSort
End Sub

Private Function AddLstFrmCbo(oCombo As ComboBox, oList As ListBox) As Boolean
Dim idx As Integer, bFound As Boolean
   
   If oCombo <> "" Then
      For idx = 0 To oList.ListCount - 1
         If oList.List(idx) = oCombo Then
            MsgBox "資料已存在！"
            oCombo.SetFocus
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem oCombo, 0
         oCombo = ""
      End If
   End If
End Function

Private Function RemoveList(oList As ListBox) As Boolean
Dim idx As Integer, ii As Integer
   
   If oList.ListCount > 0 Then
      ii = 0
      For idx = 0 To oList.ListCount - 1
         If oList.Selected(ii) = True Then
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Function

Private Function ComposeList(oList As ListBox) As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      strExc(1) = oList.List(0)
      For intI = 1 To oList.ListCount - 1
         strExc(1) = strExc(1) & "," & oList.List(intI)
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Sub txtPCU_Change(Index As Integer)
   Select Case Index
      '國籍
      Case 1, 2
         
         If (Left(txtPCU(1), 3) <> Left(txtPCU(2), 3)) Then
             cboCity.Clear
          Else
            If Len(txtPCU(1)) >= 3 And Len(txtPCU(2)) >= 3 Then
                  SetCity
            End If
         End If
 
      '開發人員
      Case 4
         If txtPCU(4) <> "" Then
            Label9.Caption = GetStaffName(txtPCU(4))
            Label9.Visible = True
         Else
            Exit Sub
         End If
   End Select
End Sub

Private Sub SetCity()
   cboCity.Clear
   If txtPCU(1) <> "" And txtPCU(2) <> "" Then
      strExc(0) = "select ct03,ct05,ct02 from city where substr(ct01,1,3)>='" & Left(txtPCU(1), 3) & "' and substr(ct01,1,3)<= '" & Left(txtPCU(2), 3) & "' order by ct02 desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            cboCity.AddItem "" & .Fields("ct03").Value & " " & .Fields("ct02").Value
            .MoveNext
         Loop
         End With
      End If
   End If
   ShowCity
End Sub

Private Sub ShowCity()
   If cboCity.ListCount > 0 Then
      For intI = 0 To cboCity.ListCount - 1
         If cboCity.ItemData(intI) = Val(txtPCU(1)) & Val(txtPCU(2)) Then
            cboCity.ListIndex = intI
            Exit For
         End If
      Next
   End If
End Sub

Private Sub SetList(oList As ListBox, p_stData As String)
Dim arrID
   oList.Clear
   If p_stData <> "" Then
      arrID = Split(p_stData, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

Private Sub txtPCU_GotFocus(Index As Integer)
   Select Case Index
      Case 0 '洲別
         If txtPCU(0).Text <> "" Then
             txtPCU(0).SelStart = 0
             txtPCU(0).SelLength = Len(txtPCU(0))
         End If
      Case 1, 2 '國籍
         If txtPCU(1).Text <> "" And txtPCU(2) <> "" Then
             txtPCU(1).SelStart = 0
             txtPCU(1).SelLength = Len(txtPCU(1))
             txtPCU(2).SelStart = 0
             txtPCU(2).SelLength = Len(txtPCU(2))
         End If
'      '類別
'      Case 3
'         If txtPCU(3).Text <> "" Then
'            txtPCU(3).SelStart = 0
'            txtPCU(3).SelLength = Len(txtPCU(3))
'         End If
      '開發人員
      Case 4
         If txtPCU(4).Text <> "" Then
            txtPCU(4).SelStart = 0
            txtPCU(4).SelLength = Len(txtPCU(4))
         End If
      Case Else
         TextInverse Text1(Index)
   End Select
End Sub

Private Sub AddCombo(p_iID As Integer)
   Select Case p_iID
      '部門
      Case 1
         PUB_AddDeptCombo cboDept
      '職稱
      Case 2
         PUB_AddTitleCombo cboTitle
   End Select
End Sub

Private Sub PUB_AddCombo(oCombo As ComboBox)
   With oCombo
      .Clear
      .AddItem "IP諮詢"
      .AddItem "非IP法律諮詢"
      .AddItem "詢價"
      .AddItem "申請所需文件"
      .AddItem "利益衝突"
      .AddItem "互惠"
      .AddItem "詢問IP侵害"
      .AddItem "訪談" & Chr(1) & "(包括來所訪問、出國拜訪、國際會議)"
      .AddItem "客戶特別指示" & Chr(1) & "(譬如:不要寄confirmation copy, 聯絡方式只限fax or e-mail, 付款方式只限credit card or cheque…)"
   End With
End Sub

'Add By Sindy 2010/11/26
Private Sub txtPCU_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPCU_Validate(Index As Integer, Cancel As Boolean)
Dim strSql As String

   Select Case Index
      Case 1, 2 '國籍
         If txtPCU(1) <> "" And txtPCU(2) <> "" Then
            If Not nickChgRan(txtPCU(1), txtPCU(2), "國籍") Then
               txtPCU(2).SetFocus
               Exit Sub
            End If
            If txtPCU(1) = 台灣國家代號 Or txtPCU(2) = 台灣國家代號 Then
               Cancel = True
               ShowMsg MsgText(9153)
               TextInverse txtPCU(Index)
            Else
               ShowCity
            End If
         End If
   End Select
End Sub

Private Sub txtPCU05_GotFocus()
   txtPCU05.SelStart = 0
   txtPCU05.SelLength = Len(txtPCU05)
End Sub

Private Sub txtPCU05_LostFocus()
   If PUB_CheckKeyInDate(txtPCU05) = -1 Then
      txtPCU05.SetFocus
      Exit Sub
   End If
End Sub

Private Sub txtPCU06_GotFocus()
   txtPCU06.SelStart = 0
   txtPCU06.SelLength = Len(txtPCU06)
End Sub

Private Sub txtPCU06_LostFocus()
   If PUB_CheckKeyInDate(txtPCU06) = -1 Then
      txtPCU06.SetFocus
      Exit Sub
   End If
   
   If Not nickChgRan(txtPCU05, txtPCU06, "往來日期") Then
      txtPCU06.SetFocus
      Exit Sub
   End If
End Sub

Private Sub FormClear()
   txtPCU(0) = ""
   txtPCU(1) = ""
   txtPCU(2) = ""
'   txtPCU(3) = ""
   cboPCU11.Text = "" 'Add By Sindy 2021/6/28
   txtPCU(4) = ""
   txtPCU05 = ""
   txtPCU06 = ""
   txtPCC(6) = ""
   txtPCC(7) = ""
   Text2 = ""
   
   lstDept.Clear
   lstTitle.Clear
   cboCity.Clear
   lstSort.Clear
   cboDept = ""
   cboTitle = ""
   cboCity = ""
   lstSort = ""
   cboSort = ""
End Sub

Private Function Process() As Boolean
Dim StrPCC06 As Variant
Dim k_PCC06 As Integer        '檢查部門筆數
Dim StrPCC07 As Variant
Dim k_PCC07 As Integer        '檢查職稱筆數
Dim StrCR05 As Variant
Dim k_CR05 As Integer         '檢查往來類別筆數
Dim rsA As New ADODB.Recordset
'檢查客戶編號是否相同
Dim tm_str As String
'計算列印筆數
Dim Print_Count As Integer

   stSQLa = "": tm_str = ""
   Print_Count = 0
   
   '洲別
   If Len(txtPCU(0)) <> 0 Then
      Select Case txtPCU(0)
         Case "1"
            stSQLa = stSQLa & "(SUBSTR(NA02,1,1)='A' OR NA02='B00' OR NA02='C00') AND "
            pub_QL05 = pub_QL05 & ";" & Label1(0) & "1.亞洲" 'Add By Sindy 2010/12/24
         Case "2"
            stSQLa = stSQLa & "NA02='C10' AND "
            pub_QL05 = pub_QL05 & ";" & Label1(0) & "2.美洲" 'Add By Sindy 2010/12/24
         Case "3"
            stSQLa = stSQLa & "NA02='C20' AND "
            pub_QL05 = pub_QL05 & ";" & Label1(0) & "3.歐洲" 'Add By Sindy 2010/12/24
         Case "4"
            stSQLa = stSQLa & "NA02='C30' AND "
            pub_QL05 = pub_QL05 & ";" & Label1(0) & "4.非洲" 'Add By Sindy 2010/12/24
         Case "5"
            stSQLa = stSQLa & "NA02='C40' AND "
            pub_QL05 = pub_QL05 & ";" & Label1(0) & "5.大洋洲" 'Add By Sindy 2010/12/24
      End Select
   End If
   
   '國籍
   If Len(txtPCU(1)) <> 0 And Len(txtPCU(2)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txtPCU(1) & "-" & txtPCU(2) 'Add By Sindy 2010/12/24
      stSQLa = stSQLa & "substr(PCU09,1,3)>='" & txtPCU(1) & "' AND substr(PCU09,1,3) <='" & txtPCU(2) & "' AND "
   End If
   
   '類別
   'Modify By Sindy 2021/6/28
'   If Len(txtPCU(3)) <> 0 Then
'      Select Case txtPCU(3)
'         Case "1"
'            stSQLa = stSQLa & "PCU11='1' AND "
'            pub_QL05 = pub_QL05 & ";" & Label1(2) & "1.廠商" 'Add By Sindy 2010/12/24
'         Case "2"
'            stSQLa = stSQLa & "PCU11='2' AND "
'            pub_QL05 = pub_QL05 & ";" & Label1(2) & "2.事務所" 'Add By Sindy 2010/12/24
'      End Select
'   End If
   If Len(cboPCU11) <> 0 Then
      If Trim(Left(cboPCU11, 1)) <> "" Then
         stSQLa = stSQLa & "PCU11='" & Trim(Left(cboPCU11, 1)) & "' AND "
         pub_QL05 = pub_QL05 & ";" & Label1(2) & cboPCU11
      End If
   End If
   '2021/6/28 END
   
   '城市
   If cboCity <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & cboCity.Text 'Add By Sindy 2010/12/24
      stSQLa = stSQLa & "PCU10='" & Right(cboCity, 3) & "' AND "
   End If
   
   '開發人員
   If Len(txtPCU(4)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & txtPCU(4) & Label9 'Add By Sindy 2010/12/24
      stSQLa = stSQLa & "(instr(PCU38,'" & ChgSQL(txtPCU(4)) & "') > 0) AND "
   End If
   
   '往來日期
   If Len(txtPCU05) <> 0 And Len(txtPCU06) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & txtPCU05 & "-" & txtPCU06 'Add By Sindy 2010/12/24
      stSQLa = stSQLa & "CR02 >='" & ChangeTStringToWString(txtPCU05) & "' AND CR02 <='" & ChangeTStringToWString(txtPCU06) & "' AND "
   End If
   
   '部門
   StrPCC06 = Split(txtPCC(6), ",")
   If Trim(txtPCC(6)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(4) & Trim(txtPCC(6)) 'Add By Sindy 2010/12/24
   End If
   '職稱
   StrPCC07 = Split(txtPCC(7), ",")
   If Trim(txtPCC(7)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & Trim(txtPCC(7)) 'Add By Sindy 2010/12/24
   End If
   '往來類別
   StrCR05 = Split(Text2, ",")
   If Trim(Text2) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(8) & Trim(Text2) 'Add By Sindy 2010/12/24
   End If
   
   If txtPCU05 = "" And txtPCU06 = "" And Text2 = "" Then
      '2008/12/9 MODIFY BY SONIA
      'strExc(0) = "select DISTINCT(PCC01) as PCC,PCC02 ,PCC03,PCC04,PCC05,PCU36  from Potcustomer,PotCustcont,NATION where PCU09=na01(+) and " & stSQLa & " PCU01=PCC01(+) " & _
                   "AND  PCC02=01"
      'Modify by Amy 2018/04/11 若無聯絡人資料編號會空,後續會error 原:DISTINCT(PCC01)
      If txtPCC(6) = "" And txtPCC(7) = "" Then
        strExc(0) = "select DISTINCT(PCU01) as PCC,PCU36,PCC06,PCC07  from Potcustomer,PotCustcont,NATION where PCU09=na01(+) and " & stSQLa & " PCU01=PCC01(+) "
      Else
        strExc(0) = "select DISTINCT(PCU01)||PCC02 as PCC,PCC03,PCC04,PCC05,PCU36,PCC06,PCC07  from Potcustomer,PotCustcont,NATION where PCU09=na01(+) and " & stSQLa & " PCU01=PCC01(+) "
      End If
      'end 2018/04/11
      '2008/12/9 END
   Else
      '2008/12/9 MODIFY BY SONIA
      'strExc(0) = "select DISTINCT(PCC01) as PCC,PCC02 ,PCC03,PCC04,PCC05,PCU36,CR04,CR05 from Potcustomer,PotCustcont,ContactRecord,NATION,CITY where PCU09=na01(+) and " & stSQLa & " PCU01=PCC01(+) AND PCU01(+)=SUBSTR(CR03,1,8) " & _
                   "AND  PCC02=01"
      'Modify by Amy 2018/04/11 若無聯絡人資料編號會空,後續會error 原:DISTINCT(PCC01)
      If txtPCC(6) = "" And txtPCC(7) = "" Then
         strExc(0) = "select DISTINCT(PCU01) as PCC,PCU36,PCC06,PCC07,CR04,CR05 from Potcustomer,PotCustcont,ContactRecord,NATION,CITY where PCU09=na01(+) and " & stSQLa & " PCU01=PCC01(+) AND PCU01(+)=SUBSTR(CR03,1,8) "
      Else
         strExc(0) = "select DISTINCT(PCU01)||PCC02 as PCC,PCC03,PCC04,PCC05,PCU36,PCC06,PCC07,CR04,CR05 from Potcustomer,PotCustcont,ContactRecord,NATION,CITY where PCU09=na01(+) and " & stSQLa & " PCU01=PCC01(+) AND PCU01(+)=SUBSTR(CR03,1,8) "
      End If
      'end 2018/04/11
      '2008/12/9 END
   End If
   
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      While Not rsA.EOF
         Do While Not rsA.EOF
             '檢查客戶編號是否相同
            If rsA.Fields("PCC") = tm_str Then
               GoTo NextRecord
            Else
               '2008/12/9 ADD BY SONIA
               '檢查部門
               k_PCC06 = 0
               For i = 0 To UBound(StrPCC06)
                  If InStr(UCase(rsA.Fields("PCC06")), UCase(StrPCC06(i))) > 0 Then
                     k_PCC06 = k_PCC06 + 1
                  End If
               Next i
               If k_PCC06 < UBound(StrPCC06) + 1 Then
                  GoTo NextRecord
               End If
               '檢查職稱
               k_PCC07 = 0
               For i = 0 To UBound(StrPCC07)
                  If InStr(UCase(rsA.Fields("PCC07")), UCase(StrPCC07(i))) > 0 Then
                     k_PCC07 = k_PCC07 + 1
                  End If
               Next i
               If k_PCC07 < UBound(StrPCC07) + 1 Then
                  GoTo NextRecord
               End If
               '2008/12/9 END
               '檢查往來類別
               k_CR05 = 0
               For i = 0 To UBound(StrCR05)
                  If InStr(UCase(rsA.Fields("CR05")), UCase(StrCR05(i))) > 0 Then
                     k_CR05 = k_CR05 + 1
                  End If
               Next i
               If k_CR05 < UBound(StrCR05) + 1 Then
                  GoTo NextRecord
               End If
               
               '列印地址條
               Print_Count = Print_Count + 1
               Load frm083014
               frm083014.Hide
               frm083014.Opt1(3).Value = True
               
'               If txtPCC(6) = "" And txtPCC(7) = "" Then
'                  Debug.Print rsA.Fields("PCC")
'               Else
'                  Debug.Print rsA.Fields("PCC") & rsA.Fields("PCC03") & rsA.Fields("PCC04") & rsA.Fields("PCC05")
'               End If
'
               '客戶編號
               frm083014.Text1(16).Text = "" & rsA.Fields("PCC")
               tm_str = "" & rsA.Fields("PCC")
               '定稿語文
               '97/12/04 add by Toni
               '定稿語文為日文時改成英文
               If rsA.Fields("PCU36") = 3 Then
                  frm083014.Text1(4).Text = 2
               Else
                  frm083014.Text1(4).Text = rsA.Fields("PCU36")
               End If
               '2008/12/9 add by sonia
               If txtPCC(6) = "" And txtPCC(7) = "" Then
                  'Modified by Lydia 2024/05/07 變更欄位Text1(13)=>textFM2(0)
                  frm083014.textFM2(0).Text = ""
               Else
                  If frm083014.Text1(4).Text = 1 Then
                     'Modified by Lydia 2024/05/07 變更欄位Text1(13)=>textFM2(0)
                     frm083014.textFM2(0).Text = rsA.Fields("PCC05")
                  ElseIf frm083014.Text1(4).Text = 2 Then
                     'Modified by Lydia 2024/05/07 變更欄位Text1(13)=>textFM2(0)
                     frm083014.textFM2(0).Text = rsA.Fields("PCC03")
                  End If
               End If
               '2008/12/9 end
               frm083014.SetPrinter Printer.DeviceName
               frm083014.cmdPrint_Click
               Unload frm083014
         
            End If
NextRecord:
            rsA.MoveNext
            '----------------------
         Loop
            
         If rsA.EOF = False Then
            rsA.MoveNext
         End If
                     
      Wend
      
      InsertQueryLog (Print_Count) 'Add By Sindy 2010/12/24
      If Print_Count = 0 Then
         ShowNoData
         cboSort.SetFocus
         Exit Function
      Else
         MsgBox "列印結束，共 " & Format(Print_Count, "#,##0") & " 筆 !!!", vbInformation
         FormClear
         txtPCU(0).SetFocus
      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/24
      ShowNoData
      txtPCU(0).SetFocus
   End If
End Function

Private Function RemoveList1(oList As ListBox) As Boolean
Dim ii As Integer
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList1 = True
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function
