VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040114_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "非台灣案發文後暫緩"
   ClientHeight    =   5745
   ClientLeft      =   -2670
   ClientTop       =   1575
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8388
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7530
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   120
      TabIndex        =   13
      Top             =   528
      Width           =   9072
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "收文號"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "P"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   6624
         TabIndex        =   7
         Top             =   180
         Width           =   800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1740
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      TabIndex        =   10
      Top             =   1230
      Width           =   7995
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14102;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   9180
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   9180
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   1230
      Width           =   765
   End
End
Attribute VB_Name = "frm040114_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (Combo1)
'Create By Sindy 2014/7/21 參考:內專發文(frm040104_1)改寫
Option Explicit

Dim intLastRow As Integer
Dim m_bolActivated As Boolean


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 1 '確定
         strKey5 = MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 2)
         strExc(0) = "select cp47 from caseprogress where cp09='" & strKey5 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Trim("" & RsTemp.Fields("cp47")) <> "" Then
               MsgBox "此程序代理人已提申不可暫緩！", vbExclamation
               Exit Sub
            End If
         End If
         'Added by Morgan 2021/12/22
         '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm040114_2") = False Then
            Set frm040114_2 = Nothing
         End If
         'end 2021/12/22
         frm040114_2.Show
         Command1.SetFocus
         Me.Hide
      Case 2 '離開
         Unload Me
   End Select
End Sub

Public Sub Command1_Click()
Dim ii As Integer
   
   '選擇本所案號
   If Option1(1).Value = True Then
      strExc(0) = "select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & Text5 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text1 = "" & RsTemp.Fields("cp01")
         Text2 = "" & RsTemp.Fields("cp02")
         Text3 = "" & RsTemp.Fields("cp03")
         Text4 = "" & RsTemp.Fields("cp04")
      Else
         MsgBox "無該收文號資料！", vbExclamation
         Exit Sub
      End If
   End If
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    If FMP2open = True Then
      If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1, Text2, Text3, Text4) = False Then Exit Sub
    End If
        
   strKey1 = Text1
   StrKey2 = Text2
   strKey3 = Text3
   strKey4 = Text4
   Combo1.Clear
   If Text1 = "P" Then
      strExc(0) = "SELECT PA05,PA06,PA07,PA09 FROM PATENT WHERE " & ChgPatent(strKey1 & StrKey2 & strKey3 & strKey4)
   ElseIf Text1 = "PS" Then
      strExc(0) = "SELECT SP05,SP06,SP07,SP09 FROM SERVICEPRACTICE WHERE " & ChgService(strKey1 & StrKey2 & strKey3 & strKey4)
   End If
   intI = 1
   strExc(1) = "CPM03,"
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         Combo1.AddItem "中 : " & .Fields(0)
         Combo1.AddItem "英 : " & .Fields(1)
         Combo1.AddItem "日 : " & .Fields(2)
         strKey6 = "" & .Fields(0)
         strKey7 = "" & .Fields(1)
         strKey8 = "" & .Fields(2)
         Combo1.ListIndex = 0
         If IsNull(.Fields(3)) = False Then
            If .Fields(3) = 台灣國家代號 Then
               strExc(1) = "CPM03,"
               MsgBox "不可為台灣案！", vbExclamation
               Exit Sub
            Else
               strExc(1) = "CPM04,"
            End If
         End If
      End With
   End If
   strExc(0) = "select ''," & SQLDate("CP05") & ",cp09," & strExc(1) & "staff.st02 as st1," & _
      "staff1.st02 as st2,Decode(CP06,Null,Null,CP06 - 19110000),DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),cp64,cp10,cp12,cp13,CP79,CP43 from caseprogress, casepropertymap," & _
      "staff,staff staff1,Customer where " & ChgCaseprogress(strKey1 & StrKey2 & strKey3 & strKey4) & _
      " AND CP27 IS not NULL AND CP57 IS NULL" & _
      " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) "
   If Option1(1).Value = True Then strExc(0) = strExc(0) & " and cp09='" & Text5 & "'"
   strExc(0) = strExc(0) & " order by cp27 desc,cp05 desc,cp09 desc"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set MSHFlexGrid1.Recordset = RsTemp
   '相關總收號案件性質
   For ii = 1 To Me.MSHFlexGrid1.Rows - 1
      Me.MSHFlexGrid1.TextMatrix(ii, 3) = Me.MSHFlexGrid1.TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(Me.MSHFlexGrid1.TextMatrix(ii, 2), "1")
   Next ii
   GridHead
   
   '若只搜尋到一筆時直接勾選
   If Me.MSHFlexGrid1.Rows = 2 Then
      Me.MSHFlexGrid1.row = 1
      GridClick MSHFlexGrid1, 1, 0
      strKey5 = MSHFlexGrid1.TextMatrix(1, 2)
      cmdOK(1).SetFocus
      If Option1(1).Value = True Then cmdOK_Click 1
   End If
End Sub

Private Sub Form_Activate()
Dim i As Integer, j As Integer
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
        If .CellBackColor = &HFFC0C0 Then
            For j = 0 To .Cols - 1
               .col = j
               .CellBackColor = .BackColor
            Next
         End If
      Next
   End With
   '預設本所號
   If Not m_bolActivated Then
      m_bolActivated = True
      Option1(0).Value = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'Combo1.ListIndex = 0 'Removed by Morgan 2021/12/22
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   Text4.Enabled = False
   InitGrid 14, MSHFlexGrid1
   GridHead
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040114_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(1).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
   Select Case Index
      Case 0
         Text1.Enabled = True
         Text2.Enabled = True
         Text3.Enabled = True
         Text4.Enabled = True
         Text5.Enabled = False
         Me.Text2.SetFocus
      Case 1
         Text1.Enabled = False
         Text2.Enabled = False
         Text3.Enabled = False
         Text4.Enabled = False
         Text5.Enabled = True
         Text5.SetFocus
   End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "P" And Text1 <> "PS" And Text1 <> "" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 900: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 900: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 900: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 900: .Text = "本所期限"
      .col = 7: .ColWidth(7) = 1400: .Text = "相關人"
      .col = 8: .ColWidth(8) = 1400: .Text = "進度備註"
      For i = 9 To 13
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub ReQuery()
    If Me.Option1(1).Value Then
        Me.Option1(0).Value = True
    End If
    Command1_Click
End Sub

Public Sub Clear()
   '保留系統類別
'   Text1 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   Option1(0).Value = False
   '預設本所案號
   Option1(0).Value = True
   Combo1.Clear
   InitGrid 14, MSHFlexGrid1
   GridHead
   strKey1 = ""
   StrKey2 = ""
   strKey3 = ""
   strKey4 = ""
   strKey5 = ""
   strKey6 = ""
   strKey7 = ""
   strKey8 = ""
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub
