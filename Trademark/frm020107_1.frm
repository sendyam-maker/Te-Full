VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020107_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "T案大陸指示信"
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
      Left            =   7560
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   120
      TabIndex        =   12
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
         Left            =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "T"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   6624
         TabIndex        =   5
         Top             =   180
         Width           =   800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   11
      Top             =   1770
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   12
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1170
      TabIndex        =   13
      Top             =   1260
      Width           =   7995
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14097;503"
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
      Y1              =   1870
      Y2              =   1870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   9180
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標名稱："
      Height          =   180
      Left            =   210
      TabIndex        =   10
      Top             =   1320
      Width           =   900
   End
End
Attribute VB_Name = "frm020107_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/06 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Combo1
'Create By Sindy 2014/4/29
Option Explicit

Dim tm(0 To 10) As String
Dim intWhere As Integer
Dim intLastRow As Integer

Public m_Nation As String


Private Sub cmdOK_Click(Index As Integer)
   Dim i As Integer, bolChk As Boolean
   Dim strCP09 As String
   
   Select Case Index
      Case 1 '確定
         If Option1(0).Value = True Then '本所案號
            With MSHFlexGrid1
               For i = 1 To .Rows - 1
                  If .TextMatrix(i, 0) = "v" Then
                     bolChk = True
                     Me.Tag = .TextMatrix(i, 2)
                     '記錄總收文號
                     strCP09 = .TextMatrix(i, 2)
                     tm(10) = .TextMatrix(i, 7) '案件性質
                     Exit For
                  End If
               Next
            End With
            If bolChk = False Then
               MsgBox "請選擇資料 !", vbInformation
               Exit Sub
            End If
            frm020107_2.Caption = Me.Caption
            frm020107_2.Show
            cmdQuery.SetFocus
            Me.Hide
         Else
            cmdQuery_Click '收文號
         End If
      Case 2 '離開
         Unload Me
   End Select
End Sub

Public Sub cmdQuery_Click()
 Dim i As Integer
   '選擇本所案號
   If Option1(0).Value = True Then
      If Me.Text2.Text = "" Then
         MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
         Me.Text2.SetFocus
         Text2_GotFocus
         Exit Sub
      End If
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
      tm(1) = Text1
      tm(2) = Text2
      tm(3) = Text3
      tm(4) = Text4
      Combo1.Clear
      strExc(0) = "SELECT TM05,TM06,TM07,TM10,TM28,TM14 FROM Trademark WHERE " & ChgTradeMark(tm(1) & tm(2) & tm(3) & tm(4))
      intI = 1
      strExc(1) = "CPM03,"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
            Combo1.AddItem .Fields(0)
            Combo1.ListIndex = 0
            If IsNull(.Fields(3)) = False Then
               m_Nation = .Fields(3)
               If .Fields(3) = 台灣國家代號 Then
                  strExc(1) = "CPM03,"
               Else
                  strExc(1) = "CPM04,"
               End If
            End If
         End With
      End If
      
'      strExc(0) = "Select ''," & SQLDate("CP05") & ",cp09," & strExc(1) & "staff.st02 as st1," & _
'                  "staff1.st02 as st2,cp64,cp10,cp12,cp13,CP73" & _
'                  " From CaseProgress, CasePropertyMap,Staff,Staff Staff1,Trademark" & _
'                  " Where " & ChgCaseprogress(tm(1) & tm(2) & tm(3) & tm(4)) & _
'                  " AND CP27 IS NULL AND CP57 IS NULL AND CP09<'C' AND TM10='020'" & _
'                  " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)" & _
'                  " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)"
      strExc(0) = "Select ''," & SQLDate("CP05") & ",cp09," & strExc(1) & "staff.st02 as st1," & _
                  "staff1.st02 as st2,cp64,cp10,cp12,cp13,CP73" & _
                  " From CaseProgress, CasePropertyMap,Staff,Staff Staff1,Trademark" & _
                  " Where " & ChgCaseprogress(tm(1) & tm(2) & tm(3) & tm(4)) & _
                  " AND CP57 IS NULL AND CP09<'C' AND TM10='020'" & _
                  " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)" & _
                  " AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)"
      intI = 0
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
      '若只搜尋到一筆時直接勾選
      If Me.MSHFlexGrid1.Rows = 2 Then
         Me.MSHFlexGrid1.row = 1
         MSHFlexGrid1_Click
         cmdOK_Click (1)
      End If
   '選擇收文號
   Else
      If Me.Text5.Text = "" Then
         MsgBox "請輸入收文號!!!", vbExclamation + vbOKOnly
         Me.Text5.SetFocus
         Text5_GotFocus
         Exit Sub
      End If
      Me.Combo1.Clear
      GridHead
      
      strExc(0) = "Select CP10,CP12,CP13,CP73,TM10,CP01,CP02,CP03,CP04,TM14,CP05,CP06,CP07,TM05,TM06,TM07" & _
                  " from CaseProgress,Trademark" & _
                  " Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
                  " AND TM01='T' AND CP27 IS NULL AND CP57 IS NULL AND CP09<'C' AND TM10='020'" & _
                  " AND CP09='" & Me.Text5.Text & "'"
'      strExc(0) = "Select CP10,CP12,CP13,CP73,TM10,CP01,CP02,CP03,CP04,TM14,CP05,CP06,CP07,TM05,TM06,TM07" & _
'                  " from CaseProgress,Trademark" & _
'                  " Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)" & _
'                  " AND TM01='T' AND CP57 IS NULL AND CP09<'C' AND TM10='020'" & _
'                  " AND CP09='" & Me.Text5.Text & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
            Me.Tag = Text5
            If Not IsNull(.Fields(5)) Then Text1.Text = .Fields(5)
            If Not IsNull(.Fields(6)) Then Text2.Text = .Fields(6)
            If Not IsNull(.Fields(7)) Then Text3.Text = .Fields(7)
            If Not IsNull(.Fields(8)) Then Text4.Text = .Fields(8)
            m_Nation = ""
            If Not IsNull(.Fields(4)) Then m_Nation = .Fields(4)
            Combo1.AddItem .Fields("TM05").Value
            Combo1.ListIndex = 0
            frm020107_2.Caption = Me.Caption
            frm020107_2.Show
            Me.Hide
         End With
      Else
         MsgBox "無符合發文條件之資料 !", vbCritical
      End If
   End If
End Sub

Private Sub Form_Activate()
   Dim i As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
      Next
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   Text4.Enabled = False
   InitGrid 11, MSHFlexGrid1
   GridHead
   '改預設本所案號
   Option1(0).Value = True
   Option1_Click 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020107_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(1).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0 '選擇本所案號
         Text1.Enabled = False
         Text2.Enabled = True
         Text3.Enabled = True
         Text4.Enabled = True
         Text5.Enabled = False
         If Me.Visible = True Then
            Text2.SetFocus
         End If
      Case 1 '選擇收文號
         Text1.Enabled = False
         Text2.Enabled = False
         Text3.Enabled = False
         Text4.Enabled = False
         Text5.Enabled = True
         If Me.Visible = True Then
            Text5.SetFocus
         End If
   End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "T" And Text1 <> "" Then
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
      .col = 1: .ColWidth(1) = 1200: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "案件備註"
      For i = 7 To 10
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub ReQuery()
   cmdQuery_Click
End Sub

Public Sub Clear()
   Text1 = "T"
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   If Option1(0).Value = True Then
       Text2.SetFocus
   Else
      Text5.SetFocus
   End If
   Combo1.Clear
   InitGrid 11, MSHFlexGrid1
   GridHead
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
