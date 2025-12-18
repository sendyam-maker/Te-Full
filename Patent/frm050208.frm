VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050208 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF代理人報價附件查詢"
   ClientHeight    =   5745
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   2700
      MaxLength       =   9
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   2250
      MaxLength       =   4
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1470
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   1
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1470
      MaxLength       =   7
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   360
      Left            =   5790
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "附件記錄(&N)"
      Height          =   360
      Left            =   6690
      TabIndex        =   7
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8010
      TabIndex        =   8
      Top             =   60
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   4455
      Left            =   60
      TabIndex        =   9
      Top             =   1230
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7858
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(民國年月日)"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   3
      Left            =   3840
      TabIndex        =   13
      Top             =   180
      Width           =   1020
   End
   Begin VB.Line Line3 
      X1              =   2460
      X2              =   2940
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   900
      Width           =   1080
   End
   Begin VB.Line Line2 
      X1              =   2010
      X2              =   2490
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人國籍："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   540
      Width           =   1080
   End
   Begin VB.Line Line1 
      X1              =   2460
      X2              =   2940
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   540
      TabIndex        =   10
      Top             =   180
      Width           =   900
   End
End
Attribute VB_Name = "frm050208"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; GRD1改字型=新細明體-ExtB
'Create by Sindy 2012/12/19
Option Explicit

'Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow As Double


'查詢明細資料
Private Sub cmdDetail_Click()
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         frm050208_1.Hide
         frm050208_1.cmdExit.Visible = True
         frm050208_1.m_CurrKEY1 = GRD1.TextMatrix(i, 2)
         frm050208_1.m_CurrKEY2 = Val(DBDATE(GRD1.TextMatrix(i, 3))) - 19110000
         frm050208_1.UpdateCtrlData
         frm050208_1.Show
         Me.Hide
         Exit For
      End If
   Next i
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdQuery_Click()
   Call QueryData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
   
'   m_blnColOrderAsc = True
   QueryData = False
   GRD1.Clear
   SetGrd
   
   strCon = ""
   '日期
   If Trim(txt1(0)) <> "" And Trim(txt1(1)) <> "" Then
      strCon = strCon & " and cq02>=" & DBDATE(txt1(0)) & " and cq02<=" & DBDATE(txt1(1))
'   Else
'      MsgBox "日期不可空白！", vbExclamation
'      If Trim(txt1(0)) = "" Then txt1(0).SetFocus: Exit Function
'      If Trim(txt1(1)) = "" Then txt1(1).SetFocus: Exit Function
   End If
   '代理人國籍
   If Trim(txt1(2)) <> "" And Trim(txt1(3)) <> "" Then
      strCon = strCon & " and fa10>='" & txt1(2) & "' and fa10<='" & txt1(3) & "'"
   End If
   '代理人編號
   If Trim(txt1(4)) <> "" And Trim(txt1(5)) <> "" Then
      strCon = strCon & " and cq01>='" & Left(txt1(4) & "00000000", 9) & "' and cq01<='" & Left(txt1(5) & "00000000", 9) & "'"
   End If
   
   Screen.MousePointer = vbHourglass
   strSql = "select ' ' V,na03 代理人國籍,cq01 代理人編號,sqldatet(cq02) 日期,cq03 內容,counting(cq04) 附件個數 from cfquotation,fagent,nation " & _
            "where substr(cq01,1,8)=fa01(+) and substr(cq01,9,1)=fa02(+) " & _
            "and fa10=na01(+) " & strCon & _
            " order by 2,3,4"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryData = True
      Set GRD1.Recordset = rsTmp
   Else
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      ShowNoData
      Exit Function
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   dblPrevRow = GRD1.row
   If rsTmp.RecordCount > 0 Then
      GRD1.Text = "V"
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050208 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "代理人國籍", "代理人編號", "日期", "內容", "附件個數")
   arrGridHeadWidth = Array(200, 1000, 1000, 800, 4300, 1000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 0
      GRD1.row = dblPrevRow
      GRD1.Text = ""
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
   If GRD1.Text = "V" Then
      GRD1.Text = ""
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   Else
      If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
'   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
'      If Me.GRD1.Text = "表單編號" Or Me.GRD1.Text = "天數" Or Me.GRD1.Text = "時數" Then
'         If m_blnColOrderAsc = True Then
'            Me.GRD1.Sort = 3  '數值昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD1.Sort = 4 '數值降冪
'            m_blnColOrderAsc = True
'         End If
'      Else
'         If m_blnColOrderAsc = True Then
'            Me.GRD1.Sort = 5 '字串昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD1.Sort = 6 '字串降冪
'            m_blnColOrderAsc = True
'         End If
'      End If
'   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = Pub_NumAscii(KeyAscii, True)
      Case 4, 5
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If IsEmptyText(txt1(Index)) = True Then Exit Sub
   
   Cancel = False
   Select Case Index
      Case 0, 1
         If CheckIsTaiwanDate(txt1(Index), False) = False Then
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Call txt1_GotFocus(Index)
            Exit Sub
         End If
         If Trim(txt1(0)) <> "" And Trim(txt1(1)) = "" Then txt1(1).Text = Trim(txt1(0).Text)
      Case 2, 3
         If Trim(txt1(2)) <> "" And Trim(txt1(3)) = "" Then txt1(3).Text = Trim(txt1(2).Text)
      Case 4, 5
         If Left(txt1(Index), 1) <> "Y" Then
            Cancel = True
            MsgBox "必須輸入代理人編號！", vbInformation, "檢核資料"
            Call txt1_GotFocus(Index)
            Exit Sub
         End If
         If Trim(txt1(4)) <> "" And Trim(txt1(5)) = "" Then txt1(5).Text = Trim(txt1(4).Text)
   End Select
End Sub
