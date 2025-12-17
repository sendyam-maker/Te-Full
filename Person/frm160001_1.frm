VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160001_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "健保異動資料維護"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9165
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Frame1 
      Caption         =   "眷屬資料"
      Height          =   1695
      Left            =   90
      TabIndex        =   25
      Top             =   480
      Width           =   8970
      Begin VB.TextBox textHL05 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2535
         TabIndex        =   20
         Top             =   1275
         Width           =   4110
      End
      Begin VB.TextBox textSR05 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5865
         MaxLength       =   7
         TabIndex        =   12
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox textSR03 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1035
         MaxLength       =   7
         TabIndex        =   10
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox textSR06 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1035
         MaxLength       =   7
         TabIndex        =   13
         Top             =   585
         Width           =   1095
      End
      Begin VB.TextBox textSR07 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3630
         MaxLength       =   10
         TabIndex        =   14
         Top             =   585
         Width           =   1485
      End
      Begin VB.CheckBox chkSR08 
         Caption         =   "健保眷屬"
         Enabled         =   0   'False
         Height          =   285
         Left            =   45
         TabIndex        =   19
         Top             =   1275
         Width           =   1035
      End
      Begin VB.TextBox textSR09 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5865
         MaxLength       =   20
         TabIndex        =   15
         Top             =   585
         Width           =   1440
      End
      Begin VB.TextBox textSR10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   17
         Top             =   900
         Width           =   1095
      End
      Begin VB.CheckBox chkSR13 
         Caption         =   "歿"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7785
         TabIndex        =   16
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox textSR12 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7755
         MaxLength       =   7
         TabIndex        =   21
         Top             =   1275
         Width           =   1095
      End
      Begin MSForms.TextBox textSR11 
         Height          =   285
         Left            =   2850
         TabIndex        =   18
         Top             =   900
         Width           =   6015
         VariousPropertyBits=   679495707
         MaxLength       =   70
         Size            =   "10610;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSR04 
         Height          =   285
         Left            =   3630
         TabIndex        =   11
         Top             =   285
         Width           =   1485
         VariousPropertyBits=   679495707
         MaxLength       =   12
         Size            =   "2619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "健保補助類別："
         Height          =   180
         Left            =   1215
         TabIndex        =   35
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "稱謂："
         Height          =   180
         Left            =   75
         TabIndex        =   34
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   180
         Left            =   3015
         TabIndex        =   33
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "性別："
         Height          =   180
         Left            =   5295
         TabIndex        =   32
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "出生日期："
         Height          =   180
         Left            =   75
         TabIndex        =   31
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "身分證字號："
         Height          =   180
         Left            =   2475
         TabIndex        =   30
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Left            =   5295
         TabIndex        =   29
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "郵遞區號："
         Height          =   180
         Left            =   75
         TabIndex        =   28
         Top             =   945
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Left            =   2295
         TabIndex        =   27
         Top             =   945
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "刪除日期："
         Height          =   180
         Left            =   6795
         TabIndex        =   26
         Top             =   1320
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Left            =   7875
      TabIndex        =   24
      Top             =   30
      Width           =   1125
   End
   Begin VB.ComboBox cboHL05 
      Height          =   300
      ItemData        =   "frm160001_1.frx":0000
      Left            =   4050
      List            =   "frm160001_1.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   3000
      Width           =   4005
   End
   Begin VB.ComboBox cboHL04 
      Height          =   300
      ItemData        =   "frm160001_1.frx":0004
      Left            =   1125
      List            =   "frm160001_1.frx":0006
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   3030
      Width           =   1125
   End
   Begin VB.TextBox txtHL03 
      Height          =   285
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   3450
      TabIndex        =   4
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2625
      TabIndex        =   3
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "新增"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "刪除"
      Enabled         =   0   'False
      Height          =   345
      Index           =   3
      Left            =   1785
      TabIndex        =   1
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "修改"
      Enabled         =   0   'False
      Height          =   345
      Index           =   4
      Left            =   945
      TabIndex        =   0
      Top             =   2280
      Width           =   795
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1875
      Left            =   135
      TabIndex        =   9
      Top             =   3420
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3307
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      FixedCols       =   0
      ForeColorSel    =   16777215
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "異動日期|異動原因|補助類別"
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
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "健保補助類別："
      Height          =   180
      Left            =   2745
      TabIndex        =   23
      Top             =   3060
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "異動原因："
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   3060
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "異動日期："
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   2730
      Width           =   900
   End
End
Attribute VB_Name = "frm160001_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/15 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/6/10
Option Explicit

Public strHL01 As String
Public strHL02 As String

Dim m_EditMode  As Integer
Dim iLstSelRow As Integer

Private Sub cboHL04_Click()
   If m_EditMode <> 0 Then
      If cboHL04.ListIndex = 1 Then
         cboHL05.ListIndex = 0
         cboHL05.Enabled = False
      Else
         cboHL05.Enabled = True
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim i As Integer, iOldValue As Integer
   Dim bCheck As Boolean
   
   '要控制已設定為歿不可再維護否則若又新增異動且為健保眷屬時會矛盾！
   If chkSR13.Value Then
      MsgBox "【" & textSR04 & "】已設定為歿，不可再有異動！"
      Exit Sub
   End If
   
   bCheck = False
   
   Select Case Index
      Case 0 '新增
         iLstSelRow = -1
         For i = 1 To grd1.Rows - 1
            grd1.row = i
            grd1.col = 0
            If grd1.CellBackColor = &HFFC0C0 Then
               iLstSelRow = i
               Exit For
            End If
         Next
         ClearData
         grd1.Rows = grd1.Rows + 1
         grd1.row = grd1.Rows - 1
         grd1_SelChange
         
         EnableData True
         cmdOK(0).Enabled = False
         cmdOK(1).Enabled = True
         cmdOK(2).Enabled = True
         cmdOK(3).Enabled = False
         cmdOK(4).Enabled = False
         grd1.Enabled = False
         txtHL03.SetFocus
         m_EditMode = 1
         SetcboHL04
                     
      Case 1 '確定
         For i = 1 To grd1.Rows - 1
            grd1.row = i
            grd1.col = 0
            If grd1.CellBackColor = &HFFC0C0 Then
               If TxtValidate = True Then
                  iOldValue = chkSR08.Value
                  If grd1.TextMatrix(i, 0) = "" Then
                     If InsertRec() = False Then
                        Exit Sub
                     End If
                  Else
                     If UpdateRec(i) = False Then
                        Exit Sub
                     End If
                  End If
                  EnableData False
                  cmdOK(0).Enabled = True
                  cmdOK(1).Enabled = False
                  cmdOK(2).Enabled = False
                  cmdOK(3).Enabled = True
                  cmdOK(4).Enabled = True
                  grd1.Enabled = True
                  m_EditMode = 0
                  SetcboHL04
                  
                  ClearData
                  SetGrid
                  If iOldValue <> chkSR08.Value Then
                     MsgBox "健保眷屬設定已自動變更！"
                  End If
               End If
               Exit For
            End If
         Next
                  
      Case 2 '取消
         SetGrid
         EnableData False
         cmdOK(0).Enabled = True
         cmdOK(1).Enabled = False
         cmdOK(2).Enabled = False
         cmdOK(3).Enabled = True
         cmdOK(4).Enabled = True
         grd1.Enabled = True
         
         m_EditMode = 0
         SetcboHL04
         
         If iLstSelRow >= 0 And iLstSelRow < grd1.Rows Then
            grd1.row = iLstSelRow
            grd1.col = 0
            grd1.CellBackColor = QBColor(15) '顏色還原以便重新讀取資料
            grd1_SelChange
         Else
            ClearData
         End If
         
      Case 3 '刪除
         If txtHL03 = "" Then
            MsgBox "請點選異動資料！"
         Else
            strExc(1) = Left(cboHL04.Text, 1)
            If Pub_StrUserSt03 = "M21" And strExc(1) = "3" Then
               MsgBox "不可刪除異動原因為【" & cboHL04 & "】的資料！"
               Exit Sub
            ElseIf Pub_StrUserSt03 = "M31" And strExc(1) <> "3" Then
               MsgBox "不可刪除異動原因為【" & cboHL04 & "】的資料！"
               Exit Sub
            End If
            
            For i = 1 To grd1.Rows - 1
               grd1.row = i
               grd1.col = 0
               If grd1.CellBackColor = &HFFC0C0 Then
                  If MsgBox("是否確定要刪除！", vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  Else
                     iOldValue = chkSR08.Value
                     If DeleteRec = False Then
                        Exit Sub
                     End If
                     bCheck = True
                     m_EditMode = 0
                     SetcboHL04
                     
                     ClearData
                     SetGrid
                     If iOldValue <> chkSR08.Value Then
                        MsgBox "健保眷屬設定已自動變更！"
                     End If
                     Exit For
                  End If
               End If
            Next i
         End If
         
      Case 4 '修改
         If txtHL03 = "" Then
            MsgBox "請點選異動資料！"
         Else
            strExc(1) = Left(cboHL04.Text, 1)
            If Pub_StrUserSt03 = "M21" And strExc(1) = "3" Then
               MsgBox "不可修改異動原因為【" & cboHL04 & "】的資料！"
               Exit Sub
            ElseIf Pub_StrUserSt03 = "M31" And strExc(1) <> "3" Then
               MsgBox "不可修改異動原因為【" & cboHL04 & "】的資料！"
               Exit Sub
            End If
            For i = 1 To grd1.Rows - 1
               grd1.row = i
               grd1.col = 0
               If grd1.CellBackColor = &HFFC0C0 Then
                  iLstSelRow = i
                  EnableData True
                  txtHL03.Enabled = False '異動日期不可改
                  cboHL04.SetFocus
                  cmdOK(0).Enabled = False
                  cmdOK(1).Enabled = True
                  cmdOK(2).Enabled = True
                  cmdOK(3).Enabled = False
                  cmdOK(4).Enabled = False
                  grd1.Enabled = False
                  m_EditMode = 2
                  SetcboHL04 Left(cboHL04.Text, 1)
                  Exit For
               End If
            Next
         End If
      Case Else
   End Select
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_EditMode = 0
   SetcboHL04
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160001_1 = Nothing
End Sub

Public Sub InitForm()
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer

   SetcboHL05
   
   stSQL = "SELECT sr03||' '||decode(sr03,'1','父親','2','母親','3','配偶','4','子女','其他')" & _
      ",Sr04,DECODE(SR12,NULL,NULL,'刪') STATUS,sr05||' '||decode(sr05,'M','男','F','女','不詳')" & _
      ",sqldatet(sr06),sr07,sr08,sr13,sr09,sr10,sr11,sqldatet(sr12),sr02,hl05" & _
      " FROM staff_relation,(select hl02,hl05 from HIrelationlog a where hl01='" & strHL01 & "' and hl02=" & strHL02 & _
      " and hl03=(select max(b.hl03) from HIrelationlog b where b.hl01=a.hl01 and b.hl02=a.hl02) )X" & _
      " WHERE SR01='" & strHL01 & "' and SR02=" & strHL02 & " and hl02(+)=sr02"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoRst
      textSR03.Text = "" & .Fields(0)
      textSR04.Text = "" & .Fields(1)
      textSR05.Text = "" & .Fields(3)
      textSR06.Text = ChangeTDateStringToTString("" & .Fields(4))
      textSR07.Text = "" & .Fields(5)
      chkSR08.Value = IIf("" & .Fields(6) = "Y", vbChecked, vbUnchecked)
      chkSR13.Value = IIf("" & .Fields(7) = "Y", vbChecked, vbUnchecked)
      textSR09.Text = "" & .Fields(8)
      textSR10.Text = "" & .Fields(9)
      textSR11.Text = "" & .Fields(10)
      textSR12.Text = ChangeTDateStringToTString("" & .Fields(11))
      setHL05 "" & .Fields(13)
      End With
      SetGrid
   End If
   
   EnableCmdBtn 0
   EnableData False
End Sub

Private Sub setHL05(pValue As String)
   If pValue = "" Then
      textHL05 = cboHL05.List(0)
   Else
      For intI = 1 To cboHL05.ListCount - 1
         If Left(cboHL05.List(intI), 2) = pValue Then
            textHL05 = cboHL05.List(intI)
            Exit For
         End If
      Next
   End If
End Sub

Private Sub SetGrid()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   
   strSql = "select sqldatet(HL03) X1,decode(HL04,'1','轉入','2','轉出','3','調整',HL04) X2" & _
      ",HR04,HL04,HL05 from HiRelationLog,HiReduce where HL01='" & strHL01 & "' and HL02=" & strHL02 & _
      " and HR01(+)=HL05 order by HL03 desc"
      
   intI = 1
   Set grd1.Recordset = ClsLawReadRstMsg(intI, strSql)
      
   arrGridHeadText = Array("異動日期", "異動原因", "補助類別", "HL04", "HL05")
   arrGridHeadWidth = Array(850, 850, 6900, 0, 0)
   grd1.Visible = False
   grd1.Cols = UBound(arrGridHeadText) + 1
   grd1.row = 0
   For iCol = 0 To grd1.Cols - 1
      grd1.col = iCol
      grd1.Text = arrGridHeadText(iCol)
      grd1.ColWidth(iCol) = arrGridHeadWidth(iCol)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
   grd1.Visible = True
End Sub

Private Sub EnableCmdBtn(ByVal iState As Integer)
   Select Case iState
      Case 1, 2 '新增,修改
         cmdOK(0).Enabled = False
         cmdOK(1).Enabled = True
         cmdOK(2).Enabled = True
         cmdOK(3).Enabled = False
         cmdOK(4).Enabled = False
      Case 9 '查詢
         cmdOK(0).Enabled = False
         cmdOK(1).Enabled = False
         cmdOK(2).Enabled = False
         cmdOK(3).Enabled = False
         cmdOK(4).Enabled = False
      Case Else
         cmdOK(0).Enabled = True
         cmdOK(1).Enabled = False
         cmdOK(2).Enabled = False
         cmdOK(3).Enabled = True
         cmdOK(4).Enabled = True
   End Select
   
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   grd1_SelChange
End Sub

Private Sub grd1_SelChange()
   Dim tmpMouseRow
   Dim i, j
   
   'tmpMouseRow = grd1.row
   If grd1.MouseRow <> 0 Then
      tmpMouseRow = grd1.MouseRow
   Else
      tmpMouseRow = grd1.row
   End If
   grd1.Visible = True
   If tmpMouseRow <> 0 Then
       grd1.row = tmpMouseRow
       grd1.col = 0
       If grd1.CellBackColor = QBColor(15) Then
             grd1.Visible = False
             For j = 1 To grd1.Rows - 1
                 grd1.row = j
                 For i = 0 To grd1.Cols - 1
                      grd1.col = i
                      grd1.CellBackColor = QBColor(15)
                 Next i
            Next j
            grd1.row = tmpMouseRow
            For i = 0 To grd1.Cols - 1
                grd1.col = i
                grd1.CellBackColor = &HFFC0C0
            Next i
            txtHL03 = ChangeTDateStringToTString(grd1.TextMatrix(tmpMouseRow, 0))
            SelCombo cboHL04, grd1.TextMatrix(tmpMouseRow, 3), 1, -1
            SelCombo cboHL05, grd1.TextMatrix(tmpMouseRow, 4)
            grd1.Visible = True
       End If
   End If
End Sub

Private Sub SetcboHL05()
   cboHL05.Clear
   cboHL05.AddItem "無"
   strSql = "select HR01||' '||HR04 from HiReduce order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         cboHL05.AddItem "" & .Fields(0).Value
         .MoveNext
      Loop
      End With
   End If
End Sub

Sub ClearData()
   txtHL03 = Empty
   cboHL04.ListIndex = -1
   cboHL05.ListIndex = 0
End Sub

Private Sub EnableData(ByVal bEnable As Boolean)
   txtHL03.Enabled = bEnable
   cboHL04.Enabled = bEnable
   cboHL05.Enabled = bEnable
End Sub

Private Function UpdateRec(ByVal iRow As String) As Boolean
   Dim strHL03 As String, strHL04 As String, strHL05 As String, strOldHL03 As String
   
   strOldHL03 = DBDATE(grd1.TextMatrix(iRow, 0))
   strHL03 = DBDATE(txtHL03)
   strHL04 = Left(cboHL04, 1)
   If cboHL05.ListIndex > 0 Then
      strHL05 = Left(cboHL05, 2)
   Else
      strHL05 = ""
   End If
   
   If CheckLogLogic(strHL01, strHL02, strHL03, strHL04) = False Then
      Exit Function
   End If
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   strSql = "Update HiRelationLog set HL03=" & strHL03 & ",HL04='" & strHL04 & "',HL05='" & strHL05 & "'" & _
      " where HL01='" & strHL01 & "' and HL02=" & strHL02 & " and HL03=" & strOldHL03
   strSql = "begin user_data.user_enabled:=1; " & strSql & "; end;"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   UpdateSR08
   
   cnnConnection.CommitTrans
   UpdateRec = True
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      If Err.Number = -2147217873 Then
         MsgBox "異動紀錄重複！"
      Else
         MsgBox Err.Description
      End If
   End If
End Function

Private Function InsertRec() As Boolean
   Dim strHL03 As String, strHL04 As String, strHL05 As String
   
   strHL03 = DBDATE(txtHL03)
   strHL04 = Left(cboHL04, 1)
   If cboHL05.ListIndex > 0 Then
      strHL05 = Left(cboHL05, 2)
   Else
      strHL05 = ""
   End If
   
   If CheckLogLogic(strHL01, strHL02, strHL03, strHL04) = False Then
      Exit Function
   End If
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
   strSql = "insert into HiRelationLog (HL01,HL02,HL03,HL04,HL05)" & _
      " values ('" & strHL01 & "'," & strHL02 & "," & strHL03 & ",'" & strHL04 & "','" & strHL05 & "')"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI

   UpdateSR08
   
   cnnConnection.CommitTrans
   InsertRec = True
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      If Err.Number = -2147217873 Then
         MsgBox "異動紀錄重複！"
      Else
         MsgBox Err.Description
      End If
   End If
End Function

Private Function DeleteRec() As Boolean
   Dim strHL03 As String, strHL04 As String
   
   strHL03 = DBDATE(txtHL03)
   strHL04 = Left(cboHL04, 1)
   
   If CheckLogLogic(strHL01, strHL02, strHL03, strHL04, True) = False Then
      Exit Function
   End If
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   strSql = "Delete HiRelationLog where HL01='" & strHL01 & "' and HL02=" & strHL02 & " and HL03=" & strHL03
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   UpdateSR08

   cnnConnection.CommitTrans
   DeleteRec = True
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Private Sub txtHL03_GotFocus()
   TextInverse txtHL03
End Sub

Private Sub txtHL03_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtHL03_Validate(Cancel As Boolean)
   If m_EditMode <> 0 Then
      If txtHL03 <> "" Then
         If ChkDate(txtHL03) = False Then
            Call txtHL03_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If txtHL03 = "" Then
      MsgBox "異動日期不可空白！"
      txtHL03.SetFocus
      Exit Function
   Else
      txtHL03_Validate bCancel
      If bCancel = True Then Exit Function
   End If
   
   If cboHL04.ListIndex = -1 Then
      MsgBox "請選擇異動類別！"
      cboHL04.SetFocus
      Exit Function
   ElseIf cboHL04.ListIndex = 1 Then
      cboHL05.ListIndex = 0
      cboHL05.Enabled = False
   End If
   
   'Add by Sindy 2021/9/1 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/9/1 END
   
   TxtValidate = True
End Function

'更新眷屬檔健保眷屬欄位
Private Sub UpdateSR08()
   Dim stSR08 As String, stCon As String
   
   stSR08 = "": stCon = " and SR08='Y'"
   chkSR08.Value = 0: Me.textHL05 = ""
   strSql = "select HL04,HL05 from hirelationlog a" & _
      " where HL01='" & strHL01 & "' and HL02=" & strHL02 & _
      " and HL03=(select max(b.HL03) from hirelationlog b where b.hl01=a.hl01 and b.hl02=a.hl02)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      '最後異動不是移出
      If RsTemp(0) <> "2" Then
         chkSR08.Value = 1
         setHL05 "" & RsTemp(1)
         stSR08 = "Y": stCon = " and SR08 is null"
      End If
   End If
   strSql = "UPDATE STAFF_RELATION SET SR08='" & stSR08 & "' WHERE SR01='" & strHL01 & "' and SR02=" & strHL02 & stCon
   cnnConnection.Execute strSql, intI
End Sub

'Add by Morgan 2009/6/24
'選取選單
Private Sub SelCombo(ByRef pCBO As ComboBox, ByVal pValue As String, Optional pLen As Integer = 2, Optional pNullIdx As Integer = 0)
   Dim idx As Integer
   If pValue = "" Then
      pCBO.ListIndex = pNullIdx
   Else
      For idx = 0 To pCBO.ListCount - 1
         If Left(pCBO.List(idx), pLen) = pValue Then
            pCBO.ListIndex = idx
            Exit For
         End If
      Next
   End If
End Sub

Private Sub SetcboHL04(Optional pValue As String)
   
   cboHL04.Clear
   If Pub_StrUserSt03 = "M21" And m_EditMode <> 0 Then
      cboHL04.AddItem "1 轉入"
      cboHL04.AddItem "2 轉出"
   ElseIf Pub_StrUserSt03 = "M31" And m_EditMode <> 0 Then
      cboHL04.AddItem "3 調整"
   Else
      cboHL04.AddItem "1 轉入"
      cboHL04.AddItem "2 轉出"
      cboHL04.AddItem "3 調整"
   End If
   If pValue <> "" Then
      SelCombo cboHL04, pValue, 1
   End If
End Sub
'Add by Morgan 2009/6/29
Private Function CheckLogLogic(pHL01 As String, pHL02 As String, pHL03 As String, pHL04 As String, Optional pDelete As Boolean) As Boolean
   Dim stLastState As String, stNextState As String
   '刪除
   If pDelete = True Then
      '檢查刪除後前後筆資料不可為移出和調整
      If pHL04 = "1" Then
         strSql = "select HL04 from HirelationLog a where HL01='" & pHL01 & "' and HL02=" & pHL02 & " and HL03=(select max(b.HL03) from HirelationLog b where b.HL01=a.HL01 and b.HL02=a.HL02 and  b.HL03<" & pHL03 & ")"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp(0) = "2" Then
               strSql = "select HL04 from HirelationLog a where HL01='" & pHL01 & "' and HL02=" & pHL02 & " and HL03=(select min(b.HL03) from HirelationLog b where b.HL01=a.HL01 and b.HL02=a.HL02 and  b.HL03>" & pHL03 & ")"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If RsTemp(0) = "3" Then
                      MsgBox "前一次為【移出】且後一次為【調整】時【加入】異動不可刪除，請確認！", vbExclamation
                     Exit Function
                  End If
               End If
            End If
         End If
      End If
   Else
      '調整的前一次異動不可為移出
      If pHL04 = "3" Then
         strSql = "select HL04 from HirelationLog a where HL01='" & pHL01 & "' and HL02=" & pHL02 & " and HL03=(select max(b.HL03) from HirelationLog b where b.HL01=a.HL01 and b.HL02=a.HL02 and  b.HL03<" & pHL03 & ")"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp(0) = "2" Then
               MsgBox "【調整】的前一次異動不可為【移出】，請確認！", vbExclamation
               Exit Function
            End If
         End If
      End If
      
      '移出的後一次異動不可為調整
      If pHL04 = "2" Then
         strSql = "select HL04 from HirelationLog a where HL01='" & pHL01 & "' and HL02=" & pHL02 & " and HL03=(select min(b.HL03) from HirelationLog b where b.HL01=a.HL01 and b.HL02=a.HL02 and  b.HL03>" & pHL03 & ")"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp(0) = "3" Then
               MsgBox "【移出】的後一次異動不可為【調整】，請確認！", vbExclamation
               Exit Function
            End If
         End If
      End If
   End If
   CheckLogLogic = True
End Function
