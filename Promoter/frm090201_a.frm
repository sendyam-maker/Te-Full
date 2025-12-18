VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_a 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護－管制期限提醒"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1515
   ClientWidth     =   9315
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "全部選取"
      Height          =   400
      Index           =   3
      Left            =   4665
      TabIndex        =   6
      Top             =   135
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8225
      TabIndex        =   5
      Top             =   135
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "預定會稿日修改(&E)"
      Height          =   400
      Index           =   1
      Left            =   6465
      TabIndex        =   4
      Top             =   135
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4668
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5784
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5565
      TabIndex        =   0
      Top             =   135
      Width           =   850
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4900
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8652
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
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
      _Band(0).Cols   =   1
   End
   Begin MSForms.Label lblName 
      Height          =   255
      Left            =   1050
      TabIndex        =   7
      Top             =   330
      Width           =   1875
      VariousPropertyBits=   27
      Caption         =   "lblName"
      Size            =   "3307;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人： "
      Height          =   180
      Index           =   35
      Left            =   150
      TabIndex        =   3
      Top             =   360
      Width           =   795
   End
End
Attribute VB_Name = "frm090201_a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; grd1改字型=新細明體-ExtB、lblName
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Create by Morgan 2010/11/16
Option Explicit
Public TextOk As Boolean
Public strContinue As Boolean
Public bolCancel As Boolean

Dim strSql As String, i As Integer, j As Integer

Private Sub cmdOK_Click(Index As Integer)
   Dim KeyWord As String, iRows As Integer

   Select Case Index
      Case 0 '繼續
         strContinue = True
         Unload Me
      Case 1  '寄mail
         With grd1
            For i = 1 To .Rows - 1
               .Visible = False
               If .TextMatrix(i, 0) = "V" Then
                  Exit For
               End If
            Next
            .Visible = True
            If i = .Rows Then
               MsgBox "請點選欲處理的資料"
               Exit Sub
            End If
         End With
         
         Me.Enabled = False
         With grd1
            iRows = 1
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "V" Then
                  With frm090201_a_1
                  .lblReveiver = grd1.TextMatrix(i, 10)
                  .SetContent grd1.TextMatrix(i, 7), grd1.TextMatrix(i, 8)
                  .lblCP06 = grd1.TextMatrix(i, 12)
                  .stCP10 = grd1.TextMatrix(i, 13)
                  .lblRecNo = grd1.TextMatrix(i, 11)
                  .lblCaseNo = grd1.TextMatrix(i, 1)
                  .lblCaseName = grd1.TextMatrix(i, 2)
                  .lblCaseProperty = grd1.TextMatrix(i, 3)
                  .lblEP06 = grd1.TextMatrix(i, 4)
                  .lblCP48 = grd1.TextMatrix(i, 5)
                  .lblEP09 = grd1.TextMatrix(i, 6)
                  .txtEP28 = Replace(grd1.TextMatrix(i, 7), "/", "")
                  .txtEP28.Tag = .txtEP28
                  .lblSalesDate = grd1.TextMatrix(i, 8)
                  .stToNo = grd1.TextMatrix(i, 9)
                  .lblSubject = grd1.TextMatrix(i, 1) & "(" & grd1.TextMatrix(i, 11) & ") 管制期限提醒!!"
                  End With
                  frm090201_a_1.Show vbModal
                  If bolCancel Then
                     Exit For
                  End If
                  .Visible = False
                  .TextMatrix(i, 0) = ""
                  For j = 0 To .Cols - 1
                     .row = i
                     .col = j
                     .CellBackColor = QBColor(15)
                  Next j
                  .RowHeight(i) = 0
                  iRows = iRows + 1
                  .Visible = True
               End If
            Next
         End With
         Me.Enabled = True
         If iRows = grd1.Rows Then
            strContinue = True
            Unload Me
         End If
         
      Case 2 '結束
         strContinue = False
         Unload Me
         
      Case 3
         Screen.MousePointer = vbHourglass
         If Trim(cmdok(3).Caption) = "全部選取" Then
            KeyWord = "V"
            cmdok(3).Caption = "全部取消"
         Else
            KeyWord = ""
            cmdok(3).Caption = "全部選取"
         End If
         With grd1
         .Visible = False
         For i = 1 To .Rows - 1
            If .RowHeight(i) > 0 Then
               .row = i
               .col = 0
               .Text = KeyWord
               For j = 0 To grd1.Cols - 1
                  .col = j
                  If KeyWord = "V" Then
                     .CellBackColor = &HFFC0C0
                  Else
                     .CellBackColor = QBColor(15)
                  End If
               Next j
            End If
         Next i
         .Visible = True
         End With
         Screen.MousePointer = vbDefault
   End Select
End Sub

Private Sub Form_Load()
   Me.Hide
   Screen.MousePointer = vbHourglass
   MoveFormToCenter Me
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache
   If strContinue = True Then Nextstep
   Set frm090201_a = Nothing
End Sub

Private Sub Nextstep()
   PUB_AddExcuteLog "frm090201_a" 'Added by Morgan 2013/10/8
   frm090201_2.Show   '工作進度維護
End Sub

Public Sub StrMenu1()
   Dim stCon As String
   strContinue = True
   
   'Modified by Morgan 2012/8/16 預定會稿日曾經改過的都不再顯示(主管修改也算)--柄佑
   'stCon = " and not exists(select * from mailcache where mc01=ep05 and mc03>ep06 and instr(mc07,ep02)>0 and instr(mc07,'管制期限提醒')>0)"
   stCon = " and nvl(ep30,0)=0"
   
   'Added by Morgan 2012/7/19
   stCon = stCon & " and ep28>=" & strSrvDate(1)
   
   'Modify By Sindy 2016/9/5 and cp57 is null ==> and cp159=0
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   strSql = "select ''" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",nvl(pa05,nvl(pa06,pa07)) 案件名稱" & _
      ",decode(pa09,'000',CPM03,CPM04) 案件性質" & _
      ",substrb(sqldatet(ep06),1,10) 齊備日" & _
      ",substrb(sqldatet(cp48),1,10) 承辦期限" & _
      ",substrb(sqldatet(ep09),1,10) 完稿日" & _
      ",substrb(sqldatet(ep28),1,10) 預定會稿日" & _
      ",substrb(sqldatet(workdayadd(decode(cp01,'P',10,21),ep06)),1,10) 智權人員管制日" & _
      ",cp13,st02,ep02,cp06,cp10" & _
      " From engineerprogress, caseprogress, patent, casepropertymap,staff" & _
      " where ep05='" & strUserNum & "' and ep07||ep08||ep10 is null and ep02<'B'" & _
      " and ep06>to_char(sysdate-365,'yyyymmdd')" & stCon & _
      " and cp09(+)=ep02 and cp159=0 and cp10 in (" & NewCasePtyList & ")" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and to_char(sysdate,'yyyymmdd')>=workdayadd(decode(cp01,'P',8,17),ep06)" & _
      " and st01(+)=cp13 order by 9,EP02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      TextOk = True
      grd1.Clear
      Set grd1.Recordset = RsTemp.Clone
      SetGrd
   Else
      TextOk = False
      Nextstep
   End If
End Sub

Private Sub SetGrd()
   Me.lblName.Caption = strUserName
   With grd1
      .Visible = False
      .row = 0
      .col = 0:   .Text = "V"
      .ColWidth(0) = 200
      .CellAlignment = flexAlignCenterCenter
      .col = 1:   .Text = "本所案號"
      .ColWidth(1) = 1300
      .CellAlignment = flexAlignCenterCenter
      .col = 2:   .Text = "案件名稱"
      .ColWidth(2) = 2050
      .CellAlignment = flexAlignCenterCenter
      .col = 3:   .Text = "案件性質"
      .ColWidth(3) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 4:  .Text = "齊備日"
      .ColWidth(4) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 5:  .Text = "承辦期限"
      .ColWidth(5) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 6:  .Text = "完稿日"
      .ColWidth(6) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 7:  .Text = "預定會稿日"
      .ColWidth(7) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 8:  .Text = "智權人員管制日"
      .ColWidth(8) = 1000
      .CellAlignment = flexAlignCenterCenter
      For i = 9 To .Cols - 1
         .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub grd1_SelChange()
   grd1.Visible = False
   grd1.col = 0
   grd1.row = grd1.MouseRow
   If grd1.MouseRow <> 0 Then
      If grd1.Text = "V" Then
         grd1.Text = ""
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = QBColor(15)
         Next i
      Else
         grd1.Text = "V"
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   grd1.Visible = True
End Sub
