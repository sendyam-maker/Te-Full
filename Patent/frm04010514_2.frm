VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010514_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "初審及公佈通知來函輸入"
   ClientHeight    =   5760
   ClientLeft      =   -2970
   ClientTop       =   4590
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "內部收文(&E)"
      Height          =   400
      Index           =   3
      Left            =   5088
      TabIndex        =   14
      Top             =   72
      Width           =   1200
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   8
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   7
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   6
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   5
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4980
      TabIndex        =   4
      Top             =   660
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm04010514_2.frx":0000
      Left            =   1080
      List            =   "frm04010514_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   1020
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8352
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6300
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7128
      TabIndex        =   0
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4245
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   7488
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3900
      TabIndex        =   12
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   1020
      Width           =   768
   End
   Begin MSForms.Label Label8 
      Height          =   270
      Left            =   1800
      TabIndex        =   10
      Top             =   1020
      Width           =   4800
      VariousPropertyBits=   27
      Caption         =   "Label8"
      Size            =   "8467;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm04010514_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (Label8)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2009/11/24 自內專核准函輸入抽出
Option Explicit

Dim strReceiveNo As String, strTemp As String
Dim pa() As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0 '確定
         FormConfirm
      Case 1 '回前畫面
         frm04010514_1.Show
         Unload Me
      Case 2 '結束
         Unload frm04010514_1
         Unload Me
      Case 3 '內部收文
         mdiMain.mnu1102_Click 1
   End Select
End Sub

' 確認鈕
Private Sub FormConfirm()
 Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            If InStr(CaseMapIn, .TextMatrix(i, 12)) > 0 Then
               If pa(10) = "" Then
                  MsgBox "新申請案不可無申請日"
                  Exit Sub
               End If
            End If
           
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   'Add By Sindy 2016/10/5
   frm04010514_3.m_strIR01 = m_strIR01
   frm04010514_3.m_strIR02 = m_strIR02
   frm04010514_3.m_strIR03 = m_strIR03
   frm04010514_3.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010514_3.Show
   Me.Hide
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label8 = pa(5)
      Case "英"
         Label8 = pa(6)
      Case "日"
         Label8 = pa(7)
   End Select
End Sub

Private Sub Form_Activate()
   If Me.Tag = "" Then
      Me.Tag = "1"
      ' 若只有一筆資料時自動選取第一筆
      If MSHFlexGrid1.Rows = 2 Then
         MSHFlexGrid1.row = 1
         GridClick MSHFlexGrid1, intLastRow, 0
         FormConfirm
      End If
   End If
End Sub

Private Sub Form_Initialize()
   ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010514_1.m_strIR01
   m_strIR02 = frm04010514_1.m_strIR02
   m_strIR03 = frm04010514_1.m_strIR03
   m_strIR04 = frm04010514_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   ReadPatent
End Sub

Private Sub ReadPatent()
   Dim Lbl As Label, txt As TextBox, i As Integer
   Dim strTmp As String
   Label8 = ""
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      Label8 = pa(5)
      Text1 = pa(11)
   End If
   'Modify by Morgan 2010/1/25 澳門改所有專利種類都可輸入(原控制新型不可輸)
   strExc(0) = "select '',CP09,CPM04,CP43,CP40,CP36,SQLDATET(CP06),SQLDATET(CP07),SQLDATET(CP27)" & _
      ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64,CP10 " & _
      ", DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
      " from caseprogress,PATENT,casepropertymap " & _
      " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " and cp27 is not null and cp24 is null and CP09<'C' " & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
      " AND (CP10 IN ('101','109','110') OR (PA08='1' AND CP10='307') OR (PA09='044' AND CP10 IN ('101','102','103')))" & _
      " and cp01=cpm01(+) and cp10=cpm02(+)" & _
      "ORDER BY SORTFIELD DESC "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010514_2 = Nothing
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1000: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1000: .Text = "對造號數"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .ColWidth(9) = 600: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 10: .ColWidth(10) = 800: .Text = "後金"
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .ColWidth(11) = 1000: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 12: .ColWidth(12) = 0: .Text = "案件性質代號"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(0).SetFocus
   
End Sub
