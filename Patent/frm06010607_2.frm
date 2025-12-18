VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010607_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "延期受理輸入"
   ClientHeight    =   5745
   ClientLeft      =   -150
   ClientTop       =   900
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7092
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6264
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8316
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm06010607_2.frx":0000
      Left            =   1080
      List            =   "frm06010607_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   960
      Width           =   615
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
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   3
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   2
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   1
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   0
      Top             =   660
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4212
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   7435
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
   Begin MSForms.Label Label8 
      Height          =   285
      Left            =   1740
      TabIndex        =   13
      Top             =   990
      Width           =   7410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13070;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3900
      TabIndex        =   11
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   660
      Width           =   768
   End
End
Attribute VB_Name = "frm06010607_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/23 Form2.0已修改
'Create By Sindy 2016/8/11 參考frm06010604_2撰寫
Option Explicit

Dim strReceiveNo As String, strTemp As String
Dim pa() As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 1
         frm06010607_1.Show
         Unload Me
      Case 2
         Unload frm06010607_1
         Unload Me
   End Select
End Sub

'確認鈕
Public Sub FormConfirm()
 Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            .TextMatrix(i, 0) = "" 'Add By Sindy 2016/8/17
            strExc(5) = .TextMatrix(i, 3)
            strExc(4) = .TextMatrix(i, 9)
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      If Me.Tag = "" Then
         MsgBox "請選擇資料 !", vbInformation
         Exit Sub
      Else
         frm06010607_1.Show
         frm06010607_1.Clear
         Unload Me
         Exit Sub
      End If
   End If
   frm06010607_3.Show
   Me.Hide
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label8 = pa(5)
      Case "英"
         Label8 = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label8 = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   ReadPatent
   Combo1.ListIndex = 0
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As Object, i As Integer, j As Integer
Dim strTmp(0 To 5) As String, varTmp As Variant
Dim intOnlyOneRow As Integer, intRow As Integer 'Add By Sindy 2016/8/11
   
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   Label8 = ""
   If pa(1) = "FCP" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         Label8 = pa(5)
         Text1 = pa(11)
      End If
   Else
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         Label8 = pa(5)
         Text1 = pa(11)
      End If
   End If
   
   '只帶出已發文之延期進度
   strExc(0) = "SELECT '',CP09,CPM03," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
      "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD,cp43" & _
      " from caseprogress,casepropertymap,CUSTOMER where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " and (cp09<'C') and cp27 is not null" & _
      " and cp01=cpm01(+) and cp10=cpm02(+)" & _
      " AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
      " and cp10='404'"
   '排序改依收文日,收文號由小至大--靜芳
   strExc(0) = strExc(0) & " ORDER BY CP05 asc,CP09 asc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   '逐筆檢查資料,只留無延期受理的資料
   MSHFlexGrid1.Visible = False
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 1) <> "" Then
            strExc(0) = "select * from caseprogress where cp43='" & .TextMatrix(i, 1) & "' and cp10='1004'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MSHFlexGrid1.RowHeight(i) = 0
            Else
               intRow = intRow + 1
               If intRow = 1 Then
                  intOnlyOneRow = i
               Else
                  intOnlyOneRow = 0
               End If
               '案件性質欄顯示延期-相關總收文號案件性質
               If Me.MSHFlexGrid1.TextMatrix(i, 11) <> "" Then
                  Me.MSHFlexGrid1.TextMatrix(i, 2) = Me.MSHFlexGrid1.TextMatrix(i, 2) & PUB_GetRelateCasePropertyName(Me.MSHFlexGrid1.TextMatrix(i, 1), "1")
               End If
            End If
         End If
      Next
   End With
   MSHFlexGrid1.Visible = True
   
   ' 若只有一筆資料時直接進入到下一個畫面   InStr(NewCasePtyList, txtCaseField(7)) > 0
   'If MSHFlexGrid1.Rows = 2 Then
   If intOnlyOneRow > 0 Then
      'MSHFlexGrid1.row = 1
      MSHFlexGrid1.row = intOnlyOneRow
      GridClick MSHFlexGrid1, intLastRow, 0
      FormConfirm
   End If

   '預設新申請案那一道
   MSHFlexGrid1.Visible = False
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .RowHeight(i) > 0 Then
            If InStr(NewCasePtyList, .TextMatrix(i, 9)) > 0 Then
               .TextMatrix(i, 0) = "v"
               .row = i
               For j = 0 To .Cols - 1
                  .col = j
                  .CellBackColor = &HFFC0C0
               Next
            End If
         End If
      Next
   End With
   MSHFlexGrid1.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010607_2 = Nothing
End Sub

Private Sub GridHead()
Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .col = 1: .ColWidth(1) = 900: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 600: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 1500: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      For i = 9 To .Cols - 1
         .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0, 1 '多筆選取
End Sub
