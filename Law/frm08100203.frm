VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm08100203 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件進度查詢"
   ClientHeight    =   5745
   ClientLeft      =   90
   ClientTop       =   720
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7164
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7992
      TabIndex        =   1
      Top             =   70
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4212
      Left            =   168
      TabIndex        =   2
      Top             =   1488
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      RowSizingMode   =   1
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
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1428
      TabIndex        =   9
      Top             =   705
      Width           =   7245
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12779;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2460
      TabIndex        =   8
      Top             =   1080
      Width           =   6435
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11351;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 ："
      Height          =   252
      Left            =   228
      TabIndex        =   7
      Top             =   396
      Width           =   1092
   End
   Begin VB.Label lbeCaseNumber 
      Height          =   252
      Left            =   1428
      TabIndex        =   6
      Top             =   396
      Width           =   1932
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   705
      Width           =   1095
   End
   Begin VB.Label lbeCustomer 
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "當  事  人："
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   915
   End
End
Attribute VB_Name = "frm08100203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; cboCaseName、lbeCusName、MSHFlexGrid1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim LcTmp As String
Dim intLastRow As Integer, intCols As Integer

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdSure_Click()
 Dim strNum As String
   With MSHFlexGrid1
      .col = 0
      If .Text = "v" Then
         .col = 1
         strNum = .Text
      End If
   End With
   frm081002.Text(15) = strNum
   Unload Me
   frm081002.Show
End Sub

Private Sub Form_Load()
 Dim i As Integer, temp(2 To 4) As String
   IsNoExistData = False
   MoveFormToCenter Me
   lbeCaseNumber = frm081002.lbeNumber
   lbeCustomer = frm081002.Text(1)
   lbeCusName = frm081002.lbe(1)
   temp(2) = "中:"
   temp(3) = "英:"
   temp(4) = "日:"
   For i = 2 To 4
      If frm081002.Text(i) <> "" Then
         cboCaseName.AddItem temp(i) + frm081002.Text(i)
      End If
   Next
   If cboCaseName.ListCount <> 0 Then
      cboCaseName.ListIndex = 0
   End If
   LcTmp = frm081002.lbeNumber.Tag
   GetGridData
   GridHead
   cmdSure.Enabled = False
End Sub

Private Sub GridHead()
   With MSHFlexGrid1
      .row = 0
      .col = 0
      .ColWidth(0) = 200
      .Text = "v"
      .col = 1
      .ColWidth(1) = 1000
      .Text = "收文號"
      .col = 2
      .ColWidth(2) = 800
      .Text = "收文日"
      .col = 3
      .ColWidth(3) = 1000
      .Text = "案件性質"
      .col = 4
      .ColWidth(4) = 900
      .Text = "發文日"
      .col = 5
      .ColWidth(5) = 800
      .Text = "後金"
      .col = 6
      .ColWidth(6) = 800
      .Text = "結果"
      .col = 7
      .ColWidth(7) = 900
      .Text = "相關人"
   End With
End Sub

Private Sub GetGridData()
 Dim rs As New ADODB.Recordset, i As Integer
   strExc(1) = "select '',cp09,decode(cp05,null,'',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2))" + _
      ", decode(lc15,020, CPM04, CPM03), decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 " + _
      " || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)),cp19,decode(cp24,1,'淮/勝',2,'駁/敗'),cp50 from caseprogress," + _
      " CASEPROPERTYMAP ,lawcase where " & ChgCaseprogress(LcTmp) + " and cp01=cpm01(+) and cp10=cpm02(+)" + _
      " and " & ChgLawcase(LcTmp) & " and cp09<>" + CNULL(frm081002.lbePaperNum)
   intI = 0
   Set rs = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      rs.MoveFirst
      With MSHFlexGrid1
         Set .Recordset = rs
      End With
   Else
      IsNoExistData = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm08100203 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
 Dim i As Integer
   With MSHFlexGrid1
      intCols = .Cols - 1
      i = .row
      .row = intLastRow
      .col = 0
      If .Text = "v" Then
         .Text = ""
      Else
         .Text = "v"
      End If
      .row = i
   End With
   If Not CheckGridChoese(MSHFlexGrid1, intLastRow, intCols) Then Exit Sub
   cmdSure.Enabled = True
End Sub
