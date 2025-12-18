VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_1_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外相關案件"
   ClientHeight    =   4590
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7275
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   5820
      TabIndex        =   9
      Top             =   30
      Width           =   1200
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   5
      Top             =   144
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2220
      MaxLength       =   1
      TabIndex        =   4
      Top             =   144
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   3
      Top             =   144
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   900
      MaxLength       =   3
      TabIndex        =   2
      Top             =   144
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060104_1_1.frx":0000
      Left            =   900
      List            =   "frm060104_1_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3645
      Left            =   15
      TabIndex        =   0
      Top             =   840
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   6429
      _Version        =   393216
      BackColor       =   255
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   90
      TabIndex        =   8
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   90
      TabIndex        =   7
      Top             =   180
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   5670
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "10001;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm060104_1_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/12 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer


Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(0) = pa(5)
      Case "英"
         Label2(0) = pa(6)
      Case "日"
         Label2(0) = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()


   intWhere = 國內
   With frm060104_1
      pa(1) = .text1
      pa(2) = .Text2
      pa(3) = .Text3
      pa(4) = .Text4
   End With
   text1 = pa(1)
   Text2 = pa(2)
   Text3 = pa(3)
   Text4 = pa(4)
   If ClsPDReadPatentDatabase(pa, intWhere) Then  'edit by nickc 2007/02/02 不用 dll 了  If objPublicData.ReadPatentDatabase(pA, intWhere) Then
      Label2(0) = pa(5)
      Combo1.ListIndex = 0
   End If
   
   strExc(0) = "select cm01||cm02||cm03||cm04,ST02,nvl(pa05,nvl(pa06,pa07)) FROM " & _
      "CASEMAP,CASEPROGRESS,PATENT,STAFF WHERE CM05='" & pa(1) & "' AND CM06='" & pa(2) & "' AND " & _
      "CM07='" & pa(3) & "' AND CM08='" & pa(4) & "' AND CM10='0' AND " & _
      "cm01=pa01 and cm02=pa02 and cm03=pa03 and cm04=pa04 AND " & _
      "cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04 AND " & _
      "CP27 IS NULL and CP57 IS NULL and cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + ")" & _
      " and cp14=st01(+) ORDER BY cm01,cm02,cm03,CM04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_1_1 = Nothing
End Sub

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 0: .ColWidth(0) = 1200: .Text = "國外案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 4500: .Text = "案件名稱"
      .Visible = True
   End With
End Sub
