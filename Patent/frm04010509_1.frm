VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010509_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "所外鑑定報告結果"
   ClientHeight    =   5130
   ClientLeft      =   -3090
   ClientTop       =   2025
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8040
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(F)"
      Default         =   -1  'True
      Height          =   330
      Left            =   3312
      TabIndex        =   4
      Top             =   624
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   3
      Top             =   648
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   2
      Top             =   648
      Width           =   255
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   7080
      TabIndex        =   6
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   6228
      TabIndex        =   5
      Top             =   48
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1752
      MaxLength       =   6
      TabIndex        =   1
      Top             =   648
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1260
      MaxLength       =   3
      TabIndex        =   0
      Top             =   648
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3735
      Left            =   165
      TabIndex        =   10
      Top             =   1350
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   8
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1245
      TabIndex        =   9
      Top             =   990
      Width           =   5130
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9049;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   1005
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   645
      Width           =   900
   End
End
Attribute VB_Name = "frm04010509_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Combo1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2016/10/7 END


'Add By Sindy 2022/7/1
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmkok_Click(Index As Integer)
 Dim i As Integer, bolChk As Boolean
   Select Case Index
      Case 0
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               bolChk = True
               Me.Tag = MSHFlexGrid1.TextMatrix(i, 6)
               Exit For
            End If
         Next
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         End If
         frm04010509_1.Hide
         If m_strIR01 <> "" Then
            If Not m_PrevForm Is Nothing Then
               Call frm04010509_2.SetParent(m_PrevForm)
            End If
            'Add By Sindy 2016/10/7
            frm04010509_2.m_strIR01 = m_strIR01
            frm04010509_2.m_strIR02 = m_strIR02
            frm04010509_2.m_strIR03 = m_strIR03
            frm04010509_2.m_strIR04 = m_strIR04
            '2016/10/7 END
         End If
         frm04010509_2.Show
       Case 1
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
'Add By Cheng 2002/07/08
Dim StrSQLa As String

   If Text1(0).Text <> "P" And Text1(0).Text <> "PS" Then
      MsgBox "只可為 P 或 PS 案件 !", vbInformation
      Text1(0).SetFocus
      Exit Sub
   End If
   
   If Text1(2) = "" Then Text1(2) = "0"
   If Text1(3) = "" Then Text1(3) = "00"

   pa(1) = Text1(0)
   pa(2) = Text1(1)
   pa(3) = Text1(2)
   pa(4) = Text1(3)
   
   If pa(1) = "P" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
      Else
         Text1(1).SetFocus
         TextInverse Text1(1)
         Exit Sub
      End If
   ElseIf pa(1) = "PS" Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
      Else
         Text1(1).SetFocus
         TextInverse Text1(1)
         Exit Sub
      End If
   End If
   'Modify By Cheng 2002/07/08
   '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'   strExc(0) = "select ''," & SQLDate("cp05") & ", decode(cp10,'000',cpm03,cpm04)," & _
'      SQLDate("cp27") & ",decode(cp24,'1','准','2','駁','')," & _
'      "nvl(fa05,nvl(fa04,nvl(fa06,''))),CP09 from " & _
'      "caseprogress,fagent,casepropertymap where " & _
'      "cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and " & _
'      "cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' and " & _
'      "cp10='906' and (cp27<>'' or cp27 is not null) and (cp44<>'' or  cp44 is not null) and " & _
'      "( cp09<'C' ) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND " & _
'      "SUBSTR(CP44,9,1)=FA02(+)"
   StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
' 91.09.13 modify by louis (,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD)
'   strExc(0) = "select ''," & SQLDate("cp05") & ", decode(cp10,'000',cpm03,cpm04)," & _
'      SQLDate("cp27") & ",decode(cp24,'1','准','2','駁','')," & _
'      strSQLA & "CP09 From " & _
'      "caseprogress,fagent,casepropertymap,SystemKind Where " & _
'      "cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and " & _
'      "cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' and " & _
'      "cp10='906' and (cp27<>'' or cp27 is not null) and (cp44<>'' or  cp44 is not null) and " & _
'      "( cp09<'C' ) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND " & _
'      "SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) "
   strExc(0) = "select ''," & SQLDate("cp05") & ", decode(cp10,'000',cpm03,cpm04)," & _
      SQLDate("cp27") & ",decode(cp24,'1','准','2','駁','')," & _
      StrSQLa & "CP09,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD FROM " & _
      "caseprogress,fagent,casepropertymap,SystemKind Where " & _
      "cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and " & _
      "cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' and " & _
      "cp10='906' and (cp27<>'' or cp27 is not null) and (cp44<>'' or  cp44 is not null) and " & _
      "( cp09<'C' ) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND " & _
      "SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & _
      "ORDER BY SORTFIELD DESC "
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1_Click
      cmkok_Click 0
   ElseIf MSHFlexGrid1.Rows = 1 Then
      
   Else
      cmkok(0).SetFocus
   End If
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      Text1(0).Text = m_strCP01
      Text1(1).Text = m_strCP02
      Text1(2).Text = m_strCP03
      Text1(3).Text = m_strCP04
      Command1.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Initialize()
   'add by nickc 2007/02/02
   ReDim pa(1 To TF_PA) As String
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   InitGrid 7, MSHFlexGrid1
   GridHead
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/7/1
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/7/1 END
   
   Set frm04010509_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(0) <> "" Then
      If Text1(0).Text <> "P" And Text1(0).Text <> "PS" Then
         MsgBox "只可為 P 或 PS 案件 !", vbInformation
         Cancel = True
         TextInverse Text1(Index)
      End If
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
      .col = 2: .ColWidth(2) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1400: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "代理人"
      .col = 6: .ColWidth(6) = 0
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
End Sub
