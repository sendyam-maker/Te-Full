VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060108 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦/核稿期限、會稿日輸入"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7185
      TabIndex        =   6
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8025
      TabIndex        =   7
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6345
      TabIndex        =   4
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   1
      Left            =   1770
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "FCP"
      Top             =   630
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   2250
      MaxLength       =   6
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   3090
      MaxLength       =   1
      TabIndex        =   2
      Top             =   630
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   3
      Top             =   630
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3750
      Left            =   90
      TabIndex        =   5
      Top             =   1950
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   6615
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
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   3
      Left            =   1770
      TabIndex        =   17
      Top             =   1620
      Width           =   7000
      VariousPropertyBits=   27
      Size            =   "12347;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   2
      Left            =   1770
      TabIndex        =   16
      Top             =   1290
      Width           =   7000
      VariousPropertyBits=   27
      Size            =   "12347;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   1
      Left            =   1770
      TabIndex        =   15
      Top             =   960
      Width           =   7000
      VariousPropertyBits=   27
      Size            =   "12347;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblAppDate 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   6120
      TabIndex        =   14
      Top             =   630
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Index           =   1
      Left            =   5400
      TabIndex        =   13
      Top             =   630
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   450
      TabIndex        =   12
      Top             =   990
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1245
      TabIndex        =   11
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1245
      TabIndex        =   10
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1245
      TabIndex        =   9
      Top             =   1620
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   825
      TabIndex        =   8
      Top             =   630
      Width           =   765
   End
End
Attribute VB_Name = "frm060108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/25 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intLastRow As Integer


Private Sub SetGridHead()
    Dim i As Integer
    FixGrid MSHFlexGrid1
    With MSHFlexGrid1
        .Visible = False
        .Cols = 19
        .row = 0
        .col = 0: .ColWidth(.col) = 200: .Text = "v"
        .CellAlignment = flexAlignCenterCenter
        .col = 1: .ColWidth(.col) = 900: .Text = "收文日"
        .CellAlignment = flexAlignCenterCenter
        .col = 2: .ColWidth(.col) = 1300: .Text = "收文號"
        .CellAlignment = flexAlignCenterCenter
        .col = 3: .ColWidth(.col) = 1400: .Text = "案件性質"
        .CellAlignment = flexAlignCenterCenter
        .col = 4: .ColWidth(.col) = 1200: .Text = "承辦人"
        .CellAlignment = flexAlignCenterCenter
        .col = 5: .ColWidth(.col) = 1200: .Text = "核稿人"
        .CellAlignment = flexAlignCenterCenter
        .col = 6: .ColWidth(.col) = 1200: .Text = "完稿日"
        For i = 7 To 18
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
    End With
End Sub

Private Sub ClearGrid()
    Dim rstGrid As New ADODB.Recordset, stSQL As String
    
    stSQL = "SELECT 0, 1,2,3,4,5,6,7,8, 9, 10, 11 FROM DUAL WHERE ROWNUM<1"
    rstGrid.CursorLocation = adUseClient
    rstGrid.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    Set MSHFlexGrid1.Recordset = rstGrid
    SetGridHead
    Set rstGrid = Nothing
End Sub

Private Sub cmdExit_Click()
    blnIsFormBack = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim ii As Integer
    
    If MSHFlexGrid1.Rows < 2 Then Exit Sub
    
    With MSHFlexGrid1
        .Visible = False
        For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = "v" Then Exit For
        Next ii
        .Visible = True
        If ii = .Rows Then
            MsgBox "請點選欲輸入資料！"
        Else
            frm060108_1.Show
            Call frm060108_1.SetData(MSHFlexGrid1.Recordset, ii)
            Me.Hide
        End If
        
    End With
   
    
End Sub

Public Sub SetGrid(Optional ByVal bolMsg As Boolean = True)

On Error GoTo flgErr

    Dim rstGrid As New ADODB.Recordset
    Dim stSQL As String
    Dim arrCaseNo(1 To 4) As String
    
    arrCaseNo(1) = txtCaseNo(1)
    arrCaseNo(2) = Right("000000" & txtCaseNo(2), 6)
    arrCaseNo(3) = Right("0" & txtCaseNo(3), 1)
    arrCaseNo(4) = Right("00" & txtCaseNo(4), 2)
    'Modify by Morgan 2009/11/9 +判斷系統別及國外部智權人員收文案件(FMP案)
    stSQL = "SELECT '' V" & _
            ", DECODE(CP05,NULL,NULL,(SUBSTR(CP05,1,4)-1911)||SUBSTR(CP05,5,2)||SUBSTR(CP05,7,2)) CP05T" & _
            ", CP09,NVL(CPM03,CP10) CP10T" & _
            ", S1.ST02 CP14T" & _
            ", S2.ST02 EP04T" & _
            ", DECODE(EP09,NULL,NULL,EP09-19110000) EP09T" & _
            ", DECODE(PA10,NULL,NULL,PA10-19110000) PA10T" & _
            ", PA08, PTM03 PA08T, CP64, S1.ST03" & _
            ", PA05, PA06, PA07, CP05, CP06, CP10, CP14, EP04, EP09, PA10" & _
            ", DECODE(EP07,NULL,NULL,EP07-19110000) EP07T" & _
            ", DECODE(EP08,NULL,NULL,EP08-19110000) EP08T" & _
            ", DECODE(EP09,NULL,NULL,EP09-19110000) EP09T, EP34" & _
            ", DECODE(CP48,NULL,NULL,CP48-19110000) CP48T" & _
            " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1, STAFF S2, PatentTrademarkMap" & _
            " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
            " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
            " AND EP02=CP09" & _
            " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
            " AND S1.ST01(+)=CP14 AND S2.ST01(+)=EP04" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
            " AND CP10 IN ('201','209','210') AND CP27 IS NULL AND CP57 IS NULL" & _
            " AND PTM02=PA08 AND PTM01='1'" & _
            " AND CP01 IN ('FCP','FG','P') AND SUBSTR(CP12,1,1)='F'" & _
            " ORDER BY CP05, CP09"

    rstGrid.CursorLocation = adUseClient
    rstGrid.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    
    If rstGrid.RecordCount > 0 Then
        txtCaseNo(1) = arrCaseNo(1)
        txtCaseNo(2) = arrCaseNo(2)
        txtCaseNo(3) = arrCaseNo(3)
        txtCaseNo(4) = arrCaseNo(4)
        lblAppDate = "" & rstGrid.Fields("PA10T")
        lblCaseName(1) = "" & rstGrid.Fields("PA05")
        lblCaseName(2) = "" & rstGrid.Fields("PA06")
        lblCaseName(3) = "" & rstGrid.Fields("PA07")
    ElseIf bolMsg Then
        ShowNoData
    End If
    
    Set MSHFlexGrid1.Recordset = rstGrid
    SetGridHead
    Set rstGrid = Nothing
    
flgErr:

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
    End If

End Sub

Private Sub cmdSearch_Click()
    SetGrid
    If Me.MSHFlexGrid1.Rows = 2 And Me.Visible = True Then
        MSHFlexGrid1.row = 1
        GridClick MSHFlexGrid1, intLastRow, 0
        cmdOK_Click
   End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    ClearGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frm060108 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK.SetFocus
End Sub

Private Sub txtCaseNo_Change(Index As Integer)
    lblAppDate = ""
    lblCaseName(1) = ""
    lblCaseName(2) = ""
    lblCaseName(3) = ""
    ClearGrid
End Sub

Private Sub txtCaseNo_GotFocus(Index As Integer)
    TextInverse txtCaseNo(Index)
    Select Case Index
        Case 2, 3, 4
            'edit by nickc 2007/07/11 切換輸入法改用API
            'txtCaseNo(Index).IMEMode = 2
            CloseIme
    End Select
End Sub

Private Sub txtCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1, 3
         KeyAscii = UpperCase(KeyAscii)
         
      Case 2, 4
         If KeyAscii <> 8 And (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") Then
            KeyAscii = 0
         End If
   End Select
End Sub
