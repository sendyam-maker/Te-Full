VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060113 
   BorderStyle     =   1  '單線固定
   Caption         =   "FMP案完稿日/核稿完成日輸入"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.CheckBox Check1 
      Caption         =   "已發文補資料"
      Height          =   285
      Left            =   7344
      TabIndex        =   16
      Top             =   1368
      Width           =   1416
   End
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
      Left            =   3720
      TabIndex        =   4
      Top             =   330
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   1
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "P"
      Top             =   450
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   1
      Top             =   450
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   2
      Top             =   450
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   3
      Top             =   450
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3870
      Left            =   90
      TabIndex        =   5
      Top             =   1800
      Width           =   8715
      _ExtentX        =   15363
      _ExtentY        =   6826
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
      Left            =   1710
      TabIndex        =   15
      Top             =   1470
      Width           =   5475
      VariousPropertyBits=   27
      Size            =   "9657;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   2
      Left            =   1710
      TabIndex        =   14
      Top             =   1140
      Width           =   5475
      VariousPropertyBits=   27
      Size            =   "9657;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   1
      Left            =   1710
      TabIndex        =   13
      Top             =   810
      Width           =   5475
      VariousPropertyBits=   27
      Size            =   "9657;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   450
      TabIndex        =   12
      Top             =   810
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1245
      TabIndex        =   11
      Top             =   810
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1245
      TabIndex        =   10
      Top             =   1140
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1245
      TabIndex        =   9
      Top             =   1470
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   825
      TabIndex        =   8
      Top             =   450
      Width           =   765
   End
End
Attribute VB_Name = "frm060113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/25 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Create by Morgan 2012/5/21
Option Explicit

Dim intLastRow As Integer

Private Sub SetGridHead()
    Dim i As Integer
    FixGrid MSHFlexGrid1
    With MSHFlexGrid1
        .Visible = False
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
        .col = 5: .ColWidth(.col) = 1200: .Text = "承辦期限"
        .CellAlignment = flexAlignCenterCenter
        .col = 6: .ColWidth(.col) = 1200: .Text = "核稿人"
        .CellAlignment = flexAlignCenterCenter
        .col = 7: .ColWidth(.col) = 1200: .Text = "核稿期限"
        For i = 8 To .Cols - 1
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
    End With
End Sub

Private Sub ClearGrid()
    Dim rstGrid As New ADODB.Recordset, stSQL As String
    
    stSQL = "SELECT 0,1,2,3,4,5,6,7,8,9,10,11 FROM DUAL WHERE ROWNUM<1"
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
            frm060113_1.Show
            Call frm060113_1.SetData(MSHFlexGrid1.Recordset, ii)
            Me.Hide
        End If
        
    End With
   
    
End Sub

Public Function SetGrid(Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim stSQL As String
   Dim stConCP As String
   Dim arrCaseNo(1 To 4) As String
   
On Error GoTo flgErr
   
   stConCP = " AND CP01 IN ('P','PS','CFP','CPS') AND SUBSTR(CP12,1,1)='F' AND SUBSTR(S1.ST03,1,1)='F'"
      
   arrCaseNo(1) = txtCaseNo(1)
   arrCaseNo(2) = Right("000000" & txtCaseNo(2), 6)
   arrCaseNo(3) = Right("0" & txtCaseNo(3), 1)
   arrCaseNo(4) = Right("00" & txtCaseNo(4), 2)
   
   txtCaseNo(1) = arrCaseNo(1)
   txtCaseNo(2) = arrCaseNo(2)
   txtCaseNo(3) = arrCaseNo(3)
   txtCaseNo(4) = arrCaseNo(4)
   
   If Check1.Value = vbUnchecked Then
      stConCP = stConCP & " AND CP27 IS NULL"
   End If
   
   'Modify by Amy 2015/01/14 +CP06/CP07/CP43
   'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
   If arrCaseNo(1) = "FG" Or arrCaseNo(1) = "PS" Or arrCaseNo(1) = "CPS" Then
      stSQL = " SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM04,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", '' EP04C, '' EP08T" & _
           ", '' PA08T,'' PA08,SP05 PA05,SP06 PA06,SP07 PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08," & IIf(strSrvDate(1) >= FCP核完日改用EP39, "ep39", "ep33") & ",cp113,cp114" & _
           ",sqldatet(cp06) cp06,sqldatet(cp07) cp07,cp43" & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, servicepractice, CASEPROPERTYMAP, STAFF S1" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           " AND CP09<'C' AND CP57 IS NULL AND CP10<>'201'" & stConCP & _
           " AND EP02(+)=CP09" & _
           " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14"
   
   Else
      '上完稿日
      stSQL = "SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM04,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", '' EP04C, '' EP08T" & _
           ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08," & IIf(strSrvDate(1) >= FCP核完日改用EP39, "ep39", "ep33") & ",cp113,cp114" & _
           ",sqldatet(cp06) cp06,sqldatet(cp07) cp07,cp43" & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1,PatentTrademarkMap" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           " AND CP09<'C' AND CP57 IS NULL AND CP10<>'201'" & stConCP & _
           " AND EP02(+)=CP09" & _
           " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14" & _
           " AND PTM02=PA08 AND PTM01='1' "
           
      '上核稿完成日(已完稿)
      stSQL = stSQL & " UNION SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM04,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", S2.ST02 EP04C, sqldatet(EP08) EP08T" & _
           ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08," & IIf(strSrvDate(1) >= FCP核完日改用EP39, "ep39", "ep33") & ",cp113,cp114" & _
           ",sqldatet(cp06) cp06,sqldatet(cp07) cp07,cp43" & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1, STAFF S2,PatentTrademarkMap" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           " AND CP09<'C' AND CP57 IS NULL AND CP10='201'" & stConCP & _
           " AND EP02(+)=CP09 and EP09>0" & _
           " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14 AND S2.ST01(+)=EP04" & _
           " AND PTM02=PA08 AND PTM01='1' "
   End If
            
   stSQL = stSQL & " ORDER BY CP05, CP09"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   Set MSHFlexGrid1.Recordset = RsTemp.Clone
   SetGridHead
   If intI = 1 Then
      With RsTemp
         lblCaseName(1) = "" & .Fields("PA05")
         lblCaseName(2) = "" & .Fields("PA06")
         lblCaseName(3) = "" & .Fields("PA07")
      End With
      SetGrid = True
   ElseIf bolMsg Then
      ShowNoData
      txtCaseNo(2).SetFocus
   End If
   Exit Function
   
flgErr:
   MsgBox Err.Description, vbCritical

End Function

Private Sub cmdSearch_Click()
    Call SetGrid
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    ClearGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frm090901 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdok.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdok.SetFocus
End Sub

Private Sub txtCaseNo_Change(Index As Integer)
    lblCaseName(1) = ""
    lblCaseName(2) = ""
    lblCaseName(3) = ""
    ClearGrid
End Sub

Private Sub txtCaseNo_GotFocus(Index As Integer)
    TextInverse txtCaseNo(Index)
    Select Case Index
        Case 2, 3, 4
            CloseIme
    End Select
End Sub

