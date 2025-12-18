VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071016 
   BorderStyle     =   1  '單線固定
   Caption         =   "會稿日輸入"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7185
      TabIndex        =   5
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8025
      TabIndex        =   6
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
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   0
      Top             =   270
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   1650
      MaxLength       =   6
      TabIndex        =   1
      Top             =   270
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   2490
      MaxLength       =   1
      TabIndex        =   2
      Top             =   270
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   3
      Top             =   270
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3480
      Left            =   180
      TabIndex        =   7
      Top             =   1920
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   6138
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
   Begin MSForms.Label LblCustName 
      Height          =   285
      Left            =   2430
      TabIndex        =   18
      Top             =   1560
      Width           =   5985
      VariousPropertyBits=   27
      Size            =   "10557;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   3
      Left            =   1470
      TabIndex        =   17
      Top             =   1200
      Width           =   6915
      VariousPropertyBits=   27
      Size            =   "12197;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   2
      Left            =   1470
      TabIndex        =   16
      Top             =   889
      Width           =   6915
      VariousPropertyBits=   27
      Size            =   "12197;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   1
      Left            =   1470
      TabIndex        =   15
      Top             =   578
      Width           =   6915
      VariousPropertyBits=   27
      Size            =   "12197;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblCustNo 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1140
      TabIndex        =   14
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "當事人"
      Height          =   180
      Left            =   210
      TabIndex        =   13
      Top             =   1605
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1020
      TabIndex        =   11
      Top             =   630
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1020
      TabIndex        =   10
      Top             =   941
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(日):"
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   9
      Top             =   1252
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   8
      Top             =   270
      Width           =   765
   End
End
Attribute VB_Name = "frm071016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、lblCaseName(index)、LblCustName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
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
      .col = 1: .ColWidth(.col) = 800: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(.col) = 1200: .Text = "總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(.col) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(.col) = 800: .Text = "會稿日"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(.col) = 800: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      'Modified by Lydia 2015/10/05
      '.col = 6: .ColWidth(.col) = 800: .Text = "承辦律師"
      .col = 6: .ColWidth(.col) = 800: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      'Modified by Lydia 2015/10/05
      '.col = 7: .ColWidth(.col) = 800: .Text = "承辦法務"
      .col = 7: .ColWidth(.col) = 800: .Text = "協辦人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(.col) = 2000: .Text = "進度備註"
      For i = 9 To 18
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

Private Sub cmdok_Click()

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
         frm071017.Show
         Call frm071017.SetData(MSHFlexGrid1.Recordset, ii)
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
   If txtCaseNo(1) = "LA" Then
      stSQL = "SELECT '' V, DECODE(CP05,NULL,NULL,(SUBSTR(CP05,1,4)-1911)||SUBSTR(CP05,5,2)||SUBSTR(CP05,7,2)) CP05T" & _
              ", CP09,NVL(CPM03,CP10) CP10T" & _
              ", DECODE(EP07,NULL,NULL,(SUBSTR(EP07,1,4)-1911)||SUBSTR(EP07,5,2)||SUBSTR(EP07,7,2)) EP07T" & _
              ", S1.ST02 CP13T, S2.ST02 CP14T, S3.ST02 CP29T, CP64, HC06, '', '', HC05, NVL(CU04,DECODE(CU05,'',CU06,CU05||' '||CU88||' '||CU89)) AS CU04, CP05, CP10, CP13, CP14, CP29, EP07" & _
              " FROM CASEPROGRESS, ENGINEERPROGRESS, HIRECASE, CASEPROPERTYMAP, STAFF S1, STAFF S2, STAFF S3, CUSTOMER" & _
              " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
              " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
              " AND EP02=CP09 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
              " AND S1.ST01(+)=CP13 AND S2.ST01(+)=CP14 AND S3.ST01(+)=CP29" & _
              " AND HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04" & _
              " AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+)" & _
              " AND CP27 IS NULL AND CP57 IS NULL ORDER BY CP05 DESC, CP09"
   Else
      stSQL = "SELECT '' V, DECODE(CP05,NULL,NULL,(SUBSTR(CP05,1,4)-1911)||SUBSTR(CP05,5,2)||SUBSTR(CP05,7,2)) CP05T" & _
              ", CP09,NVL(CPM03,CP10) CP10T" & _
              ", DECODE(EP07,NULL,NULL,(SUBSTR(EP07,1,4)-1911)||SUBSTR(EP07,5,2)||SUBSTR(EP07,7,2)) EP07T" & _
              ", S1.ST02 CP13T, S2.ST02 CP14T, S3.ST02 CP29T, CP64, LC05, LC06, LC07, LC11, NVL(CU04,DECODE(CU05,'',CU06,CU05||' '||CU88||' '||CU89)) AS CU04, CP05, CP10, CP13, CP14, CP29, EP07" & _
              " FROM CASEPROGRESS, ENGINEERPROGRESS, LAWCASE, CASEPROPERTYMAP, STAFF S1, STAFF S2, STAFF S3, CUSTOMER" & _
              " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
              " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
              " AND EP02=CP09 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
              " AND S1.ST01(+)=CP13 AND S2.ST01(+)=CP14 AND S3.ST01(+)=CP29" & _
              " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
              " AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+)" & _
              " AND CP27 IS NULL AND CP57 IS NULL ORDER BY CP05 DESC, CP09"
   End If

   rstGrid.CursorLocation = adUseClient
   rstGrid.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   
   If rstGrid.RecordCount > 0 Then
      'Modify By Sindy 2011/7/7
      If txtCaseNo(1) = "LA" Then
         txtCaseNo(1) = arrCaseNo(1)
         txtCaseNo(2) = arrCaseNo(2)
         txtCaseNo(3) = arrCaseNo(3)
         txtCaseNo(4) = arrCaseNo(4)
         lblCaseName(1) = "" & rstGrid.Fields("HC06")
         lblCaseName(2) = ""
         lblCaseName(3) = ""
         LblCustNo = "" & rstGrid.Fields("HC05")
         lblCustName = "" & rstGrid.Fields("CU04")
      '2011/7/7 End
      Else
         txtCaseNo(1) = arrCaseNo(1)
         txtCaseNo(2) = arrCaseNo(2)
         txtCaseNo(3) = arrCaseNo(3)
         txtCaseNo(4) = arrCaseNo(4)
         lblCaseName(1) = "" & rstGrid.Fields("LC05")
         lblCaseName(2) = "" & rstGrid.Fields("LC06")
         lblCaseName(3) = "" & rstGrid.Fields("LC07")
         LblCustNo = "" & rstGrid.Fields("LC11")
         lblCustName = "" & rstGrid.Fields("CU04")
      End If
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
      cmdok_Click
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ClearGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   Set frm071016 = Nothing
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
      Case 1, 2, 3, 4
         'edit by nickc 2007/06/11  切換輸入法改用API
         'txtCaseNo(Index).IMEMode = 2
         CloseIme
   End Select
End Sub

Private Sub txtCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      'Add By Sindy 2011/7/7
      Case 1
        KeyAscii = UpperCase(KeyAscii)
      Case 2, 3, 4
         If KeyAscii <> 8 And (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") Then
             KeyAscii = 0
         End If
   End Select
End Sub
Private Sub txtCaseNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      If txtCaseNo(1) <> "L" And txtCaseNo(1) <> "LA" And txtCaseNo(1) <> "CFL" And _
         txtCaseNo(1) <> "FCL" And txtCaseNo(1) <> "" And txtCaseNo(1) <> "LIN" And txtCaseNo(1) <> "ACS" Then
         MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
         TextInverse txtCaseNo(1)
         Cancel = True
      End If
   End If
End Sub
