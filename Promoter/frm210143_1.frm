VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210143_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "公告公文記錄"
   ClientHeight    =   5460
   ClientLeft      =   6090
   ClientTop       =   1550
   ClientWidth     =   9140
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9140
   Begin VB.TextBox textDate 
      Height          =   270
      Index           =   0
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2250
      Width           =   915
   End
   Begin VB.TextBox textDate 
      Height          =   270
      Index           =   1
      Left            =   2190
      MaxLength       =   7
      TabIndex        =   13
      Top             =   2250
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(9)"
      Height          =   315
      Index           =   9
      Left            =   1140
      TabIndex        =   9
      Top             =   1470
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(10)"
      Height          =   315
      Index           =   10
      Left            =   3780
      TabIndex        =   10
      Top             =   1470
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(11)"
      Enabled         =   0   'False
      Height          =   315
      Index           =   11
      Left            =   1140
      TabIndex        =   11
      Top             =   1830
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(8)"
      Height          =   315
      Index           =   8
      Left            =   3780
      TabIndex        =   8
      Top             =   1110
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(7)"
      Height          =   315
      Index           =   7
      Left            =   1140
      TabIndex        =   7
      Top             =   1110
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(6)"
      Height          =   315
      Index           =   6
      Left            =   6390
      TabIndex        =   6
      Top             =   750
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(5)"
      Height          =   315
      Index           =   5
      Left            =   3780
      TabIndex        =   5
      Top             =   750
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(4)"
      Height          =   315
      Index           =   4
      Left            =   1140
      TabIndex        =   4
      Top             =   750
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(3)"
      Height          =   315
      Index           =   3
      Left            =   1140
      TabIndex        =   3
      Top             =   390
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(2)"
      Height          =   315
      Index           =   2
      Left            =   6390
      TabIndex        =   2
      Top             =   30
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(1)"
      Height          =   315
      Index           =   1
      Left            =   3780
      TabIndex        =   1
      Top             =   30
      Width           =   2505
   End
   Begin VB.CheckBox ChkPLF 
      Caption         =   "ChkPLF(0)"
      Height          =   315
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   30
      Width           =   2505
   End
   Begin VB.CommandButton cmdPLB 
      Caption         =   "顯示公文附件"
      Height          =   375
      Left            =   6300
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Height          =   375
      Left            =   5370
      TabIndex        =   14
      Top             =   2160
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm210143_1.frx":0000
      Height          =   2805
      Left            =   60
      TabIndex        =   21
      Top             =   2580
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   4957
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|系統類別|公告日期|主旨"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   1860
      X2              =   2460
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "公告日期："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   22
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "外　商："
      Height          =   315
      Left            =   150
      TabIndex        =   20
      Top             =   1530
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "法務處："
      Enabled         =   0   'False
      Height          =   315
      Left            =   150
      TabIndex        =   19
      Top             =   1890
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "商標處："
      Height          =   315
      Left            =   150
      TabIndex        =   18
      Top             =   810
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "專利處："
      Height          =   315
      Left            =   150
      TabIndex        =   17
      Top             =   90
      Width           =   915
   End
End
Attribute VB_Name = "frm210143_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Create By Sindy 2014/3/6
Option Explicit

' 變數宣告區
Dim m_AttachPath As String
Private Declare Function SendMessageByNum Lib "user32" _
   Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
   wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim ii As Integer

Private Sub cmdExit_Click()
   Unload Me
End Sub

'顯示公文附件
Private Sub cmdPLB_Click()
Dim hLocalFile As Long
Dim stFileName As String
Dim ii As Integer
      
   Screen.MousePointer = vbHourglass
   For ii = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(ii, 0) = "V" And GRD1.TextMatrix(ii, 4) <> "" Then
         If GetAttachFile(stFileName, GRD1.TextMatrix(ii, 4), GRD1.TextMatrix(ii, 2)) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
      End If
   Next ii
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery_Click()
   Call QueryData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strQueryLimit As String, strSQLCon As String
Dim Cancel As Boolean
   
   If ChkPLF(0).Value = 0 _
      And ChkPLF(1).Value = 0 _
      And ChkPLF(2).Value = 0 _
      And ChkPLF(3).Value = 0 _
      And ChkPLF(4).Value = 0 _
      And ChkPLF(5).Value = 0 _
      And ChkPLF(6).Value = 0 _
      And ChkPLF(7).Value = 0 _
      And ChkPLF(8).Value = 0 _
      And ChkPLF(9).Value = 0 _
      And ChkPLF(10).Value = 0 Then
'      And ChkPLF(11).Value = 0 Then    'cancel by sonia 2022/11/11 杜協理通知關閉法務處
      MsgBox "請至少點選一項價目表！", vbExclamation
      Exit Function
   End If
   
   Cancel = False
   Call textDate_Validate(0, Cancel)
   If Cancel = True Then
      Exit Function
   End If
   Cancel = False
   Call textDate_Validate(1, Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   '系統類別
   For ii = 0 To 10   'modify by sonia 2022/11/11 杜協理通知關閉法務處 11->10
      If ChkPLF(ii).Enabled = True And ChkPLF(ii).Value = 1 Then
         strQueryLimit = strQueryLimit & "," & Format(ii + 1, "00")
      End If
   Next ii
   strQueryLimit = Mid(strQueryLimit, 2)
   strQueryLimit = Replace(strQueryLimit, ",", "','")
   
   If textDate(0) <> "" Then
      strSQLCon = strSQLCon & " and PLB02>=" & DBDATE(textDate(0))
   End If
   If textDate(1) <> "" Then
      strSQLCon = strSQLCon & " and PLB02<=" & DBDATE(textDate(1))
   End If
   
   strSql = "SELECT ' ',decode(PLB01,'01','國內專利','02','大陸專利','03','香港澳門專利','04','CFP','05','國內商標','06','大陸商標','07','馬德里商標','08','國內著作權','09','大陸著作權','10','CFT','11','美國著作權','12','顧問及法務',PLB01),sqldatet(PLB02),PLB03,PLB01 FROM pricelistbulletin" & _
            " WHERE PLB01 in('" & strQueryLimit & "')" & strSQLCon & _
            " order by PLB01,PLB02 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   SetGrd
   If rsTmp.RecordCount = 0 Then
      GRD1.Rows = 2
      GRD1.row = 1
      GRD1.col = 0
      MsgBox "無此資料", vbOKOnly, "查詢資料"
      QueryData = False
   Else
      QueryData = True
   End If
   
   'rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_AttachPath = App.path '& Pub_GetSpecMan("EmpFlowAttPath")
   
   SetGrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   Set frm210143_1 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\$$*.pdf"
   End If
End Sub

Private Function GetAttachFile(ByRef pFileName As String, ByVal strKEY01 As String, ByVal strKEY02 As String, _
                               Optional pSavePath As String, Optional pFileSize As Integer = 0) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
   strExc(0) = "select * from pricelistbulletin where PLB01='" & strKEY01 & "' and PLB02=" & Val(DBDATE(strKEY02))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      pFileName = SetFileName(RsTemp.Fields("PLB01"), RsTemp.Fields("PLB02"))
      
      If pSavePath = "" Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
         stAttPath = m_AttachPath & "\" & pFileName
         '檔案已存在時
         If Dir(stAttPath) <> "" Then
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(stAttPath) = True Then
               MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
               Exit Function
            End If
            Kill stAttPath
         End If
      Else
         stAttPath = pSavePath
      End If
      
      If Dir(stAttPath) <> "" Then Kill stAttPath
      
      'Add By Sindy 2017/5/31
      If "" & RsTemp.Fields("plb13") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("plb13"), stAttPath, UCase("PRICELISTBULLETIN"))
      Else
      '2017/5/31 END
         With RsTemp
            lngSize = Val(.Fields("PLB04").Value)
            ReDim bytes(lngSize)
            If lngSize > 0 Then bytes() = .Fields("PLB05").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      pFileName = stAttPath
      If pFileSize = 1 Then
         pFileName = pFileName & " (" & Round(RsTemp.Fields("PLB04") / 1024, 2) & " KB)"
      End If
      GetAttachFile = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'檔案名稱：公告日期(民國日期)＋系統類別中文名稱＋價目表公告公文.pdf
Private Function SetFileName(strSysKind As String, strDate As String) As String
   Select Case strSysKind
      Case "01"
         strSysKind = "國內專利"
      Case "02"
         strSysKind = "大陸專利"
      Case "03"
         strSysKind = "香港澳門專利"
      Case "04"
         strSysKind = "CFP"
      Case "05"
         strSysKind = "國內商標"
      Case "06"
         strSysKind = "大陸商標"
      Case "07"
         strSysKind = "馬德里商標"
      Case "08"
         strSysKind = "國內著作權"
      Case "09"
         strSysKind = "大陸著作權"
      Case "10"
         strSysKind = "CFT"
      Case "11"
         strSysKind = "美國著作權"
 '     Case "12"                      'cancel by sonia 2022/11/11 杜協理通知關閉法務處
 '        strSysKind = "顧問及法務"   'cancel by sonia 2022/11/11 杜協理通知關閉法務處
   End Select
   SetFileName = "$$" & TransDate(strDate, 1) & strSysKind & "價目表公告公文" & ServerTime & ".pdf"
End Function

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, X, Y, nCol, nRow
   If nRow < 0 Then nRow = 0
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j
   
   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      GRD1.Visible = False
      If Trim(GRD1.TextMatrix(tmpMouseRow, 4)) <> "" Then
         If Trim(GRD1.TextMatrix(tmpMouseRow, 0)) = "" Then '原白變灰藍
            GRD1.TextMatrix(tmpMouseRow, 0) = "V"
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = &HFFC0C0 '灰藍
            Next i
         Else '原灰藍變白
            GRD1.TextMatrix(tmpMouseRow, 0) = ""
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15) '白
            Next i
         End If
      End If
      GRD1.Visible = True
   End If
End Sub

'公告日期
Private Sub textDate_GotFocus(Index As Integer)
   InverseTextBox textDate(Index)
   CloseIme
End Sub
Private Sub textDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
Private Sub textDate_Validate(Index As Integer, Cancel As Boolean)
   If textDate(Index) <> "" Then
      If ChkDate(textDate(Index)) = False Then
          Call textDate_GotFocus(Index)
          Cancel = True
          Exit Sub
      End If
      Select Case Index
         Case 0
'            If textDate(Index) <> "" And textDate(Index + 1) = "" Then
'               textDate(Index + 1) = textDate(Index)
'            End If
         Case 1
            If RunNick2(textDate(Index - 1), textDate(Index)) Then
               Call textDate_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
      End Select
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("V", "系統類別", "公告日期", "主旨", "PLB01")
   arrGridHeadWidth = Array(200, 1500, 1000, 6000, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub
