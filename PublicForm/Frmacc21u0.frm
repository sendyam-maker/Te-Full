VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21u0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "帳單輸入-整批"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5736
   ScaleWidth      =   9132
   Visible         =   0   'False
   Begin VB.TextBox textUser 
      Height          =   285
      Left            =   1485
      MaxLength       =   6
      TabIndex        =   25
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除"
      Height          =   345
      Left            =   8235
      TabIndex        =   23
      Top             =   930
      Width           =   705
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟"
      Height          =   345
      Index           =   1
      Left            =   2205
      TabIndex        =   18
      Top             =   930
      Width           =   705
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟"
      Height          =   345
      Index           =   0
      Left            =   6210
      TabIndex        =   17
      Top             =   930
      Width           =   705
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   900
      TabIndex        =   13
      Top             =   5400
      Width           =   4995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入"
      Height          =   345
      Left            =   6930
      TabIndex        =   12
      Top             =   930
      Width           =   705
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "重整(&R)"
      Height          =   345
      Left            =   6255
      TabIndex        =   11
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   345
      Left            =   7200
      TabIndex        =   10
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入(&T)"
      Height          =   345
      Left            =   5310
      TabIndex        =   9
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   8115
      TabIndex        =   8
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   8685
      Picture         =   "Frmacc21u0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   540
      Width           =   350
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1905
      TabIndex        =   5
      Top             =   540
      Width           =   6795
   End
   Begin VB.Frame Frame2 
      Caption         =   "待輸入帳單："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3945
      Left            =   4140
      TabIndex        =   2
      Top             =   1020
      Width           =   4890
      Begin VB.CheckBox Check1 
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1275
         TabIndex        =   3
         Top             =   0
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3555
         Left            =   45
         TabIndex        =   4
         Top             =   270
         Width           =   4770
         _ExtentX        =   8424
         _ExtentY        =   6265
         _Version        =   393216
         Cols            =   4
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|匯入日期|檔案名稱|檔案大小"
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
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "待匯入檔案："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3525
      Left            =   135
      TabIndex        =   0
      Top             =   1020
      Width           =   3960
      Begin VB.CheckBox Check2 
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   1
         Top             =   0
         Width           =   705
      End
      Begin MSForms.ListBox lstImport 
         Height          =   3180
         Left            =   72
         TabIndex        =   28
         Top             =   288
         Width           =   3828
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "6752;5609"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   135
      TabIndex        =   15
      Top             =   4890
      Width           =   8895
      Begin VB.TextBox txtProgressBar 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         Height          =   300
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   8820
      End
   End
   Begin MSForms.Label lblUserName 
      Height          =   255
      Left            =   2820
      TabIndex        =   27
      Top             =   210
      Width           =   1440
      VariousPropertyBits=   19
      Caption         =   "lblFM2"
      Size            =   "2540;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳單匯入人員："
      Height          =   180
      Left            =   180
      TabIndex        =   26
      Top             =   225
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "檔名規則：本所案號.PDF（ex.P105116.PDF）"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   3
      Left            =   135
      TabIndex        =   24
      Top             =   4590
      Width           =   3960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "已勾選筆數："
      Height          =   180
      Index           =   2
      Left            =   7425
      TabIndex        =   22
      Top             =   5460
      Width           =   1080
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   8550
      TabIndex        =   21
      Top             =   5460
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "總筆數："
      Height          =   180
      Index           =   1
      Left            =   6075
      TabIndex        =   20
      Top             =   5460
      Width           =   720
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   6840
      TabIndex        =   19
      Top             =   5460
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   5460
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳單PDF檔存放路徑："
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   1740
   End
End
Attribute VB_Name = "Frmacc21u0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/15 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、lblUserName ; Printer列印未改
'Create by Morgan 2016/4/11
Option Explicit

Dim m_AttachPath As String

'列印報表用---
Dim PLeft() As Integer
Dim strTemp() As String
Dim iNowLine As Integer
Dim iRowHeight As Integer

Dim strPrinter As String
Dim dblMaxWidth As Double
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
Dim oFiles As files
Dim oFile As File

Dim m_iCols As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
'Add By Sindy 2018/2/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
Public m_PrevForm As Form  '前一畫面
'2018/2/22 END
Dim arrImpResult() As String 'Added by Morgan 2025/8/15

Private Sub cmdDelete_Click()
   Dim iRecord As Integer
   Dim bolInTrans As Boolean
   Dim ii As Integer
   
On Error GoTo ErrHnd

   If Val(lblCount) = 0 Then MsgBox "請先點選要刪除的記錄！", vbExclamation: Exit Sub
   If MsgBox("是否確定要刪除？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then Exit Sub
   
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) = "V" Then
      
'Modified by Morgan 2016/8/5 因匯入區共用無法控制個人帳單故取消限制(建個人資料夾又嫌太多...矛盾)--玲玲
'         'Added by Morgan 2016/7/7 只能刪除自己的
'         If Pub_StrUserSt03 <> "M51" And .TextMatrix(ii, 4) <> strUserNum Then
'            MsgBox .TextMatrix(ii, 2) & " 是由 " & .TextMatrix(ii, 4) & " 匯入，只可由本人刪除!!!", vbExclamation
'            GoTo ErrHnd
'         End If
'         'end 2016/7/7
'end 2016/8/5

         cnnConnection.BeginTrans
         bolInTrans = True
         strSql = "Update acc152 set ayf01='X' where ayf01='U' and ayf02='" & ChgSQL(.TextMatrix(ii, 2)) & "'"
         cnnConnection.Execute strSql, iRecord
         If PUB_DelFtpFile2("X", " and ayf02='" & ChgSQL(.TextMatrix(ii, 2)) & "'", "ACC152") Then
            strSql = "delete acc152 where ayf01='X' and ayf02='" & ChgSQL(.TextMatrix(ii, 2)) & "'"
            Pub_SeekTbLog strSql 'Added by Morgan 2019/7/8
            cnnConnection.Execute strSql, iRecord
         Else
            Err.Raise 999, , " 刪除Ftp檔案失敗!!"
         End If
         cnnConnection.CommitTrans
         bolInTrans = False
         .TextMatrix(ii, 0) = "X"
         .RowHeight(ii) = 0
         lblTotal = Val(lblTotal) - 1
         lblCount = Val(lblCount) - 1
      End If
   Next
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      If bolInTrans Then cnnConnection.RollbackTrans
      'Modify By Sindy 2021/2/5 + & vbCrLf & "strSQL：" & strSQL
      MsgBox Err.Description & vbCrLf & "strSQL：" & strSql, vbCritical
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdImPort_Click()
   RefreshList
   ImportFile
   QueryData
End Sub

Private Sub cmdOpen_Click(Index As Integer)
   Dim stFileName As String
   Dim hLocalFile As Long
   Dim arrList() As String
   
   If Index = 0 Then
      If Val(lblCount) = 0 Then MsgBox "請點選要開啟的檔案！", vbExclamation: Exit Sub
      
      With MSHFlexGrid1
      For intI = 1 To .Rows - 1
         If .TextMatrix(intI, 0) = "V" Then
            If PUB_GetAttachFile_Invoice("U", .TextMatrix(intI, 2), m_AttachPath, stFileName) = True Then
               ShellExecute hLocalFile, "open", m_AttachPath & "\" & stFileName, vbNullString, vbNullString, 1
            End If
            Exit For
         End If
      Next
      End With
   Else
      If lstImport.ListCount > 0 And lstImport.ListIndex <> -1 Then
         'Modified by Morgan 2025/8/15
         'If lstImport.ItemData(lstImport.ListIndex) = 1 Then
         If Val(arrImpResult(lstImport.ListIndex)) = 1 Then
         'end 2025/8/15
            MsgBox "檔案已匯入！！", vbInformation
         Else
            arrList = Split(lstImport.List(lstImport.ListIndex), " ")
            'Modified by Morgan 2025/8/15
            'ShellExecute hLocalFile, "open", txtPath & "\" & arrList(0), vbNullString, vbNullString, 1
            PUB_OpenPdf txtPath & "\" & arrList(0)
            'end 2025/8/15
         End If
      End If
   End If
End Sub

Private Sub cmdPath_Click()
   Dim sBuffer As String
   
   sBuffer = PUB_GetFolder(Me.hWnd, txtPath, "請選擇帳單PDF檔存放路徑")
   If sBuffer <> "" Then
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", sBuffer
      txtPath.Text = sBuffer
      RefreshList
   End If
   
End Sub

Private Sub cmdPrint_Click()
   PUB_RestorePrinter cmbPrinter
   DoPrint
   PUB_RestorePrinter strPrinter
End Sub

Private Sub cmdQuery_Click()
   RefreshList
   QueryData
End Sub

Private Sub Command1_Click()
   Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   Dim strErr As String
   Dim stFileName As String, strAYF02 As String
   Dim hLocalFile As Long
   Dim iRow As Integer
   
   If Val(lblCount) = 0 Then MsgBox "請點選要輸入的檔案！", vbExclamation: Exit Sub
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         'strCP01 = "P" 'Removed by Morgan 2018/10/18 不必再限制系統別
         strAYF02 = .TextMatrix(iRow, 2)
         'Modified by Morgan 2025/8/19 + pChkLike=True
         If PUB_GetCaseNoFromFileName(strAYF02, strCP01, strCP02, strCP03, strCP04, strErr, True) = True Then
            'If PUB_ChkIsNoBillCase(strCP01, strCP02, strCP03, strCP04) = True Then Exit Sub 'Added by Morgan 2019/7/8 2019/10/25 先取消--郭
            'Add By Sindy 2018/2/26
            If m_strIR01 <> "" Then
               If strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 <> m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 Then
                  MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！", , MsgText(5)
                  Exit Sub
               End If
            End If
            '2018/2/26 END
            
            If PUB_GetAttachFile_Invoice("U", strAYF02, m_AttachPath, stFileName) = True Then
               ShellExecute hLocalFile, "open", m_AttachPath & "\" & stFileName, vbNullString, vbNullString, 1
            End If
            
            ToolShow
            tool1_enabled
            With Frmacc2150
            Set .m_ParentForm = Me
            .m_eFileName = strAYF02
            'Added by Sindy 2018/2/27
            If m_strIR01 <> "" Then
               .m_strIR01 = m_strIR01
               .m_strIR02 = m_strIR02
               .m_strIR03 = m_strIR03
               .m_strIR04 = m_strIR04
               .m_RDate = m_RDate
            End If
            '2018/2/27 END
            .m_CP01 = strCP01
            .m_CP02 = strCP02
            .m_CP03 = strCP03
            .m_CP04 = strCP04
            .Show
            End With
            .TextMatrix(iRow, 0) = "X"
            .RowHeight(iRow) = 0
            lblTotal = Val(lblTotal) - 1
            lblCount = Val(lblCount) - 1
            Me.Enabled = False
         End If
      End If
   Next
   End With
End Sub

Private Sub Form_Activate()
Dim bolSelect As Boolean
Dim iRow As Integer
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String 'Added by Morgan 2025/8/19
   
   strFormName = Name
   
   'Added by Sindy 2018/2/22
   If m_strIR01 <> "" And m_Done = False Then
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "＜" & m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 & "＞）"
   End If
   
   'Add By Sindy 2018/2/22
   With MSHFlexGrid1
   If m_strIR01 <> "" And .Rows - 1 > 0 And Me.Enabled = True Then
      bolSelect = False
      '點選此本所案號資料列
      .Visible = False
      For iRow = 1 To .Rows - 1
         'Modified by Morgan 2020/2/14 PUB_CaseNo2FileName
         'Modified by Morgan 2020/2/14 配合卷宗區檔名格式統一,CP02改抓全部
         'If UCase(.TextMatrix(iRow, 2)) = UCase(m_strCP01 & Val(m_strCP02) & IIf(m_strCP04 <> "00", "-" & m_strCP03 & "-" & m_strCP04, IIf(m_strCP03 <> "0", "-" & m_strCP03, "")) & ".pdf") Then
         'Modified by Morgan 2025/8/19
         'If UCase(.TextMatrix(iRow, 2)) = UCase(PUB_CaseNo2FileName(m_strCP01, m_strCP02, m_strCP03, m_strCP04) & ".pdf") Then
         Call PUB_GetCaseNoFromFileName(UCase(.TextMatrix(iRow, 2)), strCP01, strCP02, strCP03, strCP04, , True)
         If strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 = m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 Then
         'end 2025/8/19
            bolSelect = True
            .row = iRow
            .col = 0
            ClickGrid MSHFlexGrid1
            Exit For
         End If
      Next iRow
      .Visible = True
      If bolSelect = True Then
         'Call Command1_Click
         Command1.Value = True
      Else
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致，無相同案件資料！", , MsgText(5)
      End If
   End If
   End With
   '2018/2/22 END
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath1
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   dblMaxWidth = txtProgressBar.Width
   textUser = strUserNum
   textUser_Validate False
   
   '讀取前次設定路徑
   txtPath.Text = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   If txtPath <> "" Then
      If oFileSys.FolderExists(txtPath) = False Then
         MsgBox "帳單PDF檔存放路徑 [ " & txtPath & " ] 不存在，請重新設定！", vbCritical
         txtPath = ""
      'Added by Morgan 2025/8/15
      ElseIf m_strCP01 <> "" Then
         Call PUB_ImportInvoice(m_strCP01, m_strCP02, m_strCP03, m_strCP04, True, txtPath)
      'end 2025/8/15
      End If
   Else
      MsgBox "請先設定帳單PDF檔存放路徑！", vbCritical
   End If
   
   If txtPath <> "" Then
      RefreshList
      QueryData
   End If
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   KillTemp
   
   lblCount.BackStyle = 0
   lblTotal.BackStyle = 0
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'Added by Morgan 2025/8/19
   If txtPath <> "" Then
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", txtPath
   End If
   'end 2025/8/19
   
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   
   'Add By Sindy 2018/2/23
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
      End If
      If Not m_PrevForm Is Nothing Then
         Set m_PrevForm = Nothing
      End If
   End If
   '2018/2/23 END
   
   Set oFileSys = Nothing
   Set Frmacc21u0 = Nothing
End Sub

Private Sub RefreshList()
   Dim ii As Integer
    
   lstImport.Clear
   lstImport.ToolTipText = "" 'Added by Morgan 2023/4/19
   'Modified by Morgan 2025/8/15
   'File1.path = txtPath.Text
   'File1.Refresh
   'For ii = 0 To File1.ListCount - 1
   '   If UCase(Right(Trim(File1.List(ii)), 4)) = ".PDF" Then
   '      lstImport.AddItem Trim(File1.List(ii))
   '   End If
   'Next
   Set oFolder = oFileSys.GetFolder(txtPath.Text)
   Set oFiles = oFolder.files
   If oFiles.Count > 0 Then
      For Each oFile In oFiles
         If UCase(Right(oFile.Name, 4)) = ".PDF" Then
            lstImport.AddItem oFile.Name
         End If
      Next
      If lstImport.ListCount > 0 Then
         ReDim arrImpResult(lstImport.ListCount - 1)
         lstImport.ListIndex = 0
      End If
   End If
   'end 2025/8/15
End Sub

Private Sub QueryData()
   Dim stCon As String
   
   If textUser <> "" Then
      stCon = " and ayf07='" & textUser & "'"
   End If
   
   strExc(0) = "select '' V,sqldatet(ayf08) 匯入日期,ayf02 檔案名稱,Round(ayf03 / 1024, 2)||' KB' 檔案大小, AYF07" & _
      " from acc152 where ayf01='U'" & stCon & " order by ayf04,ayf05"
   intI = 1
   lblCount = 0
   lblTotal = 0
   With MSHFlexGrid1
   .FixedCols = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '若沒有資料時不可直接設定給 Grid 否則 MouseRow 會跑掉
   If intI = 1 Then
      Set .Recordset = RsTemp
      SetCmdEnabled True
      lblTotal = RsTemp.RecordCount
      SetGrid
   Else
      SetCmdEnabled False
      SetGrid True
   End If
   
   m_iCols = .Cols
   End With
End Sub

Private Function ImportFile() As Boolean
   Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, strCP09 As String, strErr As String
   Dim strCaseNo As String
   Dim iTotRows As Integer
   Dim ii As Integer
   Dim dblFCnt As Double
   Dim stSaveName As String
   Dim bolUploadDone As Boolean
   
On Error GoTo ErrHnd

   If IsEmptyText(txtPath) = True Then
      MsgBox "請選擇帳單PDF檔存放路徑！", vbOKOnly, "檢核資料"
      cmdPath.SetFocus
      Exit Function
   ElseIf oFileSys.FolderExists(txtPath) = False Then
      MsgBox "帳單PDF檔存放路徑不存在，請重新選擇！"
      cmdPath.SetFocus
      Exit Function
   ElseIf Dir(txtPath & "\*.pdf") = "" Then
      MsgBox "資料夾 " & txtPath.Text & " 中沒有pdf檔！"
      cmdPath.SetFocus
      Exit Function
   End If
   
   RefreshList
   
   txtProgressBar.Width = 0
   dblFCnt = lstImport.ListCount
   For ii = 0 To lstImport.ListCount - 1
      strErr = "": strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = "": strCP09 = ""
      'strCP01 = "P" 'Removed by Morgan 2018/10/18 不必再限制系統別
      'Modified by Morgan 2025/8/15 +檔名可忽略後面非本所號部分
      If PUB_GetCaseNoFromFileName(lstImport.List(ii), strCP01, strCP02, strCP03, strCP04, strErr, True) = True Then
         'Added by Morgan 2019/7/8
         'Removed by Morgan 2019/7/8 2019/10/25 先取消--郭
         'If PUB_ChkIsNoBillCase(strCP01, strCP02, strCP03, strCP04) = True Then
         '   strErr = "取消匯入"
         'Else
         'end 2019/7/8
         
            stSaveName = PUB_CaseNo2FileName(strCP01, strCP02, strCP03, strCP04) & Mid(lstImport.List(ii), InStr(lstImport.List(ii), "."))
            'AddRecord lstImport.List(ii), txtPath, stSaveName, strErr
            Set oFile = oFileSys.GetFile(txtPath & "\" & lstImport.List(ii))
            If oFile.Name <> stSaveName Then oFile.Name = stSaveName 'Added by Morgan 2025/8/15 檔案含Unicode無法上傳，先更名
            
            If PUB_UploadInvoice(oFile, stSaveName, strErr) Then
               oFile.Delete True
            End If
            
         'End If 'Added by Morgan 2019/7/8
      End If
      If strErr <> "" Then
         lstImport.List(ii) = lstImport.List(ii) & " ..." & UCase(strErr)
      Else
         lstImport.List(ii) = lstImport.List(ii) & " ...成功"
         'Modified by Morgan 2025/8/15
         'lstImport.ItemData(ii) = 1
         arrImpResult(ii) = 1
         'end 2025/8/15
      End If
      
      SetListScroll lstImport
      txtProgressBar.Width = ii * (dblMaxWidth / dblFCnt): DoEvents
   Next
   txtProgressBar.Width = ii * (dblMaxWidth / dblFCnt): DoEvents
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Function AddRecord(pFileName As String, pFromPath As String, pSaveName As String, Optional pErr As String) As Boolean
   Dim stSQL As String, iRecords As Integer
   Dim bolInTrans As Boolean
   Dim stFullPath As String
   Dim stFtpPath As String

On Error GoTo ErrHand
   stFullPath = pFromPath & "\" & pFileName
   Set oFile = oFileSys.GetFile(stFullPath)
   
   cnnConnection.BeginTrans
   bolInTrans = True
   
   stSQL = "update ACC152 set ayf01=ayf01 where ayf01='U' and upper(ayf02)='" & ChgSQL(UCase(pSaveName)) & "'"
   cnnConnection.Execute stSQL, iRecords
   If iRecords > 0 Then
      Err.Raise 999, , " 檔名重複!!"
   End If
   
   If PUB_PutFtpFile(stFullPath, strSrvDate(1), pSaveName, stFtpPath, "ACC152") Then
      stSQL = "insert into ACC152(ayf01,ayf02,ayf03,ayf04,ayf05,ayf06,ayf07,ayf08,ayf09) values('U','" & ChgSQL(pSaveName) & "'," & oFile.Size & "," & Format(oFile.DateLastModified, "YYYYMMDD") & "," & Format(oFile.DateLastModified, "HHMMSS") & ",'" & ChgSQL(stFtpPath) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'))"
      cnnConnection.Execute stSQL, iRecords
   Else
      Err.Raise 999, , " 檔案上傳失敗!!"
   End If

   cnnConnection.CommitTrans
   oFile.Delete
   AddRecord = True
   
ErrHand:
   If Err.Number <> 0 Then
      'Modify By Sindy 2021/2/5 + & vbCrLf & "stSQL：" & stSQL
      pErr = Err.Description & vbCrLf & "stSQL：" & stSQL
      If bolInTrans Then cnnConnection.RollbackTrans
   End If
End Function

Private Sub DoPrint()
   Dim ii As Integer, jj As Integer
   Dim strFontName As String
   
   If Check1.Value <> vbChecked And Check2.Value <> vbChecked Then
      MsgBox "請勾選要列印的內容！", vbInformation
      Exit Sub
   End If
   
   strFontName = Printer.FontName
   Printer.FontName = "細明體"
   
   '待輸入帳單
   If Check1.Value = 1 Then
      GetPleft 1
      PrintTitle 1
      For jj = 1 To MSHFlexGrid1.Rows - 1
         If MSHFlexGrid1.TextMatrix(jj, 0) <> "N" Then
            For ii = 1 To 3
               strTemp(ii) = "" & MSHFlexGrid1.TextMatrix(jj, ii)
            Next ii
            If (iNowLine + 2) * iRowHeight > Printer.ScaleHeight Then
               Printer.NewPage
               PrintTitle 1  '列印表頭
            End If
            PrintDetail '列印明細
         End If
      Next jj
      Printer.EndDoc
   End If
   
   '匯入結果
   If Check2.Value = 1 Then
      GetPleft 2
      PrintTitle 2
      For jj = lstImport.ListCount - 1 To 0 Step -1
         strTemp(1) = lstImport.List(jj)
         If (iNowLine + 2) * iRowHeight > Printer.ScaleHeight Then
            Printer.NewPage
            PrintTitle 2 '列印表頭
         End If
         
         PrintDetail '列印明細
         
      Next jj
      Printer.EndDoc
   End If
   
   Printer.FontName = strFontName
End Sub

Sub GetPleft(Optional pIndex As Integer = 1)
   iRowHeight = 300
   If pIndex = 1 Then
      ReDim PLeft(1 To 5)
      ReDim strTemp(1 To 5)
      
      PLeft(1) = 500
      PLeft(2) = 2500
      PLeft(3) = 5000
      PLeft(4) = 6500
      PLeft(5) = 8000
   Else
      ReDim PLeft(1 To 1)
      ReDim strTemp(1 To 1)
      
      PLeft(1) = 500
   End If
End Sub

Sub PrintTitle(Optional pIndex As Integer = 1)
Dim strTitle As String

iNowLine = 0
If pIndex = 1 Then
   strTitle = "待輸入帳單"
Else
   strTitle = "匯入結果"
End If

iNowLine = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTitle) / 2)
Printer.CurrentY = iNowLine * iRowHeight
Printer.Print strTitle

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

iNowLine = iNowLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iNowLine = iNowLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iNowLine = 5
If pIndex = 1 Then
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print MSHFlexGrid1.TextMatrix(0, 1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print MSHFlexGrid1.TextMatrix(0, 2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print MSHFlexGrid1.TextMatrix(0, 3)
Else
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print strTitle
End If
iNowLine = iNowLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iNowLine * iRowHeight
Printer.Print String(148, "-")
iNowLine = iNowLine + 1
End Sub

Sub PrintDetail()
   Dim ii As Integer
   
   For ii = 1 To UBound(PLeft)
      Printer.CurrentX = PLeft(ii)
      Printer.CurrentY = iNowLine * iRowHeight
      Printer.Print strTemp(ii)
   Next ii
   iNowLine = iNowLine + 1
End Sub

Private Sub SetListScroll(oList As Object)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
      
On Error GoTo ExitP

   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
   
ExitP:

End Sub

Private Sub SetCmdEnabled(pEnabled As Boolean)
'   cmdOK(3).Enabled = pEnabled
'   cmdOK(1).Enabled = pEnabled
'   cmdOpen(0).Enabled = pEnabled
'   cmdOpen(1).Enabled = pEnabled
'   cmdDelete.Enabled = pEnabled And m_bDelete
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGrd1HeadWidth
   Dim iUbound As Integer
   Dim iRow As Integer

   arrGrd1HeadWidth = Array(250, 900, 2300, 1000)
   iUbound = UBound(arrGrd1HeadWidth)
   
   With MSHFlexGrid1
   .Visible = False
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   '.FixedCols = 2
   .FormatString = "V|匯入日期|檔案名稱|檔案大小"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGrd1HeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   .Visible = True
   End With
End Sub

Private Sub lstImport_Click()
   lstImport.ToolTipText = lstImport.List(lstImport.ListIndex)  'Added by Morgan 2023/4/19
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long, iRow As Integer
   Dim stValue As String
     
   
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   
   If nCol < 0 Or nRow < 0 Then Exit Sub
   
   If nRow = 0 Then
      If nCol = 0 Then
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) = "" Then
               stValue = "V"
               Exit For
            '已刪除資料標示為 X
            ElseIf .TextMatrix(iRow, 0) = "V" Then
               stValue = ""
               Exit For
            End If
         Next
         
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) <> "X" Then
               If .TextMatrix(iRow, 0) <> stValue Then
                  .row = iRow
                  ClickGrid MSHFlexGrid1
               End If
            End If
         Next
      Else
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If
      End If
   Else
      .row = nRow
      ClickGrid MSHFlexGrid1
   End If
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(grdDataList As MSHFlexGrid)
   Dim iCol As Integer

   With grdDataList
   If .TextMatrix(grdDataList.row, 1) <> "" Then
      If .TextMatrix(.row, 0) = "V" Then
         lblCount = Val(lblCount) - 1
         .TextMatrix(.row, 0) = ""
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
          Next
      '已刪除資料標示為 X
      ElseIf .TextMatrix(.row, 0) = "" Then
         lblCount = Val(lblCount) + 1
         .TextMatrix(.row, 0) = "V"
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = &HFFC0C0
         Next
      End If
   End If
   End With
End Sub

Private Sub textUser_GotFocus()
   TextInverse textUser
End Sub

Private Sub textUser_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textUser_Validate(Cancel As Boolean)
   lblUserName = ""
   If textUser <> "" Then
      lblUserName = GetStaffName(textUser, True)
      If lblUserName = "" Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtPath_Validate(Cancel As Boolean)
   If txtPath <> "" Then
      If oFileSys.FolderExists(txtPath) = False Then
         MsgBox "帳單PDF檔存放路徑 [ " & txtPath & " ] 不存在，請重新設定！", vbCritical
         Cancel = True
      End If
   End If
End Sub
