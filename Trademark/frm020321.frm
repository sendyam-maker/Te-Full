VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm020321 
   BorderStyle     =   1  '單線固定
   Caption         =   "台灣商標延展開拓(貝爾)"
   ClientHeight    =   3855
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6000
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "frm020321.frx":0000
      Top             =   3150
      Width           =   5835
   End
   Begin VB.TextBox text1 
      Height          =   315
      Index           =   2
      Left            =   1590
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2370
      Width           =   645
   End
   Begin VB.FileListBox File2 
      Height          =   270
      Left            =   630
      TabIndex        =   12
      Top             =   150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Text            =   "C:\temp\貝爾商標\商標圖檔"
      Top             =   1740
      Width           =   3795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   345
      Left            =   5400
      TabIndex        =   3
      Top             =   1740
      Width           =   345
   End
   Begin VB.CommandButton cmdWord 
      Cancel          =   -1  'True
      Caption         =   "產生定稿(&W)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1260
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5265
      Left            =   0
      TabIndex        =   8
      Top             =   3870
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9287
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<="
      Height          =   345
      Left            =   5400
      TabIndex        =   1
      Top             =   1290
      Width           =   345
   End
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1590
      TabIndex        =   0
      Text            =   "C:\temp\貝爾商標\貝爾文字檔.txt"
      Top             =   1290
      Width           =   3795
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   756
   End
   Begin VB.CommandButton cmdImPort 
      Caption         =   "匯入資料(&E)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "分機："
      Height          =   180
      Left            =   1020
      TabIndex        =   15
      Top             =   2430
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "服務人員："
      Height          =   180
      Left            =   660
      TabIndex        =   14
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1590
      TabIndex        =   13
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標圖檔案路徑："
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "注意：當程式正在執行〔產生定稿〕時，請暫時不要使用Word！"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   2910
      Width           =   5835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "文字檔案路徑："
      Height          =   180
      Left            =   300
      TabIndex        =   9
      Top             =   1350
      Width           =   1260
   End
End
Attribute VB_Name = "frm020321"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/22 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Create By Sindy 2012/7/13
Option Explicit

Dim p_Recs1 As Integer
Dim m_WordFilePath As String
Dim m_intFileCnt As Integer, m_iRow As Integer
Dim bolRetry As Boolean '是否已發生錯誤且重試
Dim m_AppName As String '商標註冊人
Dim m_AppAddrZip As String '申請人地址郵遞區號
Dim m_AppAddr As String '申請人地址

'加入代表圖用
'Const msoBringInFrontOfText = 4
'Const msoFalse = 0
'Const msoLineSolid = 1
'Const msoLineSingle = 1
Const msoTrue = -1
'Const msoPictureAutomatic = 1
Dim intHeight As Integer, intCnt As Integer
Dim ff3 As Integer, m_PrintRpt3 As Boolean, m_strFileName3 As String 'Add By Sindy 2013/1/28


Private Sub ResetGrid(ByRef p_Grid As MSHFlexGrid, Index As Integer)
   With p_Grid
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      If Index = 0 Then
         .FormatString = "審定號數|商標名稱|專用權人|郵遞區號|專用權人地址|專用期限|是否為本所案件|商標圖檔名"
      End If
   End With
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdImPort_Click()
Dim i As Integer, j As Integer
Dim intRow As Integer
Dim fs, f
Dim strText As String, strTemp As String, strErrText As String
Dim intTab As Integer
Dim strChkTxt As Variant
Dim bolExecuteWord As Boolean
Dim intQ As Integer 'Add By Sindy 2018/3/15
   
On Error GoTo ErrHnd
   
   If txtPath1.Text = "" Then
      MsgBox "文字檔案路徑不可空白！", vbExclamation
      txtPath1.SetFocus
      Exit Sub
   End If
   If Dir(txtPath1) = "" Then
      MsgBox "檔案不存在！", vbExclamation
      txtPath1.SetFocus
      Exit Sub
   End If
   
   bolExecuteWord = False
   'Modify By Sindy 2018/3/15
'   If MsgBox("匯入資料後，要一併執行[產生定稿]嗎？", vbExclamation + vbYesNo) = vbYes Then
   intQ = MsgBox("匯入資料後，要一併執行[產生定稿]嗎？" & vbCrLf & vbCrLf & _
             "（按取消：放棄執行）", vbExclamation + vbYesNoCancel + vbDefaultButton3, "重要訊息！")
   If intQ = vbYes Then '一併產生定稿
      bolExecuteWord = True
      
      If txtPath2.Text = "" Then
         MsgBox "商標圖檔案路徑不可空白！", vbExclamation
         txtPath2.SetFocus
         Exit Sub
      End If
      If InStr(txtPath2, ".") > 0 Then
         For i = Len(txtPath2) To 1 Step -1
            If Mid(txtPath2, i, 1) = "\" Then
               txtPath2 = Mid(txtPath2, 1, i - 1)
               Exit For
            End If
         Next i
      Else
         If Right(txtPath2, 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
      End If
      File2.Refresh
      If File2.ListCount = 0 Then
         MsgBox "找不到商標圖檔！"
         Exit Sub
      End If
      If text1(2).Text = "" Then
         MsgBox "分機不可空白！", vbExclamation
         text1(2).SetFocus
         Exit Sub
      End If
   ElseIf intQ = vbNo Then '僅匯入作業
   Else '放棄執行
      Exit Sub
   End If
   '2018/3/15 END
   
   p_Recs1 = 0
   intRow = 0
   m_PrintRpt3 = False 'Add By Sindy 2013/1/28
   
   Screen.MousePointer = vbHourglass
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.OpenTextFile(txtPath1, ForReading, TristateFalse)
   ResetGrid MSHFlexGrid1, 0
   
   cnnConnection.BeginTrans
   
   '清除暫存檔資料
   strSql = "delete from BaireTrademark"
   cnnConnection.Execute strSql

   Do While f.AtEndOfLine <> True
      intRow = intRow + 1
      strText = f.ReadLine
      strErrText = strText 'Add By Sindy 2018/12/4
      If intRow > 1 And Left(strText, InStr(strText, vbTab) - 1) <> "" Then
         p_Recs1 = p_Recs1 + 1
         MSHFlexGrid1.Rows = p_Recs1 + 1
         For i = 0 To 5
            intTab = InStr(strText, vbTab)
            If i = 5 Then '最後一個欄位
               strTemp = Trim(strText)
            Else
               strTemp = Trim(Mid(strText, 1, intTab - 1))
               strText = Mid(strText, intTab + 1, Len(strText))
            End If
            If i = 0 Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 6) = strTemp '商標圖檔名
               If Left(strTemp, 1) = "T" Or Left(strTemp, 1) = "S" Then
                  strTemp = Trim(Right(strTemp, Len(strTemp) - 1))
               End If
               MSHFlexGrid1.TextMatrix(p_Recs1, 0) = strTemp '審定號數
            End If
            If i = 1 Then MSHFlexGrid1.TextMatrix(p_Recs1, 1) = strTemp '商標名稱
            If i = 2 Then '專用權人
               If InStr(strTemp, "<") > 0 Then
                  strChkTxt = Split(strTemp, "<")
                  strTemp = Trim(strChkTxt(0))
               ElseIf InStr(strTemp, ",") > 0 Then
                  strChkTxt = Split(strTemp, ",")
                  strTemp = Trim(strChkTxt(0))
               End If
               MSHFlexGrid1.TextMatrix(p_Recs1, 2) = strTemp
            End If
            If i = 3 Then '郵遞區號
               If strTemp > "" Then
                  For j = 1 To Len(strTemp)
                     MSHFlexGrid1.TextMatrix(p_Recs1, 7) = MSHFlexGrid1.TextMatrix(p_Recs1, 7) & Chr(Asc(Mid(strTemp, j, 1))) ' - 23937
                  Next j
               End If
            End If
            If i = 4 Then MSHFlexGrid1.TextMatrix(p_Recs1, 3) = strTemp '專用權人地址
            If i = 5 Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 4) = DBDATE(strTemp) '專用期限
               '檢查是否為本所案件
               '審定號數先用原資料檢核,若Find不到資料,再看是否有須要補足碼數再檢核一次
'               If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "S" Then '服務商標
'                  'Modify By Sindy 2013/7/2 原為and TM08 in('4','5','6'),因T-086449及T-085452之故
'                  strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & MSHFlexGrid1.TextMatrix(p_Recs1, 0) & "' and TM10='000' and TM28='1' and (TM08 in('4','5','6') or instr(TM58,'原為服務標章')>0 or instr(TM58,'原為聯合服務標章')>0) and tm29 is null and tm57 is null"
'               Else '商標
               'Modify By Sindy 2013/8/13 閉卷不通知,銷卷要通知
               'Modify By Sindy 2013/8/15 銷卷也不通知
               If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "S" Or Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "T" Then
                  strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & MSHFlexGrid1.TextMatrix(p_Recs1, 0) & "' and TM10='000' and TM28='1' and TM08 in('1','2','3')"
               Else
                  MsgBox "審定號（" & MSHFlexGrid1.TextMatrix(p_Recs1, 0) & "）之商標種類（" & Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) & "）有問題，請確認！"
                  Exit Sub
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI <> 1 And Len(MSHFlexGrid1.TextMatrix(p_Recs1, 0)) < 8 Then
                  '補足8碼檢核
                  strTemp = Right("00000000" & MSHFlexGrid1.TextMatrix(p_Recs1, 0), 8)
'                  If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "S" Then '服務商標
'                     'Modify By Sindy 2013/7/2 原為and TM08 in('4','5','6'),因T-086449及T-085452之故
'                     strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & strTemp & "' and TM10='000' and TM28='1' and (TM08 in('4','5','6') or instr(TM58,'原為服務標章')>0 or instr(TM58,'原為聯合服務標章')>0) and tm29 is null and tm57 is null"
'                  Else '商標
                  If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "S" Or Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "T" Then
                     strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & strTemp & "' and TM10='000' and TM28='1' and TM08 in('1','2','3')"
                  End If
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     MSHFlexGrid1.TextMatrix(p_Recs1, 5) = "Y"
                  End If
               Else
                  MSHFlexGrid1.TextMatrix(p_Recs1, 5) = "Y"
               End If
               
               '商標名稱
               If MSHFlexGrid1.TextMatrix(p_Recs1, 1) = "@" Then
                  MSHFlexGrid1.TextMatrix(p_Recs1, 1) = MSHFlexGrid1.TextMatrix(p_Recs1, 2) + "標章"
               End If

               '新增資料至DB
'               If MSHFlexGrid1.TextMatrix(p_Recs1, 0) = "1073457" Then
'                  MsgBox "TEST"
'               End If
               
               'Add By Sindy 2014/4/18 跨類商標會資料重覆出現,因此過濾重覆出現的審定號數,只收錄一筆
               strExc(0) = "SELECT * FROM BaireTrademark WHERE bt07='" & MSHFlexGrid1.TextMatrix(p_Recs1, 6) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
               '2014/4/18 END
                  strSql = "insert into BaireTrademark(bt01,bt02,bt03,bt04,bt05,bt06,bt07,bt08) values(" & _
                           CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 0)) & "," & _
                           CNULL(ChgSQL(MSHFlexGrid1.TextMatrix(p_Recs1, 1))) & "," & _
                           CNULL(ChgSQL(MSHFlexGrid1.TextMatrix(p_Recs1, 2))) & "," & _
                           CNULL(ChgSQL(MSHFlexGrid1.TextMatrix(p_Recs1, 3))) & "," & _
                           CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 4)) & "," & _
                           CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 5)) & "," & _
                           CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 6)) & "," & _
                           CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 7)) & _
                           ")"
                  cnnConnection.Execute strSql
               End If
            End If
         Next i
      End If
   Loop
   cnnConnection.CommitTrans
   f.Close
   Screen.MousePointer = vbDefault
   'Add By Sindy 2013/1/28
   If m_PrintRpt3 = True Then
      Close ff3
      MsgBox "資料匯入完畢！新增時有錯誤資料，請至C:\temp\貝爾商標\" & m_strFileName3 & "查看"
   Else
   '2013/1/28 End
      If bolExecuteWord = True Then
         Call cmdWord_Click
      Else
         MsgBox "資料匯入完畢！"
      End If
   End If
   Exit Sub
   
ErrHnd:
   'Add By Sindy 2013/1/28
   If Err.Number = -2147217900 Then 'ORA-00917: 遺漏逗點
      '寫Log
      Call ReadTxt3(strSql)
      '接著發生錯誤陳述式的下個陳述式開始執行
      Resume Next
   End If
   '2013/1/28 End
   cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      'Modify By Sindy 2018/12/4
      MsgBox "第" & p_Recs1 & "筆," & vbCrLf & vbCrLf & strErrText & vbCrLf & vbCrLf & Err.Description, vbCritical
   End If
End Sub

'Add By Sindy 2013/1/28
'新增失敗記錄檔
Private Sub ReadTxt3(strSql As String)
   If m_PrintRpt3 = False Then
      m_PrintRpt3 = True
      If ff3 > 0 Then Close #ff3
      ff3 = FreeFile
      m_strFileName3 = "新增失敗記錄檔.txt"
      'Open PUB_Getdesktop & "\" & m_strFileName3 For Output As ff3
      Open "C:\temp\貝爾商標\" & m_strFileName3 For Output As ff3
   End If
   Print #ff3, strSql
End Sub

Private Sub Command1_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.txt"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "文字檔案 (*.txt)|*.txt"
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtPath1.Text = .FileName
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Command2_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.gif"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "GIF (*.GIF)"
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtPath2.Text = .FileName
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Label8.Caption = strUserName
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020321 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   CloseIme
   TextInverse text1(Index)
End Sub

Private Sub txtPath1_GotFocus()
   TextInverse txtPath1
End Sub

Private Sub txtPath2_GotFocus()
   TextInverse txtPath2
End Sub

Private Sub cmdWord_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strTime As String
Dim fs As Object
Dim i As Integer
Dim strCU01 As String, strTempAppAddr As String, strTempAppAddrZip As String
   
On Error GoTo ErrHnd
   
   strTime = time()
   
   If txtPath2.Text = "" Then
      MsgBox "商標圖檔案路徑不可空白！", vbExclamation
      txtPath2.SetFocus
      Exit Sub
   End If
   If InStr(txtPath2, ".") > 0 Then
      For i = Len(txtPath2) To 1 Step -1
         If Mid(txtPath2, i, 1) = "\" Then
            txtPath2 = Mid(txtPath2, 1, i - 1)
            Exit For
         End If
      Next i
   Else
      If Right(txtPath2, 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   End If
   File2.Refresh
   If File2.ListCount = 0 Then
      MsgBox "找不到商標圖檔！"
      Exit Sub
   End If
   If text1(2).Text = "" Then
      MsgBox "分機不可空白！", vbExclamation
      text1(2).SetFocus
      Exit Sub
   End If
   
   m_WordFilePath = "c:\temp\貝爾商標\WordFile"
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.DeleteFolder m_WordFilePath, True
NotFolder76:
   fs.CreateFolder m_WordFilePath
   
   '產生Word檔
   m_intFileCnt = 0
   bolRetry = True
   Screen.MousePointer = vbHourglass
   For i = 1 To 1
      'strSql = "select distinct bt03 from bairetrademark Where bt06 Is Null order by bt03 asc"
      'Modify By Sindy 2012/8/16 開拓函除了過濾本所案件,還要排除國內商標公報特定公司不列印者
      'Modify By Sindy 2013/3/1 + and tbnp08='T'
      strSql = "select distinct bt03 from bairetrademark,tmbulletinnp Where bt06 Is Null and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null order by bt03 asc"
      '2012/8/16 End
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         For m_iRow = 1 To rsTmp.RecordCount
            '一舜科技股 / " & rsTmp.Fields(0) & " / 七寶旅行社股份有限公司
            'Modify By Sindy 改以專用權人的審定號數小到大排序,抓最小號數的資料
            'strSql = "select * from bairetrademark Where bt06 Is Null and bt03='七寶旅行社股份有限公司' order by bt01 asc"
            'Modify By Sindy 2012/10/2 龔說要以專用權人+商標種類(商標前服務標章後)+審定號數小到大排序,抓最小號數的資料
            strSql = "select bairetrademark.*,substr(bt07,1,1) as T1 from bairetrademark Where bt06 Is Null and bt03='" & rsTmp.Fields(0) & "' order by T1 desc,bt01 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               m_AppName = "" & RsTemp.Fields("bt03") '專用權人
               m_AppAddr = "" & RsTemp.Fields("bt04") '專用權人地址
               m_AppAddrZip = PUB_ChangeZIPToSir("" & RsTemp.Fields("bt08")) '郵遞區號
               '非個人, 抓相同客戶名稱之最小客戶編號,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
               strCU01 = "": strTempAppAddr = "": strTempAppAddrZip = ""
               If Len(m_AppName) > 4 Then
                  strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     If RsTemp.RecordCount > 0 Then
                        RsTemp.MoveFirst
                        If RsTemp.Fields("cu02") <> "0" Then
                           strCU01 = "" & RsTemp.Fields("cu01")
                        Else
                           strTempAppAddrZip = Trim("" & RsTemp.Fields("cu30"))
                           strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
                        End If
                     End If
                  End If
                  If strCU01 <> "" Then
                     strSql = "SELECT * FROM customer WHERE cu01='" & strCU01 & "' and cu02='0'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        strTempAppAddrZip = Trim("" & RsTemp.Fields("cu30"))
                        strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
                     End If
                  End If
                  If strTempAppAddr <> "" Then
                     m_AppAddrZip = PUB_ChangeZIPToSir(strTempAppAddrZip)
                     m_AppAddr = strTempAppAddr
                  End If
               End If
               'Add By Sindy 2013/6/3
               If m_AppAddrZip = "" Then
                  m_AppAddrZip = PUB_ChangeZIPToSir(Left(PUB_AddrChangeZIPCode(m_AppAddr), 3))
               End If
               '2013/6/3 End
               '列印定稿
               If WordEdit() = False Then
                  GoTo ErrHnd
               End If
            End If
            
            If (m_iRow Mod 100) = 0 Or m_iRow = rsTmp.RecordCount Then
               g_WordAp.Documents.Save
               g_WordAp.Documents.Close
               bolRetry = True
            End If
'            If (m_iRow Mod 40) = 0 Then
'               Exit For
'            End If
            
            rsTmp.MoveNext
         Next m_iRow
      End If
   Next i
   If bolRetry = False Then
      g_WordAp.Documents.Save
      g_WordAp.Documents.Close
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
   End If
   
   Screen.MousePointer = vbDefault
   
   Set rsTmp = Nothing
   Set g_WordAp = Nothing
   MsgBox "作業完成！請至" & m_WordFilePath & "\資料夾中列印開拓函。（花費時間：" & strTime & "  " & time() & "）"
   Exit Sub
   
ErrHnd:
   If Err.Number = 76 Then
      GoTo NotFolder76
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Set g_WordAp = Nothing
   'Resume
End Sub

Private Function WordEdit() As Boolean
'+信頭
Dim stFileName As String
Dim iPicNo As Integer
Dim iPicNo2 As Integer
Dim oShape

'Added by Morgan 2020/3/30
If strSrvDate(1) >= 智慧所更名日 Then
   PUB_GetLetterPicID "1", "T", iPicNo, iPicNo2
Else
'end 2020/3/30
   iPicNo = 12
   iPicNo2 = 11
End If 'Added by Morgan 2020/3/30

'end
Dim rsTmp As New ADODB.Recordset
Dim i As Integer
   
On Error GoTo ERRORSECTION1
   
   WordEdit = True
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize 'wdWindowStateMinimize  wdWindowStateMaximize
   With g_WordAp
   
      If bolRetry = True Then
         m_intFileCnt = m_intFileCnt + 1
         g_WordAp.Documents.add.SaveAs m_WordFilePath & "\台灣商標延展開拓函" & Format(m_intFileCnt, "00") & ".doc"
      End If
   
      If bolRetry = False Then .Selection.InsertBreak Type:=wdPageBreak
      
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
            
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.Width = .CentimetersToPoints(21)
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = .CentimetersToPoints(0)
         oShape.Top = .CentimetersToPoints(0.5)
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.Width = .CentimetersToPoints(21)
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = .CentimetersToPoints(0)
            'oShape.Top = .CentimetersToPoints(27.3)
            oShape.Top = .CentimetersToPoints(27)
         End If
         .Selection.EndKey Unit:=wdStory
      End If
      
      '配合新的開窗定稿改固定行高
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
      .Selection.ParagraphFormat.LineSpacing = 15
      
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      '靠左
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      If m_AppAddrZip = "" Then
         .Selection.TypeParagraph
      End If
      .Selection.TypeText getAddrData
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "致：" & m_AppName
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "敬啟者："
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      'Modify By Sindy 改以專用權人的審定號數小到大排序,抓最小號數的資料
      'Modify By Sindy 2012/10/2 龔說要以專用權人+商標種類(商標前服務標章後)+審定號數小到大排序,抓最小號數的資料
      strSql = "select bairetrademark.*,substr(bt07,1,1) as T1 from bairetrademark Where bt06 Is Null and bt03='" & m_AppName & "' order by T1 desc,bt01 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         intHeight = 0: intCnt = 0
         For i = 1 To rsTmp.RecordCount
            If i = 1 Then
               .Selection.TypeText "　　貴公司／台 端所有註冊第" & rsTmp.Fields("bt01") & "號「" & rsTmp.Fields("bt02") & "」" & IIf(rsTmp.RecordCount > 1, "等" & rsTmp.RecordCount & "件", "") & "商標專用期限將於民國" & Left(ChangeWStringToTString(rsTmp.Fields("bt05")), 3) & "年" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 4, 2) & "月" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 6, 2) & "日屆滿，　貴公司／台 端若有意繼續使用前揭商標，即應辦理商標延展註冊。"
               .Selection.TypeParagraph
               .Selection.TypeText "　　商標專用期限屆滿，未辦理延展註冊者，商標權當然消滅，為維護　貴公司／台 端之商標權益，特函提醒　貴公司／台 端，請儘速與本所聯繫，本所將竭誠為　貴公司／台 端提供服務！"
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeText "　　　　耑此　　順頌"
               .Selection.TypeParagraph
               .Selection.TypeText "商　祺"
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
               '.Selection.TypeText "　　　　　　　　　　　　　　　台一國際專利商標事務所  敬上"
               .Selection.TypeText "　　　　　　　　　　　　　　　" & PUB_GetCompName2("1") & "  敬上"
               'end 2020/3/30
               .Selection.TypeParagraph
               .Selection.TypeText "　　　　　　　　　　　　　　　服務人員：" & Label8.Caption
               .Selection.TypeParagraph
               .Selection.TypeText "　　　　　　　　　　　　　　　服務專線：02-25061023" & IIf(text1(2).Text <> "", " 轉 " & text1(2), "")
               .Selection.TypeParagraph
            End If
            AddInPicToWordR g_WordAp, txtPath2 & "\" & rsTmp.Fields("bt07") & ".gif", i '插入圖檔
            rsTmp.MoveNext
         Next i
      End If
   End With
   bolRetry = False
   Exit Function
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            If bolRetry = False Then
               bolRetry = True
               Resume
            End If
         'Add By Sindy 2013/1/28
         Case 5152
            Resume Next
         '2013/1/28 End
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
            WordEdit = False
      End Select
   End If
End Function

Private Function getAddrData() As String
Dim strAddrData As String
Dim m_line As Variant
Dim ii As Integer
   
   '地址
   If m_AppAddr = "" Then
      m_AppAddr = String(20, "　")
   Else
      m_AppAddr = ToWide(Trim(CheckStr(m_AppAddr)))
   End If
   '收件人
   m_AppName = Trim(CheckStr(m_AppName))
   If m_AppAddrZip <> "" Then
      strAddrData = m_AppAddrZip & vbCrLf & m_AppAddr & vbCrLf & m_AppName & "　鈞啟"
   Else
      strAddrData = m_AppAddr & vbCrLf & m_AppName & "　鈞啟"
   End If
   If strAddrData <> "" Then
      m_line = Split(strAddrData, vbCrLf)
      For ii = 0 To UBound(m_line)
         strAddrData = m_line(ii)
         Do While strAddrData <> StrToStr(strAddrData, 17)
               If InStr(1, m_line(ii), StrToStr(strAddrData, 17)) = 1 Then
                   m_line(ii) = Mid(m_line(ii), 1, InStr(1, m_line(ii), StrToStr(strAddrData, 17)) - 1) & StrToStr(strAddrData, 17) & vbCrLf & Replace(m_line(ii), StrToStr(strAddrData, 17), "")
               Else
                   m_line(ii) = Mid(m_line(ii), 1, InStr(1, m_line(ii), StrToStr(strAddrData, 17)) - 1) & StrToStr(strAddrData, 17) & vbCrLf & Replace(Mid(m_line(ii), InStr(1, m_line(ii), StrToStr(strAddrData, 17))), StrToStr(strAddrData, 17), "")
               End If
               strAddrData = Replace(strAddrData, StrToStr(strAddrData, 17), "")
         Loop
      Next ii
      strAddrData = Join(m_line, vbCrLf)
      m_line = Split(strAddrData, vbCrLf)
      For ii = 0 To UBound(m_line)
           m_line(ii) = m_line(ii)
      Next ii
      strAddrData = Join(m_line, vbCrLf)
      m_line = Split(strAddrData, vbCrLf)
      If UBound(m_line) < 3 Then
           strAddrData = strAddrData & vbCrLf
      End If
   End If
   
   getAddrData = strAddrData
End Function

Private Sub AddInPicToWordR(ByRef oWord As Word.Application, strFileName As String, intFileCnt As Integer)
Dim oShape 'Added by Lydia 2016/09/29

   With oWord
      '筆數為3的倍數時,接下一頁
      'If (intFileCnt Mod 3) = 1 Then
      If intHeight > 453 Or intHeight = 0 Then
         intCnt = 1
         intHeight = 5
         .Selection.InsertBreak
      Else
         intCnt = intCnt + 1
      End If
      
      '插入圖片檔案
      'Modified by Lydia 2016/09/29 用舊寫法會造成Word2010出錯
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=strFileName, LinkToFile:=False, SaveWithDocument:=True
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:= _
      'strFileName, LinkToFile:= _
      'False, SaveWithDocument:=True
      '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
      Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=strFileName, LinkToFile:=False, SaveWithDocument:=True)
      oShape.Select
      
      '定義大小
      '鎖定最高 圖區
      '圖大小
      'Modified by Lydia 2016/09/29
      '.Selection.ShapeRange.LockAspectRatio = msoTrue
      '.Selection.ShapeRange.Line.Visible = True '加框線
      oShape.LockAspectRatio = msoTrue
      oShape.Line.Visible = True '加框線
      '移到指定位置
      '3個圖檔高度:210
      '4個圖檔高度:155
      'Modified by Lydia 2016/09/29
      'If Selection.ShapeRange.Height > 200 Then
      '   Selection.ShapeRange.Height = 200 '210
      If oShape.Height > 200 Then
         oShape.Height = 200
      End If
      If intCnt = 1 Then
         intHeight = 5
      Else
         intHeight = intHeight + 6
      End If
      'Modified by Lydia 2016/09/29
      '.Selection.ShapeRange.Top = intHeight
      '.Selection.ShapeRange.Left = 5 'Add By Sindy 2012/9/3
      'intHeight = intHeight + Selection.ShapeRange.Height
      oShape.Top = intHeight
      oShape.Left = 5
      intHeight = intHeight + oShape.Height
'      '3個圖檔起始位置:0/216/432
'      '4個圖檔起始位置:0/165/327/489
'      If (intFileCnt Mod 3) = 1 Then
'         .Selection.ShapeRange.Top = 0
'      ElseIf (intFileCnt Mod 3) = 2 Then
'         .Selection.ShapeRange.Top = intHeight + 7 '207 '216
'      Else
'         .Selection.ShapeRange.Top = (intHeight + 7) * 2 '414 '432
'      End If
      'Modified by Lydia 2016/09/29
      '.Selection.ShapeRange.LockAnchor = False
      oShape.LockAnchor = False
      
      .Selection.EndKey Unit:=wdStory
   End With
   Exit Sub
   
ErrHnd:
   Err.Raise Err.Number
End Sub
