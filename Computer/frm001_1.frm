VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm001_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "整批更新造字資料"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdDel 
      Caption         =   "清除"
      Height          =   465
      Left            =   4830
      TabIndex        =   13
      Top             =   270
      Width           =   915
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3270
      Width           =   675
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3270
      Width           =   675
   End
   Begin VB.CommandButton cmdRunProc 
      Caption         =   "直接更新"
      Height          =   465
      Left            =   6270
      TabIndex        =   8
      Top             =   2130
      Width           =   1005
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "確認"
      Height          =   465
      Left            =   3690
      TabIndex        =   2
      Top             =   2130
      Width           =   915
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入"
      Height          =   465
      Left            =   3690
      TabIndex        =   0
      Top             =   270
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "離開"
      Height          =   405
      Left            =   6270
      TabIndex        =   3
      Top             =   300
      Width           =   1035
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "檢查"
      Height          =   465
      Left            =   3690
      TabIndex        =   1
      Top             =   1080
      Width           =   915
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6270
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "目前匯入：　　　　筆　　　　　　已確認：　　　　筆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   3300
      Width           =   6915
   End
   Begin VB.Label Label5 
      Caption         =   "現在啟動"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5190
      TabIndex        =   9
      Top             =   2220
      Width           =   1065
   End
   Begin VB.Label Label4 
      Caption         =   "3.確認後的資料，將於次日凌晨　　更新所有造字資料；"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2130
      Width           =   3525
   End
   Begin VB.Label Label3 
      Caption         =   "若需修正，請重新匯入Excel檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   390
      TabIndex        =   6
      Top             =   1620
      Width           =   3195
   End
   Begin VB.Label Label2 
      Caption         =   "2.檢查：會寄Excel檔，請收信後　檢查BIG5造字或Unicode字是否正確"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3525
   End
   Begin VB.Label Label1 
      Caption         =   "1.先將造字替換表Excel檔匯入DB"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   330
      Width           =   3525
   End
End
Attribute VB_Name = "frm001_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create table by Lydia 2022/03/21
'因為111/3/12 的批次更換造字是在M51-Win7執行exe, 程式直接開啟Excel檔,直接讀取Excel值動丟到String變數,再組合語法,
'然後Unicode呈現?造成大量資料的造字都變成「?」，之後3/14~3/17都在檢查和還原資料。
'發生原因尚未查清楚(ex.在M51-Win7開啟excel的字體看起來正常,excel的字型從明朝體改為細明體)
'現在將步驟改為：匯入DB=>檢查產生的Excel=>確認加入排程=> 每日批次 or (人工啟動)直接更新
'-----------------------------------------
Option Explicit

Dim intJ As Integer
Dim intQ As Integer, strQuery As String
Dim rsQuery As New ADODB.Recordset

Private Sub cmdDel_Click()
    If MsgBox("是否清除目前匯入(含已確認)的資料？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
        cnnConnection.Execute "delete from editeudclog where eel01 in (9999,9998) "
        GetRecSum
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdImPort_Click()
   Dim stFileName As String
   Dim sFile
   Dim fs, f
   Dim strFile As String
   Dim ii As Integer
   
On Error GoTo ErrHnd
   
   If Val(Text1(0)) > 0 Then
      If MsgBox("匯入將會清空之前的匯入資料，是否繼續執行？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
          Exit Sub
      End If
   End If
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.*"
      .Filter = "All Files *.*|(*.*)"
      '預設上一次的路徑
      If GetSetting("TAIE", "EUDCL", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "EUDCL", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      '去掉多選Or cdlOFNAllowMultiselect
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            '多選
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
             SaveSetting "TAIE", "EUDCL", UCase(Me.Name) & "Dir", sFile(0)
             '路徑排除&
             If InStr(CStr(sFile(0)), "&") > 0 Then
                  MsgBox CStr(sFile(0)) & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於路徑！", vbExclamation
                  Exit Sub
             End If
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Or InStr(CStr(sFile(ii)), "&") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名！", vbExclamation
                  Exit Sub
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
                '限制EXCEL檔
                If Right(UCase(stFileName), 4) = ".XLS" Or Right(UCase(stFileName), 5) = ".XLSX" Then
                Else
                     MsgBox "請選擇Excel檔！", vbExclamation
                     Exit Sub
                End If
            
                '檢查檔案是否正在使用中
                If PUB_ChkFileOpening(stFileName) = True Then
                     MsgBox stFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                     Exit Sub
                End If
                
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set f = fs.GetFile(stFileName)
                '檔案大小為 0 KB 有誤
                If f.Size = 0 Then
                   ShowMsg sFile(ii) & MsgText(9221)
                   Exit Sub
                End If
            Next ii

         Else '單選
             '路徑排除&
             strExc(1) = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
             If InStr(strExc(1), "&") > 0 Then
                  MsgBox strExc(1) & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於路徑！", vbExclamation
                  Exit Sub
             End If
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Or InStr(strFile, "&") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名！", vbExclamation
               Exit Sub
            End If
            
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "EUDCL", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            
            stFileName = .FileName
            '限制EXCEL檔
            If Right(UCase(stFileName), 4) = ".XLS" Or Right(UCase(stFileName), 5) = ".XLSX" Then
            Else
                 MsgBox "請選擇Excel檔！", vbExclamation
                 Exit Sub
            End If
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(stFileName) = True Then
                 MsgBox stFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                 Exit Sub
            End If
               
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               Exit Sub
            End If
         End If
      End If
   End With
   
   If stFileName <> "" Then
      If ProcImportExcel(stFileName, strExc(1)) = True Then
          MsgBox "匯入完成，" & IIf(strExc(1) <> "", vbCrLf & strExc(1), "") & vbCrLf & "請檢查資料！", vbInformation
      Else
          MsgBox "匯入失敗！" & vbCrLf & strExc(1), vbCritical
      End If
      GetRecSum
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox "匯入失敗：" & Err.Description
   End If
End Sub

Private Function ProcImportExcel(ByVal stFileName, ByRef strReturnMsg As String) As Boolean
Dim xlsSalesPoint
Dim wksrpt
Dim intRow As Integer
Dim strErrMsg As String
Dim strWd1
Dim strWd2
Dim intCounter As Integer, intCheck As Integer

    ProcImportExcel = False
       Set xlsSalesPoint = CreateObject("Excel.Application")
       '目前只做單檔匯入
        xlsSalesPoint.Workbooks.Open stFileName
        'xlsSalesPoint.Visible = False
        Set wksrpt = xlsSalesPoint.Worksheets(1)
        '把Excel的警告訊息關掉
        xlsSalesPoint.DisplayAlerts = False
        '檢查標題欄
        intRow = 1
        strReturnMsg = ""
        If Trim("" & wksrpt.Range("A" & intRow).Value) <> "更改前" Then
             strReturnMsg = strReturnMsg & "Excel格式檢查：A欄為更改前;"
             GoTo EXITSUB
        End If
        If Trim("" & wksrpt.Range("B" & intRow).Value) <> "內碼" Then
             strReturnMsg = strReturnMsg & "Excel格式檢查：B欄為內碼;"
             GoTo EXITSUB
        End If
        If Trim("" & wksrpt.Range("C" & intRow).Value) <> "更改後" Then
             strReturnMsg = strReturnMsg & "Excel格式檢查：C欄為更改後;"
             GoTo EXITSUB
        End If
        If Trim("" & wksrpt.Range("D" & intRow).Value) <> "內碼" Then
             strReturnMsg = strReturnMsg & "Excel格式檢查：D欄為內碼;"
             GoTo EXITSUB
        End If
        If Trim("" & wksrpt.Range("E" & intRow).Value) <> "處理備註" Then
             strReturnMsg = strReturnMsg & "Excel格式檢查：E欄為處理備註;"
             GoTo EXITSUB
        End If

        strErrMsg = ""
        cnnConnection.Execute "delete from editeudclog where eel01 in (9999,9998) "
        wksrpt.Range("A:E").Cells.Font.Name = "新細明體-ExtB" '統一字型
        intRow = intRow + 1
        Do While Trim(wksrpt.Range("A" & intRow).Value & wksrpt.Range("C" & intRow).Value) <> ""
            intCounter = intCounter + 1 '流水號
            strWd1 = Trim(PUB_StringFilter(wksrpt.Range("A" & intRow).Value))
            strWd2 = Trim(PUB_StringFilter(wksrpt.Range("C" & intRow).Value))
            If Len(strWd1) = 0 Then
               strErrMsg = strErrMsg & "更改前為空白;"
            End If
            If Len(strWd2) = 0 Then
               strErrMsg = strErrMsg & "更改後為空白;"
            End If
            If strWd1 = strWd2 Then
               strErrMsg = strErrMsg & "更改前後文字相同;"
            End If
            If Len(strWd1) <> Len(strWd2) And strWd1 <> "" And strWd2 <> "" Then
               strErrMsg = strErrMsg & "更改前後文字數量不同;"
            End If
            If strErrMsg = "" Then  '基本檢查沒有問題的筆數
               intCheck = intCheck + 1
            End If
            strSql = "insert into EditEudcLog (EEL01,EEL02,EEL03,EEL04,EEL05,EEL06,EEL07,EEL08,EEL09) VALUES " & _
                        "(9998, " & intCounter & ", '" & strUserNum & "', '" & strWd1 & "', '" & Trim(PUB_StringFilter("" & wksrpt.Range("B" & intRow).Value)) & "'," & _
                         " '" & strWd2 & "', '" & Trim(PUB_StringFilter("" & wksrpt.Range("D" & intRow).Value)) & "', null , " & CNULL(strErrMsg) & " ) "
            cnnConnection.Execute strSql
            strErrMsg = ""
            intRow = intRow + 1
        Loop
    xlsSalesPoint.Quit '離開
    Set wksrpt = Nothing
    Set xlsSalesPoint = Nothing
    
    If intCounter <> intCheck Then
        strReturnMsg = strReturnMsg & "匯入" & intCounter & "筆，基本檢查沒問題" & intCheck & "筆；"
    End If
    ProcImportExcel = True
    Exit Function
    
EXITSUB:
   xlsSalesPoint.Quit '離開
   Set wksrpt = Nothing
   Set xlsSalesPoint = Nothing

End Function

Private Sub cmdRunProc_Click()
   GetRecSum
   If Val(Text1(1)) = 0 Then
      MsgBox "查無資料可更新！"
      Exit Sub
   End If
   If MsgBox("目前資料庫" & pub_DbTerminalName & vbCrLf & "是否開始更新資料？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
       Screen.MousePointer = vbHourglass
          cmdRunProc.Enabled = False
          If Pub_BatchDay113Proc("2") = True Then
              Call PUB_BatchDay113Excel("2", "")
          End If
          GetRecSum
          cmdRunProc.Enabled = True
       Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub cmdSend_Click()
   Call PUB_BatchDay113Excel("1", "")
End Sub

Private Sub cmdUpd_Click()
    
    '將基本檢查不過的也列入，在最後執行另外處理 --- and eel09 is null
    strQuery = "select count(*) cnt from editeudclog where eel01 =9998 "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
    strExc(1) = "0"
    If intQ = 1 Then
        strExc(1) = "" & rsQuery.Fields("cnt")
    End If
    If Val(strExc(1)) = 0 Then
         MsgBox "查無資料可供確認!! ", vbInformation
    Else
         If MsgBox("可以確認的筆數：" & strExc(1) & "，是否繼續確認？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
              strSql = "Update  editeudclog set eel01=9999 where eel01 =9998 "
              cnnConnection.Execute strSql
         End If
         GetRecSum
    End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    GetRecSum
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm001_1 = Nothing
End Sub

Private Sub GetRecSum()
    Text1(0) = "0": Text1(1) = "0"
    '匯入=9998
    strExc(0) = "select count(*) cnt from editeudclog where eel01 in (9999,9998) "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       Text1(0) = "" & RsTemp.Fields("cnt")
    End If
    '確認=9998
    strExc(0) = "select count(*) cnt from editeudclog where eel01 in (9999) "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       Text1(1) = "" & RsTemp.Fields("cnt")
    End If
End Sub


