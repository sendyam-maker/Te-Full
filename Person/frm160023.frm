VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160023 
   BorderStyle     =   1  '單線固定
   Caption         =   "Excel整批匯入刷卡記錄"
   ClientHeight    =   4815
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   9090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9090
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1485
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frm160023.frx":0000
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   7
      Top             =   4320
      Width           =   9015
      Begin VB.TextBox TextCnt 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   8970
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   315
      Left            =   8580
      TabIndex        =   1
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   1530
      Width           =   6855
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯入(&T)"
      Height          =   345
      Left            =   6480
      TabIndex        =   2
      Top             =   330
      Width           =   885
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   30
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   705
      Left            =   2040
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1244
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "檔案名稱                                                             "
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   1410
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   7560
      TabIndex        =   3
      Top             =   330
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   810
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Excel存放路徑："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   1560
      Width           =   1605
   End
End
Attribute VB_Name = "frm160023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/13 Form2.0已修改
'Create By Sindy 2021/6/4
Option Explicit

Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
Dim i As Integer, j As Integer


Private Sub cmdExcel_Click()
Dim xlsSalesPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim stFileName As String
Dim intMaxRow As Integer
Dim dblMaxWidth As Double
Dim intRow As String
Dim strDate As String, strTime As String, strID As String, strChkID As String
Dim intErrRow As Integer
   
On Error GoTo flgErr
   
   intErrRow = 0
   stFileName = txtPath1
   If Dir(stFileName) = "" Then
      MsgBox "請選擇一個Excel檔案！"
      Exit Sub
   End If
   
   '開檔
   Screen.MousePointer = vbHourglass
   xlsSalesPoint.Workbooks.Open stFileName
   'xlsSalesPoint.Visible = True
   Set wksrpt = xlsSalesPoint.Worksheets(1)
   
   '檢查總筆數
   intRow = 0
   Do While Trim(wksrpt.Range("A" & (intRow + 1)).Value) <> ""
      intRow = intRow + 1
      '檢查標題
      If intRow = 1 Then
         If Trim(wksrpt.Range("A" & intRow).Value) <> "時間戳記" Then
            MsgBox "A 欄位必須是「時間戳記」！"
            GoTo RunExit
         End If
         If Trim(wksrpt.Range("B" & intRow).Value) <> "員工編號" Then
            MsgBox "B 欄位必須是「員工編號」！"
            GoTo RunExit
         End If
         If Trim(wksrpt.Range("C" & intRow).Value) <> "" And Trim(wksrpt.Range("C" & intRow).Value) <> "系統記錄" Then
            MsgBox "C 欄位有誤「C欄位是系統記錄使用」！"
            GoTo RunExit
         Else
            wksrpt.Range("C" & intRow).Value = "系統記錄"
         End If
      End If
   Loop
   intMaxRow = intRow - 1 '總筆數
   If intMaxRow <= 0 Then
      MsgBox "此Excel檔案，無內容可讀取！"
      GoTo RunExit
   End If
   
   '讀取明細資料
   intRow = 1
   Do While Trim(wksrpt.Range("A" & (intRow + 1)).Value) <> ""
      dblMaxWidth = 8820
      TextCnt.Width = 0
      intRow = intRow + 1
      TextCnt.Width = dblMaxWidth / intMaxRow * intRow
      
      If Trim(wksrpt.Range("C" & intRow).Value) <> "V" Then  'V:已匯入不須重覆執行
         If IsDate(Trim(wksrpt.Range("A" & intRow).Value)) = False Then
            wksrpt.Range("C" & intRow).Value = "非日期時間資料"
            intErrRow = intErrRow + 1
            GoTo ReadNext
         End If
         
         '日期
         strDate = Format(Trim(wksrpt.Range("A" & intRow).Value), "yyyymmdd")
         '時間
         strTime = Format(Trim(wksrpt.Range("A" & intRow).Value), "hhmmss")
         '員編
         strID = ""
         strChkID = UCase(Trim(wksrpt.Range("B" & intRow).Value))
         
         '檢查有無此同仁
         '員編
         strExc(0) = "select st01,st02 from staff" & _
                     " where st01='" & strChkID & "'" & _
                     " and st04='1'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount = 1 Then
               strID = RsTemp.Fields("st01")
            End If
         End If
         '姓名
         If strID = "" Then
            strExc(0) = "select st01,st02 from staff" & _
                        " where st02='" & strChkID & "'" & _
                        " and st04='1'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.RecordCount = 1 Then
                  strID = RsTemp.Fields("st01")
               End If
            End If
         End If
         '回寫無此同仁
         If strID = "" Then
            wksrpt.Range("C" & intRow).Value = "無此同仁"
            intErrRow = intErrRow + 1
            GoTo ReadNext
         End If
         
         'Add By Sindy 2021/6/9
         '檢查員工卡號指紋資料檔記錄是否存在
         strExc(0) = "select * from StaffCardData" & _
                     " where SCD01='" & strID & "'" & _
                     " and SCD02='" & strID & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            strSql = "INSERT INTO StaffCardData(SCD01,SCD02)" & _
                     " VALUES('" & strID & "','" & strID & "')"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         '2021/6/9 END
         
         '檢查記錄是否已存在
         strExc(0) = "select * from PollRecord" & _
                     " where PR01=" & strDate & _
                     " and PR02=" & strTime & _
                     " and PR03='" & strID & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            wksrpt.Range("C" & intRow).Value = "記錄已存在"
         Else
            '新增
            strSql = "INSERT INTO PollRecord(PR01,PR02,PR03,PR08)" & _
                     " VALUES(" & strDate & "," & strTime & ",'" & strID & "'" & _
                     ",999)"
            'Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
            wksrpt.Range("C" & intRow).Value = "V"
         End If
      End If
ReadNext:
   Loop
   TextCnt.Width = dblMaxWidth: DoEvents
   MsgBox "資料匯入完畢！ " & vbCrLf & vbCrLf & _
          "(共計 " & intMaxRow & " 筆, 錯誤有 " & intErrRow & " 筆)"
   
   '關閉
   '.SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & ".xls"
   xlsSalesPoint.Workbooks(1).Save 'FileName:=stFileName, FileFormat:=56
   
RunExit:
   xlsSalesPoint.Workbooks.Close
   '離開
   xlsSalesPoint.Quit
   Set wksrpt = Nothing
   Set xlsSalesPoint = Nothing
   Screen.MousePointer = vbDefault

   Exit Sub
   
flgErr:
   Screen.MousePointer = vbDefault
   '.SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & "_err.xls"
   xlsSalesPoint.Workbooks(1).Save 'FileName:=stFileName, FileFormat:=56
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksrpt = Nothing
   Set xlsSalesPoint = Nothing
   If Err.Number <> 0 Then
       MsgBox intRow & " 筆 : " & Err.Description
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "" '"*.xlsx"
      .Filter = "files (*.xlsx)|*.xlsx|files (*.xls)|*.xls|"
      If InStrRev(txtPath1.Text, "\") = 0 Then
         .InitDir = txtPath1.Text
      Else
         .InitDir = Mid(txtPath1.Text, 1, InStrRev(txtPath1.Text, "\") - 1) 'txtPath1.Text
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            txtPath1.Text = sFile(0) & "\" & sFile(1)
         Else
            txtPath1.Text = .FileName
         End If
      End If
   End With
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
Dim SeekPrintL As Integer
   
   MoveFormToCenter Me
   
   m_DefaultPrinter = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      'cmbPrinter2.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = m_DefaultPrinter Then
         SeekPrint = i
      End If
   Next i
   Set Printer = Printers(SeekPrint)
   
   If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
      txtPath1.Text = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
   Else
      txtPath1.Text = PUB_Getdesktop
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '記錄路徑
   If InStrRev(txtPath1.Text, "\") = 0 Then
      SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", txtPath1.Text
   Else
      SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(txtPath1.Text, 1, InStrRev(txtPath1.Text, "\") - 1)
   End If
   
   Set frm160023 = Nothing
End Sub

Private Sub txtPath1_GotFocus()
   InverseTextBox txtPath1
End Sub
