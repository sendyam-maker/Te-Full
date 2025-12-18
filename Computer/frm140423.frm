VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm140423 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人編號匯出案件統計及互惠狀況"
   ClientHeight    =   2628
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2628
   ScaleWidth      =   6180
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      ItemData        =   "frm140423.frx":0000
      Left            =   30
      List            =   "frm140423.frx":0002
      TabIndex        =   6
      Top             =   1392
      Width           =   6100
   End
   Begin VB.TextBox txtFileName 
      Height          =   264
      Left            =   30
      TabIndex        =   3
      Top             =   1104
      Width           =   5650
   End
   Begin VB.CommandButton CmdOpenFile 
      Caption         =   "<="
      Height          =   250
      Left            =   5805
      TabIndex        =   2
      Top             =   1104
      Width           =   345
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Default         =   -1  'True
      Height          =   405
      Left            =   5280
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "匯出"
      Height          =   405
      Left            =   4440
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "1."
      ForeColor       =   &H00FF0000&
      Height          =   888
      Left            =   36
      TabIndex        =   5
      Top             =   216
      Width           =   4236
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注意事項："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   960
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm140423"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2025/06/11 ---Memo by Lydia 2025/09/11 (114/9/15)公告
Option Explicit
Dim strExtension As String '副檔名

Private Sub CmdChk_Click()

    If txtFileName = MsgText(601) Then
       MsgBox "檔案不可空白！"
       txtFileName.SetFocus
       Exit Sub
    Else

       If Dir(txtFileName) = "" Then
          MsgBox txtFileName & vbCrLf & "檔案不存在！", vbCritical
          Exit Sub
       Else
          If PUB_ChkFileOpening(txtFileName) = True Then
             Exit Sub
          End If
       End If
    End If

    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    If RunExcelChk = True Then MsgBox "匯出已完成！"
    Me.Enabled = True
    Screen.MousePointer = vbDefault

End Sub


Private Sub cmdExit_Click()
     Unload Me
End Sub

Private Sub CmdOpenFile_Click()
    Dim stFileName As String
    Dim sFile
On Error GoTo ErrHnd
  
    stFileName = ""
    strExtension = "" '副檔名
    With CommonDialog1
        .CancelError = True
        .FileName = stFileName
        .Filter = "Excel檔案 (*.xls 或 *.xlsx)|*.xls;*.xlsx"
        .Filter = "Excel檔案 (*.xls 或 *.xlsx)"
        '選過的路徑
        If PUB_GetLastDate(Me.Name, "Dir") <> "" Then
            .InitDir = PUB_GetLastDate(Me.Name, "Dir")
        Else
            .InitDir = PUB_Getdesktop
        End If
        .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .ShowOpen
        If .FileName <> "" Then
            txtFileName.Text = .FileName
            If InStr(.FileName, "\") > 0 Then
               For intI = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), intI, 1) = "\" Then
                     '記錄選過的路徑
                     PUB_SaveLastDate Me.Name, "Dir", Mid(Trim(.FileName), 1, intI - 1)
                     Exit For
                  End If
               Next intI
            End If
            '記錄副檔名,避免匯入之 xlsx 檔案另存成 xls(格式可能與xls無法相容,出現相容性檢查訊息)會彈錯誤
            If Right(.FileName, 5) = ".xlsx" Then
                strExtension = ".xlsx"
            Else
                strExtension = ".xls"
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
   MoveFormToCenter Me
   Label4.Caption = "1.檔案中只會執行第一個Sheet資料(最左邊)" & vbCrLf & _
                    "2.只需在A欄輸入代理人/潛在客戶編號，輸入END或連續空白5格表示匯入結束。" & vbCrLf & _
                    "3.B欄回寫「案件往來」，C欄回寫「互惠狀況」。" & vbCrLf

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm140423 = Nothing
End Sub

Private Function RunExcelChk() As Boolean
Dim intR As Integer, intQ As Integer, intCounter As Integer, strANo As String
Dim strGrp As String, intS As Integer, strTmpB As String, strTmpC As String
Dim strF(), intWidth()
Dim rsRD As New ADODB.Recordset
Dim strMidCon As String
Dim strSqlNow As String, strSQLpass As String
Dim strSqlAreaNow As String, strSqlAreaPass As String
Dim xlsAp As New Excel.Application
Dim wksrpt As New Worksheet
Dim strNoList As String, intQuery As Integer
    
On Error GoTo ErrHnd
     
   strF = Array("編號", "案件往來", "互惠狀況")
   intWidth = Array(12, 30, 30)
   intCounter = 1
   
   RunExcelChk = False
   List1.Clear
   xlsAp.Visible = False
   List1.AddItem "開始：", 0
   xlsAp.Workbooks.Open txtFileName
   Set wksrpt = xlsAp.Worksheets(xlsAp.ActiveSheet.Name) '避免存錯工作表造成錯誤(多工作表且有資料)
   strANo = RTrim(Replace(Replace(UCase(xlsAp.Range("A" & intCounter)), "　", ""), " ", ""))
   'Excel A欄 有END或空白多行(intS)就離開
   Do While InStr(strANo & ",", "END") = 0
      If Len(strANo) >= 6 Then
         strANo = ChangeCustomerL(strANo)
         strExc(1) = ""
         If Left(strANo, 1) = "Y" Or Left(strANo, 1) = "X" Or Left(strANo, 1) = "R" Then
            If PUB_GetCustData(strANo) = True Then
               strExc(1) = Left(strANo, 1)
            Else
               wksrpt.Range("B" & intCounter).Value = "查無此編號"
               wksrpt.Range("C" & intCounter).Value = "查無此編號"
            End If
         End If
         If strExc(1) = "X" Or strExc(1) = "R" Then
            wksrpt.Range("B" & intCounter).Value = "NA"
            wksrpt.Range("C" & intCounter).Value = "NA"
         ElseIf strExc(1) = "Y" Then
            If strGrp = strANo Then
               If strGrp <> "" Then
                  wksrpt.Range("B" & intCounter).Value = strTmpB
                  wksrpt.Range("C" & intCounter).Value = strTmpC
               End If
            Else
               Call Pub_frm100114_6_StrMenu(strUserNum, Me.Name, strANo, strMidCon, strSqlNow, strSQLpass, strSqlAreaNow, strSqlAreaPass)
               intR = 1
               strTmpB = ""
               Set rsRD = ClsLawReadRstMsg(intR, strSqlNow)
               If intR = 1 Then
                  rsRD.MoveFirst
                  Do While Not rsRD.EOF
                     strTmpB = strTmpB & rsRD.Fields(0) & "-" & rsRD.Fields(1) & ";"
                     rsRD.MoveNext
                  Loop
               Else
                  strTmpB = "NA"
               End If
               strExc(0) = "SELECT fc01||fc02 AS f01,fc04,fc05,fc06||fc04||decode(fc05,'1','上半','2','下半')||'('||sum(fc07)||')' AS f02 " & _
                            ",fc04||decode(fc05,'1','上半','2','下半')||fc06||sum(fc07) as f03  FROM fagentconfig " & _
                            "WHERE fc01='" & Mid(strANo, 1, 8) & "' AND fc02='" & Mid(strANo, 9, 1) & "' " & _
                            "GROUP BY fc01||fc02,fc04,fc05,fc06,decode(fc05,'1','上半','2','下半') order by fc04 desc,fc05 asc "
               intR = 1
               strTmpC = ""
               Set rsRD = ClsLawReadRstMsg(intR, strExc(0))
               If intR = 1 Then
                  rsRD.MoveFirst
                  Do While Not rsRD.EOF
                     strTmpC = strTmpC & rsRD.Fields("F02") & ";"
                     rsRD.MoveNext
                  Loop
               Else
                  strTmpC = "NA"
               End If
               wksrpt.Range("B" & intCounter).Value = strTmpB
               wksrpt.Range("C" & intCounter).Value = strTmpC
            End If
         End If
         List1.AddItem "　" & strANo, 0
         intQ = intQ + 1
      End If
      If strGrp = strANo And strGrp = "" Then
         intS = intS + 1
         If intS > 5 Then
            Exit Do
         End If
      Else
         intS = 1
      End If
      strGrp = strANo
      If InStr(strNoList & ",", strGrp) = 0 And InStr(strGrp, "編號") = 0 And strGrp <> "" Then
         strNoList = strNoList & strGrp & ","
         intQuery = intQuery + 1
      End If
      intCounter = intCounter + 1
      strANo = RTrim(Replace(Replace(UCase(xlsAp.Range("A" & intCounter)), "　", ""), " ", ""))
   Loop
   
   '設定欄位名稱/欄寬
   For intI = 0 To UBound(strF)
       wksrpt.Range(Chr(65 + intI) & "1").Value = strF(intI)
       wksrpt.Range(Chr(65 + intI) & "1").ColumnWidth = intWidth(intI)
   Next intI
   wksrpt.Range(Chr(65) & "1:" & Chr(65 + UBound(strF)) & "1").Interior.ColorIndex = 44
   
   '自動換行
   wksrpt.Range("B2:C" & intCounter).WrapText = True
   List1.AddItem "匯出完成！", 0
   wksrpt.Range("A1").Select
   
   '查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = ";匯出編號:" & Mid(strNoList, 1, Len(strNoList) - 1)
   InsertQueryLog (intQuery)

   '另存
   If Val(xlsAp.Version) < 12 Then
       '一般活頁簿 (xlWorkbookNormal)
       xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xlsx", ""), ".xls", "") & "_" & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=-4143
   Else
       If strExtension = ".xlsx" Then
           '預設活頁簿 (xlWorkbookDefault)
           xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xlsx", ""), ".xls", "") & "_" & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=51
       Else
           'Excel 97-2003 活頁簿 (xlExcel8)
           xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xlsx", ""), ".xls", "") & "_" & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=56
       End If
   End If

   xlsAp.Workbooks.Close
   xlsAp.Quit
   Set xlsAp = Nothing
    
   RunExcelChk = True
   Exit Function
    
ErrHnd:
   List1.AddItem "匯出失敗！請通知電腦中心(" & Err.Description & ")", 0
   MsgBox "資料有誤！請洽電腦中心"
   If Val(xlsAp.Version) < 12 Then
       xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & "_" & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=-4143
   Else
       If strExtension = ".xlsx" Then
           '預設活頁簿 (xlWorkbookDefault)
           xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & "_" & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=51
       Else
           'Excel 97-2003 活頁簿 (xlExcel8)
           xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & "_" & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=56
       End If
   End If
   xlsAp.Workbooks.Close
   xlsAp.Quit
   Set xlsAp = Nothing
End Function

