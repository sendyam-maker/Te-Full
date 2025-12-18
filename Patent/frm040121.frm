VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040121 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子收據匯入"
   ClientHeight    =   5544
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9072
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5544
   ScaleWidth      =   9072
   Begin VB.CommandButton cmdOK 
      Caption         =   "歸卷"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   3
      Left            =   6165
      TabIndex        =   14
      Top             =   288
      Width           =   840
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除"
      Height          =   345
      Left            =   7200
      TabIndex        =   13
      Top             =   288
      Width           =   840
   End
   Begin VB.TextBox txtPDFPath 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1455
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   5130
      Width           =   7485
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟"
      Height          =   345
      Index           =   0
      Left            =   7335
      TabIndex        =   10
      Top             =   708
      Width           =   705
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "列印(&P)"
      Height          =   345
      Index           =   1
      Left            =   8085
      TabIndex        =   9
      Top             =   708
      Width           =   885
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   945
      TabIndex        =   5
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   6930
      Picture         =   "frm040121.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   720
      Width           =   350
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   780
      TabIndex        =   3
      Top             =   4770
      Width           =   4995
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "重整"
      Height          =   345
      Left            =   4380
      TabIndex        =   2
      Top             =   288
      Width           =   840
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   " 匯入"
      Height          =   345
      Left            =   5280
      TabIndex        =   1
      Top             =   288
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   8085
      TabIndex        =   0
      Top             =   288
      Width           =   885
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   768
      Top             =   2472
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3624
      Left            =   96
      TabIndex        =   8
      Top             =   1092
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   6392
      _Version        =   393216
      Cols            =   9
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|本所案號|申請案號|收據號碼|金額|開立日期|檔案名稱|匯入日期時間|狀態"
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "程序人員："
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   900
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1056
      TabIndex        =   20
      Top             =   48
      Visible         =   0   'False
      Width           =   1800
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3175;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "CSV檔資料夾請勿共用，因匯入後會被清空!!!"
      ForeColor       =   &H000000FF&
      Height          =   228
      Left            =   132
      TabIndex        =   19
      Top             =   432
      Width           =   3756
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "已勾選筆數："
      Height          =   180
      Index           =   1
      Left            =   7335
      TabIndex        =   18
      Top             =   4830
      Width           =   1080
   End
   Begin VB.Label lblCount 
      Height          =   180
      Left            =   8460
      TabIndex        =   17
      Top             =   4830
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總筆數："
      Height          =   180
      Index           =   0
      Left            =   5985
      TabIndex        =   16
      Top             =   4830
      Width           =   720
   End
   Begin VB.Label lblTotal 
      Height          =   180
      Left            =   6750
      TabIndex        =   15
      Top             =   4830
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   45
      TabIndex        =   12
      Top             =   5190
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "CSV檔："
      Height          =   180
      Left            =   132
      TabIndex        =   7
      Top             =   780
      Width           =   696
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   180
      Left            =   60
      TabIndex        =   6
      Top             =   4830
      Width           =   720
   End
End
Attribute VB_Name = "frm040121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/30 Form2.0已修改 (無需修改)
'Created by Morgan 2017/2/17
Option Explicit

Const MAX_FILENAME_LEN = 260
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1

Dim m_bDelete As Boolean
Dim m_Sys As String
Dim m_iCols As Integer
Dim m_AttachPath As String
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim strPrinter As String
Dim oFileSys As New FileSystemObject
Dim oFile As File
Public m_ProState As String 'Add by Sindy 2020/8/10


Private Sub cmdDelete_Click()
   Screen.MousePointer = vbHourglass
   If Val(lblCount) = 0 Then
      MsgBox "請先勾選要刪除的資料！", vbExclamation
   ElseIf MsgBox("共有 " & Val(lblCount) & " 筆電子收據將刪除，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
      FormDelete
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function FormDelete() As Boolean
   Dim iRow As Integer, stNo As String
   
On Error GoTo ErrHnd
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      .row = iRow
      If .TextMatrix(.row, 0) = "V" Then
         cnnConnection.BeginTrans
         
On Error GoTo ErrHndT
         
         stNo = GetValue(.row, "收據號碼")
         
         PUB_DelFtpFile2 stNo 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
         
         strSql = "delete casepaperpdf where cpp01='" & stNo & "'"
         cnnConnection.Execute strSql, intI
         
         strSql = "delete ereceipt where er01='" & stNo & "'"
         cnnConnection.Execute strSql, intI
         
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "X"
         MSHFlexGrid1.RowHeight(MSHFlexGrid1.row) = 0
         cnnConnection.CommitTrans
         
         lblCount = Val(lblCount) - 1
         lblTotal = lblTotal - 1
         DoEvents
      End If
   Next
   End With
   
   
   FormDelete = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Sub cmdImPort_Click()
   Dim iImpRecs As Integer, iSkipRecs As Integer
   
   Screen.MousePointer = vbHourglass
   If txtPath = "" Then
      MsgBox "請先選擇CSV檔案!", vbExclamation
   ElseIf Dir(txtPath) = "" Then
      MsgBox "檔案不存在!", vbExclamation
   ElseIf Import2DB(iImpRecs, iSkipRecs) = True Then
      QueryData
      MsgBox "匯入完成共 " & iImpRecs & " 筆!!" & IIf(iSkipRecs > 0, "(已剔除 " & iSkipRecs & " 筆非P案)", ""), vbInformation
      KillAttach
   End If
   Screen.MousePointer = vbDefault
      
End Sub

Private Sub KillAttach()
   Dim stFolder As String
   
   Dir App.path '清除Dir指令對最後執行的資料夾的鎖定
   
   stFolder = Left(txtPath, InStrRev(txtPath, "\") - 1)
   oFileSys.DeleteFile stFolder & "\*.*", True
   oFileSys.DeleteFolder stFolder & "\*", True
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 3
      'Add By Sindy 2020/9/2
      If MsgBox("收據是否已列印，確定要歸卷了嗎？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      '2020/9/2 END
      MapReceipt
   Case 0
      Unload Me
   End Select
End Sub

Private Sub MapReceipt()
   Dim iRow As Integer, stNo As String, stSaveName As String
   Dim pa(1 To 4) As String, iColStatus As Integer
   Dim lAmt As Long 'Added by Morgan 2020/3/19
   Dim stER05 As String 'Add By Sindy 2020/8/25
   
On Error GoTo ErrHnd
   
   iColStatus = GetFieldId("狀態")
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1

      .row = iRow
      If .TextMatrix(.row, 0) = "V" Then

         stNo = GetValue(.row, "收據號碼")
         'Modify By Sindy 2020/8/25
         If m_ProState = "T" Then
            pa(1) = GetValue(.row, "tm01")
            pa(2) = GetValue(.row, "tm02")
            pa(3) = GetValue(.row, "tm03")
            pa(4) = GetValue(.row, "tm04")
         Else
         '2020/8/25 END
            pa(1) = GetValue(.row, "pa01")
            pa(2) = GetValue(.row, "pa02")
            pa(3) = GetValue(.row, "pa03")
            pa(4) = GetValue(.row, "pa04")
         End If
         lAmt = Val(GetValue(.row, "金額")) 'Added by Morgan 2020/3/19
         stER05 = GetValue(.row, "ER05") 'Add By Sindy 2020/8/25
         
         'Modified by Morgan 2018/12/19 +發文日,發文時間排序(同日可能有兩張收據)
         'Modified by Morgan 2020/3/20 +同發文日時發文金額相符的優先
         'Modify By Sindy 2020/8/25
'         If m_ProState = "T" And strSrvDate(1) < 20210105 Then
'            strExc(0) = "select c1.cp09,c1.cp10,decode(nvl(c1.cp84,c2.cp84)," & lAmt & ",1,2) Srt" & _
'               " from caseprogress c1,caseprogress c2" & _
'               " Where c1.cp01='" & pa(1) & "' and c1.cp02='" & pa(2) & "' and c1.cp03='" & pa(3) & "' and c1.cp04='" & pa(4) & "'" & _
'               " and c1.cp27>0 and c1.cp159=0 and instr(c1.cp64,'" & stER05 & "')>0" & _
'               " And Not Exists (Select * From CasePaperPdf Where cpp01=c1.cp09" & _
'               " And InStr(Upper(cpp02),'.RECEIPT.PDF')>0 and cpp10<>'D')" & _
'               " and c2.cp09(+)=c1.cp43" & _
'               " order by decode(substr(c1.cp09,1,1),'C',c2.cp27,c1.cp27) asc,Srt asc,c1.cp82"
'         Else
         '2020/8/25 END
            strExc(0) = "select c1.cp09,c1.cp10,decode(nvl(c1.cp84,c2.cp84)," & lAmt & ",1,2) Srt" & _
               " from caseprogress c1,letterprogress,caseprogress c2" & _
               " Where c1.cp01='" & pa(1) & "' and c1.cp02='" & pa(2) & "' and c1.cp03='" & pa(3) & "' and c1.cp04='" & pa(4) & "'" & _
               " and c1.cp27>0 and c1.cp159=0 and lp01(+)=c1.cp09 and lp03=0 And lp19='Y'" & _
               " And Not Exists (Select * From CasePaperPdf Where cpp01=lp01" & _
               " And InStr(Upper(cpp02),'.RECEIPT.PDF')>0 and cpp10<>'D')" & _
               " and c2.cp09(+)=c1.cp43"
            'Modify By Sindy 2021/1/4
            If m_ProState = "T" Then
               strExc(0) = strExc(0) & " and instr(c1.cp64,'" & stER05 & "')>0"
               'Modify By Sindy 2021/1/15 T-213680 申請+跨類 (+,c1.cp09 asc)
               strExc(0) = strExc(0) & " order by decode(substr(c1.cp09,1,1),'C',c2.cp27,c1.cp27) asc,Srt asc,c1.cp82 asc,c1.cp09 asc"
               '2021/1/15 END
            Else
               strExc(0) = strExc(0) & " order by decode(substr(c1.cp09,1,1),'C',c2.cp27,c1.cp27) asc,Srt asc,c1.cp82"
            End If
            '2021/1/4 END
'         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            stSaveName = PUB_CaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & "." & RsTemp("cp10") & ".RECEIPT.pdf"
            'Modify By Sindy 2021/1/19 + , stER05
            If UpdateData(stNo, RsTemp(0), stSaveName, stER05) = True Then
               MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "X"
               MSHFlexGrid1.RowHeight(MSHFlexGrid1.row) = 0
               lblCount = Val(lblCount) - 1
               lblTotal = lblTotal - 1
            Else
               .TextMatrix(.row, iColStatus) = "歸卷失敗"
            End If
         Else
            .TextMatrix(.row, iColStatus) = "無缺收據"
         End If
         
         DoEvents
      End If
   Next
   End With
   Exit Sub
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'Modify By Sindy 2021/1/19 + pER05 As String : 智慧局收文文號:1108006614-0
Private Function UpdateData(pReceiptNo As String, pCPP01 As String, pCPP02 As String, pER05 As String) As Boolean
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
   strSql = "update casepaperpdf set cpp01='" & pCPP01 & "',cpp02='" & pCPP02 & "',cpp10='Y' where cpp01='" & pReceiptNo & "'"
   cnnConnection.Execute strSql, intI
   
   strSql = "update ereceipt set er18='" & pCPP01 & "' where er01='" & pReceiptNo & "'"
   cnnConnection.Execute strSql, intI
   
   'Add By Sindy 2020/9/17 收據號碼:109DP099413;存入進度備註
   'Modified by Morgan 2022/10/27 CP64後面要加空白檢查否則NULL時不會更新
   strSql = "update caseprogress set cp64=replace(cp64||';收據號碼:" & pReceiptNo & ";',';;',';')" & _
            " where cp09='" & pCPP01 & "'" & _
            " and instr(cp64||' ','收據號碼:" & pReceiptNo & ";')=0"
   cnnConnection.Execute strSql, intI
   '2020/9/17 END
   
   PUB_UpdateLP03 pCPP01
      
   'Add By Sindy 2021/1/19 申請+超項費
   If m_ProState = "T" Then
      'Add By Sindy 2022/6/17
      strExc(0) = "select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & pCPP01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strCP01 = RsTemp.Fields("cp01")
         strCP02 = RsTemp.Fields("cp02")
         strCP03 = RsTemp.Fields("cp03")
         strCP04 = RsTemp.Fields("cp04")
      End If
      '2022/6/17 END
      
      '一併送件的其他的案件性質(例:超項費)更新不缺收據
      'Add By Sindy 2022/6/17 加入案號判斷
      strSql = "UPDATE letterprogress SET lp03=" & strSrvDate(1) & ",lp05=decode(lp04||lp10,'Y'," & strSrvDate(1) & ",lp05)" & _
               " WHERE lp03=0 And lp19='Y'" & _
               " AND lp01 in(select cp09 from caseprogress where cp09<>'" & pCPP01 & "' and instr(cp64,'" & pER05 & "')>0" & _
                           " and cp10 not in('101')" & _
                           " and cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "')" & _
               " AND exists(select lp01 from letterprogress where lp01='" & pCPP01 & "' And lp03>0 and lp19='Y')"
      cnnConnection.Execute strSql, intI
   End If
   '2021/1/19 END
   
   cnnConnection.CommitTrans
   UpdateData = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub cmdOpen_Click(Index As Integer)
   With MSHFlexGrid1
   For intI = 1 To .Rows - 1
      If .TextMatrix(intI, 0) = "V" Then
         Exit For
      End If
   Next
   
   If intI < .Rows Then
      RunPdf Index
   Else
      MsgBox "請至少點選一筆資料！", vbExclamation
   End If
   End With
End Sub


Private Sub RunPdf(iAct As Integer)
   Dim stFiles As String, stSavePath As String
   Dim stFileName As String
   Dim hLocalFile As Long
   Dim arrFileName() As String
   Dim idx As Integer, iRow As Integer
   
   Dim program_name As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim strPrinterName As String
   
   '列印
   If iAct = 1 Then
      program_name = txtPDFPath
      strPrinterName = cmbPrinter
      '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
      process_id = Shell(program_name, vbHide)
      process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   End If
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         With MSHFlexGrid1
         
         stFiles = ""
         stSavePath = ""
         If GetAttachFile(GetValue(iRow, "收據號碼"), stFiles, stSavePath) = True Then
            arrFileName = Split(stFiles, ";")
            For idx = LBound(arrFileName) To UBound(arrFileName)
               If arrFileName(idx) <> "" Then
                  stFileName = stSavePath & "\" & arrFileName(idx)
                  '列印
                  If iAct = 1 Then
                     PrintOnePdf program_name, " /n /t """ & stFileName & """ """ & strPrinterName & """"
                  '開啟
                  Else
                     ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
                  End If
               End If
            Next
         End If
         End With
      End If
   Next
   End With
   
   If iAct = 1 Then
      TerminateProcess process_handle, 0&
      CloseHandle process_handle
      MsgBox "列印完畢!"
   End If
End Sub

Private Function GetAttachFile(ByVal strCPP01 As String, Optional ByRef pFileName As String, Optional ByRef pSavePath As String) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte

On Error GoTo ErrHnd

   If pSavePath = "" Then
      '建立暫存資料夾
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      pSavePath = m_AttachPath
   End If
   
   strExc(0) = "select cpp01,cpp02 from casepaperpdf where cpp01='" & strCPP01 & "'" & IIf(pFileName <> "", " and cpp02='" & ChgSQL(pFileName) & "'", "")
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         stAttPath = pSavePath & "\" & .Fields("cpp02")
         If Dir(stAttPath) <> "" Then Kill stAttPath

         If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), pSavePath) = False Then
            Exit Function
         End If
         pFileName = IIf(pFileName = "", "", pFileName & ";") & .Fields("cpp02")
         .MoveNext
      Loop
      End With
      GetAttachFile = True
   End If
   Exit Function

ErrHnd:

   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

Private Sub cmdPath_Click()
   Dim stPath As String
   
   With cd1
   .Filter = "Supported files|*.csv"
   .FilterIndex = 0
   If txtPath = "" Then
      .InitDir = PUB_Getdesktop
   Else
      stPath = Left(txtPath, InStrRev(txtPath, "\") - 1)
      If PUB_ChkDir(stPath) = True Then
         .InitDir = stPath
      Else
         .InitDir = PUB_Getdesktop
      End If
   End If
   .ShowOpen
   If Trim(.FileName) <> "" Then
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", .FileName
      txtPath.Text = .FileName
   End If
   End With
End Sub

Private Sub cmdQuery_Click()
   QueryData
End Sub

Private Sub Combo1_Click()
   If Combo1.Tag <> Combo1 Then
      cmdQuery.Value = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   
   '讀取前次設定路徑
   txtPath = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   If txtPath <> "" Then
      strExc(1) = Left(txtPath, InStrRev(txtPath, "\"))
      strExc(0) = Dir(strExc(1) & "*.CSV")
      If strExc(0) <> "" Then
         txtPath = strExc(1) & strExc(0)
      End If
   End If
   
   'Added by Morgan 2025/2/12
   If m_ProState = "" And (Pub_StrUserSt03 = "P12" Or Pub_StrUserSt03 = "M51") Then
      Label4.Visible = True
      Combo1.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label4)
   End If
   'end 2025/2/12
   
   SetFileAssociation
   m_AttachPath = App.path & "\" & strUserNum
   QueryData
   
   If Val(lblCount) > 0 Then
      MsgBox "有前次匯入收據尚未歸卷，請先處理！", vbExclamation
   End If
End Sub

'Removed by Morgan 2019/7/9
'Private Function ReadTextFile(Optional pCharset As String = "big5") As String
'   Dim adoStream As ADODB.Stream
'   Dim var_String As Variant
'   Dim strFileName As String
'
'   strFileName = txtPath
'   Set adoStream = New ADODB.Stream
'   'adoStream.Charset = "UTF-8"
'   adoStream.Charset = pCharset
'   adoStream.Open
'   adoStream.LoadFromFile strFileName
'   ReadTextFile = adoStream.ReadText
'   adoStream.Close
'   Set adoStream = Nothing
'End Function

'去除多餘的符號
Private Function GetStr(ByVal pContent As String) As String
   pContent = Trim(pContent)
   '去除右邊逗號
   If Right(pContent, 1) = "," Then
      pContent = Left(pContent, Len(pContent) - 1)
   End If
   '去除左邊雙引號
   Do While Left(pContent, 1) = """"
      pContent = Mid(pContent, 2)
   Loop
   '去除右邊雙引號
   Do While Right(pContent, 1) = """"
      pContent = Left(pContent, Len(pContent) - 1)
   Loop
   GetStr = pContent
End Function

Private Sub Form_Unload(Cancel As Integer)
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   KillTemp
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm040121 = Nothing
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub QueryData()
Dim strText As String
Dim rsQuery As ADODB.Recordset
Dim mSeqNo As String, stVTBX As String

   'Add by Sindy 2020/8/24
   If m_ProState = "T" Then
      'Modified by Morgan 2025/2/13 +PID
      strExc(0) = "select '' V,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04) 本所案號" & _
         ",ER07 申請案號,ER01 收據號碼,ER03 金額,sqldatet(ER06) 開立日期,ER15 檔案名稱" & _
         ",(to_char(ER17,'yyyy')-1911)||to_char(ER17,'/mm/dd hh24:mi:ss') 匯入日期時間,'待歸卷' 狀態,ER16 匯入人員" & _
         ",ER01,tm01,tm02,tm03,tm04,er06,ER05,'' PID from EReceipt,trademark where ER18='C' and (tm12=ER07 or tm15=ER07) and tm01 is not null and er02='商標'" & _
         " union select '' V,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",ER07 申請案號,ER01 收據號碼,ER03 金額,sqldatet(ER06) 開立日期,ER15 檔案名稱" & _
         ",(to_char(ER17,'yyyy')-1911)||to_char(ER17,'/mm/dd hh24:mi:ss') 匯入日期時間,'待歸卷' 狀態,ER16 匯入人員" & _
         ",ER01,cp01,cp02,cp03,cp04,er06,ER05,'' PID from EReceipt,trademark,caseprogress where ER18='C' and tm12(+)=ER07 and tm01 is null and cp36(+)=ER07 and cp36 is not null and er02='商標'" & _
         " order by ER05 asc"
   Else
   '2020/8/24 END
      'Modified by Morgan 2018/2/23 申請號改抓前9碼,被舉發後的更正收據申請號會有N##,如 P114288(收據申請號:105206887N01)
      'Modified by Morgan 2018/4/26 同日有設計及衍生設計收據時，因申請號前9碼相同資料會重複抓，改為分兩句用Union語法
      'Modified by Morgan 2020/3/18 若申請號無基本檔時再抓cp30(變更後改請 Ex:P-122974)
      'Modified by Morgan 2025/2/13 +PID
      strExc(0) = "select '' V,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) 本所案號" & _
         ",ER07 申請案號,ER01 收據號碼,ER03 金額,sqldatet(ER06) 開立日期,ER15 檔案名稱" & _
         ",(to_char(ER17,'yyyy')-1911)||to_char(ER17,'/mm/dd hh24:mi:ss') 匯入日期時間,'待歸卷' 狀態,ER16 匯入人員" & _
         ",ER01,pa01,pa02,pa03,pa04,er06,ER05,'' PID from EReceipt,patent where ER18='C' and instr(ER07,'N')=0 and pa11(+)=ER07 and pa01 is not null and er02='專利'" & _
         " union select '' V,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) 本所案號" & _
         ",ER07 申請案號,ER01 收據號碼,ER03 金額,sqldatet(ER06) 開立日期,ER15 檔案名稱" & _
         ",(to_char(ER17,'yyyy')-1911)||to_char(ER17,'/mm/dd hh24:mi:ss') 匯入日期時間,'待歸卷' 狀態,ER16 匯入人員" & _
         ",ER01,pa01,pa02,pa03,pa04,er06,ER05,'' PID from EReceipt,patent where ER18='C' and instr(ER07,'N')>0 and substr(pa11(+),1,9)=substr(ER07,1,9) and er02='專利'" & _
         " union select '' V,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",ER07 申請案號,ER01 收據號碼,ER03 金額,sqldatet(ER06) 開立日期,ER15 檔案名稱" & _
         ",(to_char(ER17,'yyyy')-1911)||to_char(ER17,'/mm/dd hh24:mi:ss') 匯入日期時間,'待歸卷' 狀態,ER16 匯入人員" & _
         ",ER01,cp01,cp02,cp03,cp04,er06,ER05,'' PID from EReceipt,patent,caseprogress where ER18='C' and instr(ER07,'N')=0 and pa11(+)=ER07 and pa01 is null and cp30(+)=ER07 and er02='專利'" & _
         " order by 6,2"
   End If
   intI = 1
   lblCount = 0
   lblTotal = 0
   With MSHFlexGrid1
   .FixedCols = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '若沒有資料時不可直接設定給 Grid 否則 MouseRow 會跑掉
   If intI = 1 Then
      'Added by Morgan 2025/2/12
      If Combo1 <> "" Then
         Set rsQuery = PUB_CreateRecordset(RsTemp, , , 300, Me.Name, mSeqNo)
         With rsQuery
            .MoveFirst
            Do While Not .EOF
               .Fields("PID") = PUB_GetPHandler(.Fields("本所案號"))
               .MoveNext
            Loop
            .UpdateBatch
            
            stVTBX = "select R001 as " & .Fields(0).Name
            For intI = 2 To .Fields.Count
               stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
            Next
            stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
         End With
         strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      End If
      Combo1.Tag = Combo1
      'end 2025/2/12
      
      Set .Recordset = RsTemp
      SetCmdEnabled True
      lblTotal = RsTemp.RecordCount
      SetGrid
      CheckGrid MSHFlexGrid1, "V"
   Else
      SetCmdEnabled False
      SetGrid True
   End If
   m_iCols = .Cols
   End With
   
   'Add By Sindy 2021/9/17
   If m_ProState = "T" Then
      '檢查是否有案件的申請案號或註冊號是錯的
      strExc(0) = "select * from EReceipt" & _
                  " where ER18='C' and er02='商標'" & _
                  " and ER01 not in (select ER01 from EReceipt,trademark where ER18='C' and (tm12=ER07 or tm15=ER07) and er02='商標' and tm01 is not null)" & _
                  " and ER01 not in (select ER01 from EReceipt,trademark,caseprogress where ER18='C' and tm12(+)=ER07 and er02='商標' and tm01 is null and cp36(+)=ER07 and cp36 is not null)"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            If strText <> "" Then
               strText = strText & vbCrLf
            End If
            strText = strText & "本所案號= " & RsTemp.Fields("ER21") & " 智慧局來的號數= " & RsTemp.Fields("ER07")
            RsTemp.MoveNext
         Loop
         If strText <> "" Then
            strText = strText & vbCrLf & vbCrLf & "請檢查基本檔資料是否有問題? 改完後再重整(查詢), 再繼續操作!"
         End If
         MsgBox strText, vbExclamation
      End If
   End If
   '2021/9/17 END
End Sub

Private Sub SetCmdEnabled(pEnabled As Boolean)
   cmdOK(3).Enabled = pEnabled
   cmdOpen(0).Enabled = pEnabled
   cmdOpen(1).Enabled = pEnabled
   cmdDelete.Enabled = pEnabled And m_bDelete
End Sub

Private Function Import2DB(Optional pImpRecs As Integer, Optional pSkipRecs As Integer) As Boolean
   Dim strText As String
   Dim arrRow() As String
   Dim arrCell() As String
   Dim idx1 As Integer, idx2 As Integer
   Dim stSQL As String, stValues As String, intR As Integer
   Dim stER01 As String, stER07 As String, stER21 As String
   Dim bolIsP As Boolean
   Dim stFolder As String, stFileName As String
   Dim arrERidx(23) As Integer
   Dim adoRst As New ADODB.Recordset
   Dim arrColNames() As String
   Dim strNewCol As String
   Dim iRecs As Integer
   Dim arrVar As Variant, intVar As Integer
   Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
               
   Const cDelimiter As String = """,""" '欄位區隔符號
   pImpRecs = 0
   pSkipRecs = 0
   
On Error GoTo ErrHnd
      
   'Modified by Morgan 2019/7/9
   'strText = ReadTextFile
   strText = PUB_ReadTextFile(txtPath)
   'end 2019/7/9
   
   strText = Replace(strText, Chr(9), "")
   strText = Replace(strText, Chr(13) & Chr(10), Chr(10))
   arrRow = Split(strText, Chr(10))
   arrCell = Split(arrRow(LBound(arrRow)), cDelimiter)
   ReDim arrColNames(UBound(arrCell)) As String
   
   strNewCol = ""
   arrERidx(15) = -1
   For idx1 = LBound(arrCell) To UBound(arrCell)
      strText = GetStr(arrCell(idx1))
      Select Case strText
      Case "收據號碼"
         arrColNames(idx1) = "ER01": arrERidx(1) = idx1
         
      Case "案件種類"
         arrColNames(idx1) = "ER02": arrERidx(2) = idx1
      
      Case "金額"
         arrColNames(idx1) = "ER03": arrERidx(3) = idx1
         
      Case "繳費時間"
         arrColNames(idx1) = "ER04": arrERidx(4) = idx1
      
      Case "收發文號"
         arrColNames(idx1) = "ER05": arrERidx(5) = idx1
         
      Case "開立日期"
         arrColNames(idx1) = "ER06": arrERidx(6) = idx1
         
      Case "案號"
         arrColNames(idx1) = "ER07": arrERidx(7) = idx1
         
      Case "費用類別"
         arrColNames(idx1) = "ER08": arrERidx(8) = idx1
         
      Case "案件名稱"
         arrColNames(idx1) = "ER09": arrERidx(9) = idx1
         
      Case "繳費方式"
         arrColNames(idx1) = "ER10": arrERidx(10) = idx1
         
      Case "申請人"
         arrColNames(idx1) = "ER11": arrERidx(11) = idx1
         
      Case "年度"
         arrColNames(idx1) = "ER12": arrERidx(12) = idx1
      
      Case "專利證書號"
         arrColNames(idx1) = "ER13": arrERidx(13) = idx1
         
      Case "商標註冊號"
         arrColNames(idx1) = "ER14": arrERidx(14) = idx1
         
      Case "檔案名稱"
         arrColNames(idx1) = "ER15": arrERidx(15) = idx1
         
      Case "繳款人"
         arrColNames(idx1) = "ER19": arrERidx(19) = idx1
      
      Case "實際繳款人" 'Added by Morgan 2017/5/5
         arrColNames(idx1) = "ER20": arrERidx(20) = idx1
      
      Case "自訂案件編號" 'Added by Morgan 2017/9/26
         arrColNames(idx1) = "ER21": arrERidx(21) = idx1
         
      Case "收據種類" 'Added by Morgan 2017/9/26
         arrColNames(idx1) = "ER22": arrERidx(22) = idx1
      'Added by Morgan 2018/2/23 目前沒用,略過
      'Modfied by Morgan 2024/2/2 +案由資訊(E-Set下載的CSV檔欄位名稱不同)
      Case "案由", "案由資訊"
         arrColNames(idx1) = "ER23": arrERidx(23) = idx1 'Add by Sindy 2020/9/2
      Case Else
         strNewCol = strNewCol & strText & vbCrLf
      End Select
   Next
   
   If strNewCol <> "" Then
      If MsgBox("有新增欄位如下將不會匯入，是否確定要繼續？" & vbCrLf & vbCrLf & strNewCol, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   
   If UBound(arrRow) = LBound(arrRow) Then
      MsgBox "無資料!!" & vbCrLf & vbCrLf & arrCell(0), vbCritical, "匯入失敗"
      Exit Function
   End If
      
   'Add By Sindy 2023/3/31
   '先檢查申請案號空白和不一致時彈錯誤訊息;修正後才能匯入
   If m_ProState = "T" Then
      For idx1 = LBound(arrRow) + 1 To UBound(arrRow)
         If arrRow(idx1) <> "" Then
            arrCell = Split(arrRow(idx1), cDelimiter)
            
            stFolder = ""
            stValues = ""
            stER01 = GetStr(arrCell(arrERidx(1)))
            stER07 = GetStr(arrCell(arrERidx(7)))
            stER21 = GetStr(arrCell(arrERidx(21))) '案號
            If stER21 <> "" And Left(stER21, 1) = "T" Then
               arrVar = Split(stER21, "-") '自訂案件編號
               For intVar = LBound(arrVar) To UBound(arrVar)
                  If intVar = 0 Then strCP01 = arrVar(intVar)
                  If intVar = 1 Then strCP02 = arrVar(intVar)
                  If intVar = 2 Then strCP03 = arrVar(intVar)
                  If intVar = 3 Then strCP04 = arrVar(intVar)
               Next intVar
               If strCP01 <> "" And strCP02 <> "" Then
                  strCP02 = Format(strCP02, "000000")
                  strCP03 = Format(Val(strCP03), "0")
                  strCP04 = Format(Val(strCP04), "00")
                  If GetStr(arrCell(arrERidx(23))) = "新申請" Then
                     strExc(0) = "select tm12 from Trademark" & _
                                 " where TM01='" & strCP01 & "' And TM02='" & strCP02 & "' And TM03='" & strCP03 & "' And TM04='" & strCP04 & "'"
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If Trim("" & RsTemp.Fields("TM12")) <> Trim(stER07) Then
                           MsgBox "案號 " & strCP01 & strCP02 & strCP03 & strCP04 & " 申請案號有誤！" & vbCrLf & vbCrLf & _
                                  "系統裡是 " & Trim("" & RsTemp.Fields("TM12")) & "，但收據資料是 " & Trim(stER07) & vbCrLf & vbCrLf & _
                                  "請更正後再操作！", vbCritical, "資料有誤"
                           Exit Function
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Next
   End If
   '2023/3/31 END
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHndT

   For idx1 = LBound(arrRow) + 1 To UBound(arrRow)
      
      If arrRow(idx1) <> "" Then
         arrCell = Split(arrRow(idx1), cDelimiter)
         
         stFolder = ""
         stValues = ""
         stER01 = GetStr(arrCell(arrERidx(1)))
         stER07 = GetStr(arrCell(arrERidx(7)))
         stER21 = GetStr(arrCell(arrERidx(21)))
         
         'Modified by Morgan 2017/9/26 加判斷P案才要匯入
         bolIsP = True
         'Add by Sindy 2020/8/24
         If m_ProState = "T" Then
            If stER21 = "" Or Left(stER21, 1) <> "T" Then
               strExc(0) = "select * from trademark where tm12='" & stER07 & "' and tm01<>'T'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  bolIsP = False
               End If
            'Add By Sindy 2020/9/2
            Else
               arrVar = Split(stER21, "-") '自訂案件編號
               For intVar = LBound(arrVar) To UBound(arrVar)
                  If intVar = 0 Then strCP01 = arrVar(intVar)
                  If intVar = 1 Then strCP02 = arrVar(intVar)
                  If intVar = 2 Then strCP03 = arrVar(intVar)
                  If intVar = 3 Then strCP04 = arrVar(intVar)
               Next intVar
               If strCP01 <> "" And strCP02 <> "" Then
                  strCP02 = Format(strCP02, "000000")
                  strCP03 = Format(Val(strCP03), "0")
                  strCP04 = Format(Val(strCP04), "00")
                  '案件性質為申請時,更新申請案號
                  'Modify By Sindy 2023/3/31 取消更新,因實務上是程序人員先輸入”通知申請案號”的進度程序。
'                  If GetStr(arrCell(arrERidx(23))) = "新申請" Then
'                     strSql = "Update Trademark Set TM12='" & Trim(stER07) & "'" & _
'                              " Where TM01='" & strCP01 & "' And TM02='" & strCP02 & "' And TM03='" & strCP03 & "' And TM04='" & strCP04 & "'" & _
'                              " And TM12 is null"
'                     cnnConnection.Execute strSql, intR
'                  End If
               End If
            '2020/9/2 END
            End If
         Else
         '2020/8/24 END
            If stER21 = "" Or Left(stER21, 1) <> "P" Then
               strExc(0) = "select * from patent where pa11='" & stER07 & "' and pa01<>'P'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  bolIsP = False
               End If
            End If
         End If
         
         If bolIsP Then
            stSQL = "Insert into EReceipt("
            
            For idx2 = LBound(arrCell) To UBound(arrCell)
               strText = GetStr(arrCell(idx2))
               If arrColNames(idx2) <> "" Then
                  If stValues <> "" Then
                     stSQL = stSQL & ","
                     stValues = stValues & ","
                  End If
                  
                  stSQL = stSQL & arrColNames(idx2)
                  
                  '繳費時間(目前為西元年格式,但根據電子公文經驗有可能會改為民國年故也要考慮)
                  If arrColNames(idx2) = "ER04" Then
                     If InStr(strText, "/") = 4 Then
                        strText = (Val(Left(strText, 3)) + 1911) & Mid(strText, 4)
                     End If
                     stValues = stValues & "to_date('" & strText & "','yyyy/mm/dd hh24:mi:ss')"
                  Else
                     stValues = stValues & "'" & ChgSQL(strText) & "'"
                  End If
               End If
            Next
            
            stSQL = stSQL & ") values (" & stValues & ")"
            cnnConnection.Execute stSQL, intR
            pImpRecs = pImpRecs + 1
            '上傳檔案
            stFolder = Left(txtPath, InStrRev(txtPath, "\"))
            stFileName = GetStr(arrCell(arrERidx(15)))
            Set oFile = oFileSys.GetFile(stFolder & "\" & stFileName)
            SaveAttFile_PDF stER01, stFolder & "\" & stFileName, "$" & stER01 & ".pdf", Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True
         Else
            pSkipRecs = pSkipRecs + 1
         End If
      End If
   Next
   
   cnnConnection.CommitTrans
   Import2DB = True
   Set adoRst = Nothing
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description & vbCrLf & vbCrLf & "收據號碼:" & stER01, vbCritical, "匯入失敗"
   End If
   Set adoRst = Nothing
   
End Function

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGrd1HeadWidth
   Dim iUbound As Integer
   Dim iRow As Integer

   arrGrd1HeadWidth = Array(250, 1100, 1200, 1200, 800, 810, 1500, 810, 850)
   iUbound = UBound(arrGrd1HeadWidth)
   
   With MSHFlexGrid1
   .Visible = False
   If pReset = True Then
      .Clear
      .Rows = 2
      '.RowHeight(1) = 0
   Else
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) = "N" Then
            .row = iRow
            For iCol = 0 To .Cols - 1
               .col = iCol
               .CellBackColor = &H80000018
            Next
         End If
      Next
   End If
   .FixedCols = 2
   .FormatString = "V|本所案號|申請案號|收據號碼|金額|開立日期|檔案名稱|匯入日期時間|狀態|匯入人員"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGrd1HeadWidth(iCol)
         If iCol = 4 Then
            .ColAlignment(iCol) = flexAlignRightCenter
         Else
            .ColAlignment(iCol) = flexAlignLeftCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   
   .Visible = True
   End With
End Sub


Private Sub PrintOnePdf(ByVal program_name As String, parameters As String)

Dim process_id As Long
Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError

    process_id = Shell(program_name & parameters, vbHide)

    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub

Private Sub SetFileAssociation(Optional sFile As String)
Dim i As Integer, s2 As String
Dim bNewFile As Boolean, ff1 As Integer
Dim strReaderPath As String

'預設用 Reader,找不到才檢查關聯設定
strReaderPath = FindFirstFileAPI("C:\Program Files\Adobe\", "AcroRd32.exe")
If strReaderPath <> "" Then txtPDFPath = strReaderPath: Exit Sub

'Check if the file exists
If sFile = "" Then
   sFile = "test.pdf"
End If

If Dir(sFile) = "" Or sFile = "" Then
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   Open sFile For Output As #ff1
   Close #ff1
   bNewFile = True
End If

If Dir(sFile) = "" Or sFile = "" Then
   MsgBox "檔案不存在!", vbCritical, "PDF 檔關聯檢查"
   Exit Sub
End If

'Create a buffer
s2 = String(MAX_FILENAME_LEN, 32)
'Retrieve the name and handle of the executable, associated with this file
i = FindExecutable(sFile, vbNullString, s2)
If i > 32 Then
   txtPDFPath = Left$(s2, InStr(s2, Chr$(0)) - 1)
Else
   MsgBox "PDF 檔關聯不存在，請確認是否有安裝相關應用程式 !", , "PDF 檔關聯檢查"
End If

If bNewFile = True Then
   Kill sFile
End If
End Sub

'找檔案,回傳第一個找到的路徑
Private Function FindFirstFileAPI(path As String, SearchStr As String) As String
    Dim FileName As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim i As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    Dim bolGotIt As Boolean

    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                FindFirstFileAPI = FindFirstFileAPI & path & FileName
                bolGotIt = True
                Exit Do
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Loop
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 And bolGotIt = False Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            FindFirstFileAPI = FindFirstFileAPI & FindFirstFileAPI(path & dirNames(i) & "\", SearchStr)
            If FindFirstFileAPI <> "" Then Exit For
        Next i
    End If
End Function

Private Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
   Dim nCol As Long, nRow As Long, iRow As Integer
   Dim stValue As String
      
   If nCol < 0 Or nRow < 0 Then Exit Sub
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
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
         CheckGrid MSHFlexGrid1, stValue
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

Private Sub CheckGrid(grdDataList As MSHFlexGrid, pValue As String)
   Dim iRow As Integer
   
   With grdDataList
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) <> "X" Then
         If .TextMatrix(iRow, 0) <> pValue Then
            .row = iRow
            ClickGrid grdDataList
         End If
      End If
   Next
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

Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim ii As Integer
   With MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, ii)
         Exit For
      End If
   Next
   End With
End Function

Private Function GetFieldId(pFieldName As String) As Integer
   Dim ii As Integer
   With MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         GetFieldId = ii
         Exit For
      End If
   Next
   End With
End Function
