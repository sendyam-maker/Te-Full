VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm010033 
   BorderStyle     =   1  '單線固定
   Caption         =   "掃瞄資料匯入"
   ClientHeight    =   5750
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8960
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm010033.frx":0000
      Left            =   1530
      List            =   "frm010033.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   16
      Top             =   780
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   1275
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   13
      Text            =   "frm010033.frx":003F
      Top             =   4050
      Width           =   8865
   End
   Begin VB.Frame Frame3 
      Caption         =   "匯入錯誤訊息："
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
      Height          =   2895
      Left            =   30
      TabIndex        =   11
      Top             =   1110
      Width           =   8865
      Begin VB.ListBox List1 
         Height          =   2560
         ItemData        =   "frm010033.frx":0242
         Left            =   90
         List            =   "frm010033.frx":0244
         TabIndex        =   12
         Top             =   270
         Width           =   8685
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   315
      Left            =   8610
      TabIndex        =   1
      Top             =   450
      Width           =   345
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   345
      Left            =   6600
      TabIndex        =   3
      Top             =   60
      Width           =   885
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入(&T)"
      Height          =   345
      Left            =   5580
      TabIndex        =   2
      Top             =   60
      Width           =   885
   End
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Text            =   "C:\temp"
      Top             =   420
      Width           =   7065
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   7620
      TabIndex        =   4
      Top             =   60
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   8
      Top             =   5280
      Width           =   8895
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   8820
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid3 
      Height          =   705
      Left            =   2880
      TabIndex        =   14
      Top             =   -120
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1341
      _ExtentY        =   1252
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "檔案名稱                                                             "
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   -90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   405
      Left            =   840
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   864
      _ExtentY        =   706
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frm010033.frx":0246
   End
   Begin VB.FileListBox File1 
      Height          =   420
      Left            =   1380
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   705
      Left            =   1980
      TabIndex        =   10
      Top             =   -150
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1341
      _ExtentY        =   1252
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "檔案名稱                                                             "
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "匯入的檔案類型："
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "電子檔存放路徑："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1440
   End
End
Attribute VB_Name = "frm010033"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/19 Form2.0已修改 (無需修改)
'Create By Sindy 2014/12/17
Option Explicit

Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iLine1 As Integer
Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
Dim intUpdStarRow As Integer, intUpdEndRow As Integer
Dim strUpdCP01 As String, strUpdCP02 As String, strUpdCP03 As String, strUpdCP04 As String
Dim strUpdCP09 As String, strUpdCP10 As String
Dim strTotRow_B As String 'Add by Amy 2016/10/19 圖書封面用
Dim m_strReFileN As String 'Add By Sindy 2018/10/5 副檔名
'Dim RsQ As New ADODB.Recordset 'Add by Amy 2023/01/31

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Private Sub cmdExit_Click()
   Unload Me
End Sub

'匯入
Private Sub cmdImPort_Click()
Dim fs
Dim dblFCnt As Double
Dim dblMaxWidth As Double
Dim strTotRow As String
Dim strCaseNo As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strFileName As String
Dim strErr As String
Dim strTemp As String
Dim jj As Integer 'Add By Sindy 2018/10/4
Dim varTmp As Variant 'Add By Sindy 2018/10/5
   
On Error GoTo ErrHand
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   'Add By Sindy 2018/10/5
   If Combo1.Visible = True Then
      varTmp = Split(Combo1.Text, " ")
      If UCase(varTmp(0)) = "CASE" Then
         m_strReFileN = UCase(EMP_客戶資料)
      ElseIf UCase(varTmp(0)) = "OA" Then
         m_strReFileN = "OA"
      ElseIf UCase(varTmp(0)) = "CUS" Then
         m_strReFileN = UCase(EMP_通知函)
      End If
      If m_strReFileN = "" Then
         MsgBox "匯入的檔案類型不可空白", vbExclamation
         Exit Sub
      End If
   End If
   '2018/10/5 END
   
   If Right(Trim(txtPath1), 1) = "\" Then txtPath1 = Left(txtPath1, Len(txtPath1) - 1)
   
   '檢查資料夾
   Set fs = CreateObject("Scripting.FileSystemObject")
   File1.path = txtPath1.Text
   File1.Refresh
   If File1.ListCount = 0 Then
      MsgBox txtPath1.Text & " 此資料夾中，尚無電子檔！"
      txtPath1.SetFocus
      Exit Sub
   End If
   Set fs = Nothing
   
   cmdImport.Enabled = False 'Add By Sindy 2024/12/24
   
   dblMaxWidth = 8820
   Text2.Width = 0
   List1.Clear
   Grid2.Clear
   Grid2.Cols = 1
   Grid2.Rows = 1
   'Modify by Amy 2016/10/19 +排除圖書封面(檔名8碼)
   Grid3.Clear: Grid3.Cols = 1: Grid3.Rows = 1 '圖書封面用
'   For dblFCnt = 0 To File1.ListCount - 1
'      '檔名後4碼為.PDF者才須匯入
'      If UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
'         '檢查檔案是否正在使用中
'         If PUB_ChkFileOpening(txtPath1.Text & "\" & Trim(File1.List(dblFCnt))) = True Then
'            MsgBox Trim(File1.List(dblFCnt)) & vbCrLf & "檔案正在使用中，請關閉才可執行匯入！", vbExclamation
'            GoTo ErrHand
'         End If
'         Grid2.AddItem Trim(File1.List(dblFCnt))
'      End If
'   Next dblFCnt
   For dblFCnt = 0 To File1.ListCount - 1
      '檢查檔案是否正在使用中
      If PUB_ChkFileOpening(txtPath1.Text & "\" & Trim(File1.List(dblFCnt))) = True Then
         MsgBox Trim(File1.List(dblFCnt)) & vbCrLf & "檔案正在使用中，請關閉才可執行匯入！", vbExclamation
         'Screen.MousePointer = vbDefault
         'Exit Sub
         GoTo ErrHand
      End If
      '檔名後4碼為.PDF者才須匯入
      If UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
         If Len(Trim(File1.List(dblFCnt))) = 8 Then
            Grid3.AddItem Trim(File1.List(dblFCnt))
         Else
            Grid2.AddItem Trim(File1.List(dblFCnt))
         End If
      End If
   Next dblFCnt
   'end 2016/10/19
   Grid2.col = 0
   Grid2.row = 0
   Me.Grid2.Sort = 5 '字串昇冪
   
   strTotRow = Grid2.Rows - 1
   strTotRow_B = Grid3.Rows - 1 'Add by Amy 2016/10/19
   
   'Add By Sindy 2018/10/5
   If Val(strTotRow) + Val(strTotRow_B) = 0 Then
      MsgBox "無資料！", vbInformation
      'Exit Sub
      GoTo ErrHand
   End If
   
   Screen.MousePointer = vbHourglass
   
   '清空變數值
   intUpdStarRow = 0
   intUpdEndRow = 0
   strUpdCP01 = ""
   strUpdCP02 = ""
   strUpdCP03 = ""
   strUpdCP04 = ""
   strUpdCP09 = ""
   strUpdCP10 = ""
   For dblFCnt = 1 To strTotRow
      'Modify by Amy 2016/10/19 +strTotRow_B
      Text2.Width = dblMaxWidth / (Val(strTotRow) + Val(strTotRow_B)) * dblFCnt: DoEvents
      strErr = "": strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
      strFileName = UCase(Grid2.TextMatrix(dblFCnt, 0))
      
      '取得案號
      If InStr(strFileName, ".") > 0 Then
         strCaseNo = Trim(Left(strFileName, InStr(strFileName, ".") - 1))
      End If
      'Modify By Sindy 2018/10/4
'      If Left(strCaseNo, 1) = "P" Then
'         strCP01 = "P"
'      Else
'         strErr = convForm(CheckStr(strFileName), 30) & "系統別有誤"
'         GoTo RunSave
'      End If
      For jj = 1 To 3
         If Asc(Mid(strCaseNo, jj, 1)) >= 65 And Asc(Mid(strCaseNo, jj, 1)) <= 90 And Len(strCP01) < 3 Then  '系統別
            strCP01 = strCP01 & Mid(strCaseNo, jj, 1)
         Else
            Exit For
         End If
      Next jj
      If CheckSys(strCP01) = "" Then
         strErr = convForm(CheckStr(strFileName), 30) & "系統別有誤"
         GoTo RunSave
      End If
      '2018/10/4 END
      
      If InStr(strCaseNo, "-") = 0 Then
         strCP02 = Format(Mid(strCaseNo, Len(strCP01) + 1), "000000")
         strCP03 = "0"
         strCP04 = "00"
      Else
         'Modify By Sindy 2025/5/8 修改解析案號的程式
'         strCP02 = Format(Mid(strCaseNo, Len(strCP01) + 1, InStr(strCaseNo, "-") - 1 - Len(strCP01)), "000000")
'         strCP03 = Mid(strCaseNo, InStr(strCaseNo, "-") + 1, 1)
'         If InStr(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), "-") > 0 Then
'            strCP04 = Format(Mid(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), InStr(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), "-") + 1), "00")
'         Else
'            strCP04 = "00"
'         End If
         strExc(10) = strCaseNo
         strCP02 = Format(SystemNumber(strExc(10), 2), "000000")
         strCP03 = SystemNumber(strExc(10), 3)
         strCP04 = SystemNumber(strExc(10), 4)
         '2025/5/8 END
      End If
      '檢查strCP02的長度是否為6碼且為數字
      If Len(strCP02) <> 6 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP02長度非6碼有誤"
         GoTo RunSave
      ElseIf IsNumeric(strCP02) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP02非數字型態有誤"
         GoTo RunSave
      End If
      '檢查strCP03的長度是否為1碼且為數字
      If Len(strCP03) <> 1 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP03長度非1碼有誤"
         GoTo RunSave
      ElseIf IsNumeric(strCP03) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP03非數字型態有誤"
         GoTo RunSave
      End If
      '檢查strCP04的長度是否為2碼且為數字
      If Len(strCP04) <> 2 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP04長度非2碼有誤"
         GoTo RunSave
      ElseIf IsNumeric(strCP04) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP04非數字型態有誤"
         GoTo RunSave
      End If
      
RunSave:
      If (strUpdCP01 <> "" And _
          strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 <> strCP01 & strCP02 & strCP03 & strCP04) Or _
         strErr <> "" Then
         
         If intUpdStarRow > 0 Then
            If intUpdStarRow > 0 And intUpdEndRow = 0 Then
               intUpdEndRow = intUpdStarRow
            End If
            If strUpdCP09 = "" Then
               Call GetErrText(IIf(strErr <> "", strFileName, ""))
            Else
               Call SaveFilePDF '存檔
            End If
         End If
         '清空變數值
         intUpdStarRow = 0
         intUpdEndRow = 0
         strUpdCP09 = ""
         strUpdCP10 = ""
         If strErr <> "" Then
            List1.AddItem UCase(strErr), 0: SetListScroll List1
            strErr = ""
         Else
            '讀取下一筆資料
            Call GetUpdCP09(strCaseNo, strFileName, strCP01, strCP02, strCP03, strCP04, dblFCnt)
         End If
      Else
         Call GetUpdCP09(strCaseNo, strFileName, strCP01, strCP02, strCP03, strCP04, dblFCnt)
      End If
   Next dblFCnt
   
   If strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 <> strCP01 & strCP02 & strCP03 & strCP04 Then
      Call GetUpdCP09(strCaseNo, strFileName, strCP01, strCP02, strCP03, strCP04, dblFCnt - 1)
   End If
   If intUpdStarRow > 0 Then
      If intUpdStarRow > 0 And intUpdEndRow = 0 Then
         intUpdEndRow = intUpdStarRow
      End If
      If strUpdCP09 = "" Then
         Call GetErrText(IIf(strErr <> "", strFileName, ""))
      Else
         Call SaveFilePDF '存檔
      End If
   End If
   If strErr <> "" Then
      List1.AddItem UCase(strErr), 0: SetListScroll List1
      strErr = ""
   End If
   
   'Add by Amy 2016/10/19 +圖書封面
   If strTotRow_B > 0 Then
     Grid3.col = 0:  Grid3.row = 0
     Me.Grid3.Sort = 5 '字串昇冪
     If ImportBooks(dblMaxWidth, strTotRow) = False Then
        MsgBox "圖書封面匯入有誤,請洽電腦中心！" & vbCrLf & Err.Description
        'Screen.MousePointer = vbDefault
        'Exit Sub
        GoTo ErrHand
     End If
   End If
   'end 2016/10/19
   Text2.Width = dblMaxWidth: DoEvents
   
   Screen.MousePointer = vbDefault
   cmdImport.Enabled = True 'Add By Sindy 2024/12/24
   
   'MsgBox "匯入完畢！"
   MsgBox "匯入完畢！" & IIf(List1.ListCount > 0, vbCrLf & vbCrLf & "有匯入失敗的電子檔，請查看畫面上的【匯入錯誤訊息】", "")
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cmdImport.Enabled = True 'Add By Sindy 2024/12/24
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add by Amy 2016/10/19 圖書封面
Private Function ImportBooks(ByVal dblMaxWidth As Double, ByVal strTotRow As String) As Boolean
    Dim i As Integer
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim FS_B, F_B
    Dim strFilePath As String, strFileName As String, strErr As String, strIBF02 As String
    
On Error GoTo ErrHand

    ImportBooks = False
    
    For i = 1 To strTotRow_B
        Text2.Width = dblMaxWidth / (Val(strTotRow) + Val(strTotRow_B)) * (Val(strTotRow) + i): DoEvents
        strErr = ""
        strFileName = UCase(Grid3.TextMatrix(i, 0))
        strFilePath = txtPath1.Text & "\" & strFileName
        strIBF02 = Replace(strFileName, ".PDF", "")
        Set FS_B = CreateObject("Scripting.FileSystemObject")
        Set F_B = FS_B.GetFile(strFilePath)
        '檔案大小為 0 KB 有誤
        If F_B.Size = 0 Then
            List1.AddItem strFileName & " 檔案大小為 0 KB,請確認檔案", 0: SetListScroll List1
        ElseIf ExistCheck("BooksData", "BK01", strIBF02, strErr, False) = False Then
            List1.AddItem strFileName & " 找不到可對應的圖書編號", 0: SetListScroll List1
        Else
            strQ = "Select * From ImgByteFile Where IBF01='BOK' And IBF02='00" & strIBF02 & "' " & _
                "And IBF03='0' And IBF04='00' And IBF05='6' "
            If RsQ.State <> adStateClosed Then RsQ.Close
            RsQ.CursorLocation = adUseClient
            RsQ.Open strQ, cnnConnection, adOpenStatic, adLockOptimistic
            If RsQ.RecordCount > 0 Then
                List1.AddItem strFileName & " 檔案上傳過,請通知電腦中心刪除後再上傳", 0: SetListScroll List1
            Else
                If SaveImgByteFile_BOK(strIBF02, strFilePath, strErr, RsQ) = False Then
                      List1.AddItem strFileName & " " & strErr, 0: SetListScroll List1
                End If
            End If
        End If
   Next i
   
   ImportBooks = True
   Exit Function
   
ErrHand:
    
End Function

Private Function SaveImgByteFile_BOK(ByVal stIBF02 As String, ByVal FileFullPath As String, ByRef strErr As String, ByRef PdfRs As ADODB.Recordset) As Boolean
    Dim bytes() As Byte
    Dim strSql As String
    Dim file_num As Integer
    Dim strFtpPath As String
    
On Error GoTo CheckingErr
    
    SaveImgByteFile_BOK = False
    
    file_num = FreeFile
    Open FileFullPath For Binary Access Read As #file_num
'    ReDim bytes(LOF(file_num))
'    Get #file_num, , bytes()
      PdfRs.AddNew
      PdfRs.Fields("ibf01").Value = "BOK"
      PdfRs.Fields("ibf02").Value = "00" & stIBF02
      PdfRs.Fields("ibf03").Value = "0"
      PdfRs.Fields("ibf04").Value = "00"
         
      PdfRs.Fields("ibf05").Value = "6" '種類
      PdfRs.Fields("ibf06").Value = "5" '格式
      PdfRs.Fields("ibf08").Value = Val(strSrvDate(1))
      PdfRs.Fields("ibf09").Value = Val(Format(time, "HHMM"))
      PdfRs.Fields("ibf13").Value = Trim(LOF(file_num))
'      PdfRs.Fields("ibf14").Value = Null
'      PdfRs.Fields("ibf14").AppendChunk bytes()
      Close #file_num
      'Modify By Sindy 2017/8/10
      '檔案改放FTP
      PUB_PutFtpFile FileFullPath, "BOK-" & "00" & stIBF02 & "-0-00-6", "BOK-" & "00" & stIBF02 & "-0-00-6", strFtpPath, UCase("imgbytefile")
      If strFtpPath <> "" Then
         PdfRs.Fields("ibf15") = strFtpPath
      End If
      '2017/8/10 END
      PdfRs.UPDATE
      
      If Dir(FileFullPath) <> "" Then Kill FileFullPath
      SaveImgByteFile_BOK = True
      Exit Function
      
CheckingErr:
   PdfRs.CancelUpdate
   strErr = Err.Description
End Function
'end 2016/10/19

Private Sub GetUpdCP09(strCaseNo As String, strFileName As String, _
                       strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
                       dblFCnt As Double)
Dim strConSql As String
Dim strTemp As String
Dim strChkDate As String
   
   If intUpdStarRow = 0 Then
      strUpdCP01 = strCP01
      strUpdCP02 = strCP02
      strUpdCP03 = strCP03
      strUpdCP04 = strCP04
      
      intUpdStarRow = dblFCnt
   Else
      intUpdEndRow = dblFCnt
   End If
   
   '抓取此本所案號要歸檔的文號:發文日為系統日3個月以內,A類
   If strUpdCP09 = "" Then
      strChkDate = Format(DateAdd("m", -3, DateSerial(Left(strSrvDate(1), 4), Mid(strSrvDate(1), 5, 2), Right(strSrvDate(1), 2))), "YYYYMMDD")
      '專利大陸案,且收文日>=20150101
      'Modified by Morgan 2014/12/30 改判斷發文日>=20150101
      'Modified by Sindy 2015/1/26 改判斷20140601,且AB類都抓 (and cp09<'C')
      'Modified by Sindy 2016/12/29 只抓A類 (and substr(cp09,1,1)='A')
      'Modify By Sindy 2018/10/4 開放各系統別
'      strSql = "select cp09,cp10" & _
'               " From caseprogress,patent" & _
'               " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
'               " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
'               " and cp27 is not null and cp27>=20140601 and pa09='000'" & _
'               " and cp27>=" & strChkDate & _
'               " and substr(cp09,1,1)='A'" & _
'               " order by cp27 desc,cp09 asc"
      'Modify By Sindy 2018/10/5
      If UCase(m_strReFileN) = "CASE" Then '客戶資料
         '僅匯入該案號最近3個月內已發文的A類程序
         strSql = "select cp09,cp10" & _
                  " From caseprogress" & _
                  " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
                  " and cp27 is not null and cp27>=" & strChkDate & _
                  " and substr(cp09,1,1)='A'" & _
                  " order by cp27 desc,cp66 desc,cp67 desc"
      ElseIf UCase(m_strReFileN) = "OA" Then '官方來函
         '僅匯入該案號最近的C類程序
         strSql = "select cp09,cp10" & _
                  " From caseprogress" & _
                  " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
                  " and substr(cp09,1,1)='C'" & _
                  " order by cp66 desc,cp67 desc"
      ElseIf UCase(m_strReFileN) = "CUS" Then '通知函
         '僅匯入該案號最近已發文的A、B、C類程序
         strSql = "select cp09,cp10" & _
                  " From caseprogress" & _
                  " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
                  " and cp27 is not null and cp27>0" & _
                  " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or substr(cp09,1,1)='C')" & _
                  " order by cp27 desc,cp66 desc,cp67 desc"
      'Add By Sindy 2018/10/12 桂英提出
      '最近一筆程序
      Else
         'Modify By Sindy 2019/8/29 + ,cp09 desc : FCT-42464,OA_SCAN(但歸到B類外商發文,應該歸C類敗訴)
         strSql = "select cp09,cp10" & _
                  " From caseprogress" & _
                  " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
                  " order by cp66 desc,cp67 desc,cp09 desc"
      End If
      '2018/10/5 END
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         strUpdCP09 = RsTemp.Fields("cp09")
         strUpdCP10 = RsTemp.Fields("cp10")
      End If
   End If
End Sub

Private Sub GetErrText(strFName As String)
Dim i As Integer
Dim strText As String
   
   '失敗時,則整卷不存
   If intUpdStarRow > 0 Then
      For i = intUpdStarRow To intUpdEndRow
         If UCase(Trim(strFName)) <> UCase(Trim(Grid2.TextMatrix(i, 0))) Then
            strText = convForm(CheckStr(Grid2.TextMatrix(i, 0)), 30) & IIf(strUpdCP09 = "", "找不到歸卷的文號，", "")
'            If intUpdStarRow <> intUpdEndRow Then
'               strText = strText & "整卷不存"
'            Else
               strText = Left(strText, Len(strText) - 1)
'            End If
            List1.AddItem UCase(strText), 0: SetListScroll List1
         End If
      Next
   End If
End Sub

'存檔
Private Function SaveFilePDF() As Boolean
Dim dblFCnt As Double
Dim strFileName As String
Dim strFullFileName As String
Dim stReName As String
Dim stReName2 As String 'Add By Sindy 2018/11/13
Dim fs, f
Dim strErr As String
Dim bolSave As Boolean
Dim bolCnn As Boolean
Dim strTcp01 As String, strTcp02 As String, strTcp03 As String, strTcp04 As String
Dim bolGetFileName As Boolean
Dim intRow As Integer
   
On Error GoTo ErrHand
   
'   bolSave = True
'   cnnConnection.BeginTrans
'   bolCnn = True
   
   For dblFCnt = intUpdStarRow To intUpdEndRow
      strErr = "" 'Add By Sindy 2018/12/4
      strFileName = Grid2.TextMatrix(dblFCnt, 0)
      strFullFileName = txtPath1.Text & "\" & strFileName
      bolSave = True
      cnnConnection.BeginTrans
      bolCnn = True
      
'      '檢查檔名規則
'      If PUB_ChkEmpFlowFNMRule(strUpdCP01 & "-" & strUpdCP02 & "-" & strUpdCP03 & "-" & strUpdCP04, strFileName, "Y", strUpdCP10, , , False, False, strErr) = False Then
'         bolSave = False
'         'Exit For
'         GoTo ReadNext
'      End If
'      '更名
'      If PUB_GetEmpFlowReNameFile(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04, strUpdCP10, strFileName, stReName, True, 1, False, strErr) = False Then
'         bolSave = False
'         'Exit For
'         GoTo ReadNext
'      End If
      
      '取得檔名
      bolGetFileName = False: intRow = 0
      Do While bolGetFileName = False
         'stReName = Trim(strUpdCP01) & CStr(Val(strUpdCP02)) & IIf(strUpdCP03 <> "0" Or strUpdCP04 <> "00", "-" & strUpdCP03, "") & IIf(strUpdCP04 <> "00", "-" & strUpdCP04, "")
         stReName = Trim(strUpdCP01) & CStr(strUpdCP02) & IIf(strUpdCP03 <> "0" Or strUpdCP04 <> "00", "-" & strUpdCP03, "") & IIf(strUpdCP04 <> "00", "-" & strUpdCP04, "")
         'Add By Sindy 2018/11/13
         'stReName2 = Trim(strUpdCP01) & CStr(strUpdCP02) & IIf(strUpdCP03 <> "0" Or strUpdCP04 <> "00", "-" & strUpdCP03, "") & IIf(strUpdCP04 <> "00", "-" & strUpdCP04, "")
         stReName2 = Trim(strUpdCP01) & CStr(Val(strUpdCP02)) & IIf(strUpdCP03 <> "0" Or strUpdCP04 <> "00", "-" & strUpdCP03, "") & IIf(strUpdCP04 <> "00", "-" & strUpdCP04, "")
         'Added by Lydia 2019/05/03 檔名不可為.TS.PDF,因為查名結果主要是放在查名單內
         If (strUpdCP01 = "T" Or strUpdCP01 = "TS") And Right(UCase(strFileName), Len(".TS.PDF")) = ".TS.PDF" Then
             strErr = convForm(CheckStr(strFileName), 30) & "檔名不可為.TS.PDF，請更名為.CASE.PDF"
             bolSave = False
             GoTo ReadNext
         End If
         
         'Add By Sindy 2018/10/12 桂英說她們組要自己輸副檔名
         If m_strReFileN = "" Then
            strSql = "select EFC01,EFC02 from efilecaption where EFC06='Y'" & _
                     " and instr(upper('" & ChgSQL(strFileName) & "'),upper('.'||EFC02||'.'))>0" & _
                     " and efc01 in('ALL','" & strUpdCP01 & "')" & _
                     " order by efc01 desc,efc02 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If intRow = 0 Then '取得副檔名後原始檔名
                  stReName = stReName & "." & strUpdCP10 & Mid(strFileName, InStr(UCase(strFileName), "." & UCase(RsTemp.Fields("EFC02")) & "."))
               Else '有重覆再加序號
                  stReName = stReName & "." & strUpdCP10 & "." & RsTemp.Fields("EFC02") & "." & intRow & Mid(strFileName, InStr(UCase(strFileName), "." & UCase(RsTemp.Fields("EFC02")) & ".") + Len("." & UCase(RsTemp.Fields("EFC02"))))
               End If
            Else
               'Modify By Sindy 2018/11/13 + or UCase(strFileName) = UCase(stReName2 & ".pdf")
               If UCase(strFileName) = UCase(stReName & ".pdf") Or _
                  UCase(strFileName) = UCase(stReName2 & ".pdf") Then '官方來函
                  stReName = stReName & "." & strUpdCP10 & ".pdf"
                  'Add By Sindy 2018/11/2
                  If intRow = 1 Then
                     'Modify By Sindy 2023/9/27 + 顯示時間
                     strErr = convForm(CheckStr(strFileName), 30) & "檔案已存在, 請檢查卷宗區(文號:" & strUpdCP09 & ")! " & Format(Right("000000" & ServerTime, 6), "##:##:##")
                     bolSave = False
                     GoTo ReadNext
                  End If
                  '2018/11/2 END
               Else
                  strErr = convForm(CheckStr(strFileName), 30) & "沒有輸入檔案屬性的副檔名(文號:" & strUpdCP09 & ")!"
                  bolSave = False
                  GoTo ReadNext
               End If
            End If
         Else
         '2018/10/12 END
            'Modify By Sindy 2018/10/4
            'stReName = stReName & "." & strUpdCP10 & "." & EMP_客戶資料 & IIf(intRow > 0, intRow, "") & Mid(strFileName, InStrRev(strFileName, "."))
            'Modify By Sindy 2018/10/5 增加可以匯入其他檔案類型
            'stReName = stReName & "." & strUpdCP10 & "." & EMP_客戶資料 & IIf(intRow > 0, "." & intRow, "") & Mid(strFileName, InStrRev(strFileName, "."))
            stReName = stReName & "." & strUpdCP10 & "." & m_strReFileN & IIf(intRow > 0, "." & intRow, "") & Mid(strFileName, InStrRev(strFileName, "."))
         End If
         
         '檢查檔案是否已存在
         stReName2 = stReName2 & Mid(stReName, InStr(stReName, "."))
         strSql = "select cpp02" & _
                  " From casepaperpdf" & _
                  " where cpp01='" & strUpdCP09 & "'" & _
                  " and substr(upper(cpp02),-4)<>'.DEL'" & _
                  " and (upper(cpp02)='" & UCase(stReName) & "' or upper(cpp02)='" & UCase(stReName2) & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
            bolGetFileName = True
            'Modify By Sindy 2018/10/4 Mark
'            Call PUB_InsEfileCaption(strUpdCP01, EMP_客戶資料, intRow) '檢查是否有需要新增電子檔次要副檔名說明
'            Exit Do 'Modify By Sindy 2024/12/2 Mark:因還要再檢查歷程附件區
         Else
            bolGetFileName = False
            GoTo ReadNext_GetFN '檔名重覆,重取序號
         End If
         
         'Add By Sindy 2024/12/2 檢查歷程附件區是否有重覆,以防待送件區歸檔重覆了
         strSql = "select EMPELECTRONFILE.*" & _
                  " From EMPELECTRONFILE,caseprogress" & _
                  " where eef01='" & strUpdCP09 & "' and cp09=eef01 and cp158=0 and cp159=0" & _
                  " and (upper(eef03)='" & Replace(UCase(stReName), "." & strUpdCP10 & ".", ".") & "' or upper(eef03)='" & Replace(UCase(stReName2), "." & strUpdCP10 & ".", ".") & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
            bolGetFileName = True
            Exit Do
         Else
            bolGetFileName = False
            GoTo ReadNext_GetFN '檔名重覆,重取序號
         End If
         '2024/12/2 END
         
ReadNext_GetFN: 'Add By Sindy 2024/12/2
         intRow = intRow + 1
      Loop
      
      '檢查此文號的案號是否與系統抓到的案號一致
      strSql = "select cp01,cp02,cp03,cp04" & _
               " From caseprogress" & _
               " where cp09='" & strUpdCP09 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strTcp01 = RsTemp.Fields("cp01")
         strTcp02 = RsTemp.Fields("cp02")
         strTcp03 = RsTemp.Fields("cp03")
         strTcp04 = RsTemp.Fields("cp04")
         If strTcp01 <> strUpdCP01 Or _
            strTcp02 <> strUpdCP02 Or _
            strTcp03 <> strUpdCP03 Or _
            strTcp04 <> strUpdCP04 Then
            strErr = convForm(CheckStr(strFileName), 30) & "文號" & strUpdCP09 & _
                     "本所案號" & strTcp01 & strTcp02 & strTcp03 & strTcp04 & _
                     "與系統抓到的案號" & strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 & _
                     "不一致!"
            bolSave = False
            'Exit For
            GoTo ReadNext
         End If
      End If
      
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set f = fs.GetFile(strFullFileName)
      '檔案大小為 0 KB 有誤
      If f.Size = 0 Then
         strErr = convForm(CheckStr(strFileName), 30) & MsgText(9221)
         bolSave = False
         'Exit For
         GoTo ReadNext
      End If
      
      If SaveAttFile_PDF(strUpdCP09, strFullFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "存檔失敗！" & vbCrLf & Err.Description
         bolSave = False
         PUB_SendMail strUserNum, "97038", "", "掃瞄資料匯入:" & strFullFileName & " 文號:" & strUpdCP09 & "(" & stReName & ")"
         'Exit For
         GoTo ReadNext
'      Else
'         PUB_SendMail strUserNum, "97038", "", "掃瞄資料匯入" & stReName, stReName
      End If
      
ReadNext:
      If bolSave = False Then
         cnnConnection.RollbackTrans
         bolCnn = False
         strErr = Replace(strErr, vbCrLf, "")
         List1.AddItem UCase(strErr), 0: SetListScroll List1
      Else
         cnnConnection.CommitTrans
         bolCnn = False
         fs.DeleteFile strFullFileName, True '刪檔
      End If
   Next dblFCnt
   
   Exit Function
   
ErrHand:
   If bolCnn = True Then
      cnnConnection.RollbackTrans
   End If
   MsgBox Err.Description
End Function

'列印
Private Sub cmdPrint_Click()
Dim i As Integer, j As Integer
   
   iLine1 = 0
   For j = List1.ListCount - 1 To 0 Step -1
      For i = 1 To 1
         strTemp(i) = ""
      Next i
      strTemp(1) = List1.List(j)
      If iLine1 > 52 Or iLine1 = 0 Then
         If iLine1 > 0 Then Printer.NewPage: iLine1 = 0
         PrintTitle '列印表頭
      End If
      PrintDetail '列印明細
   Next j
   Printer.EndDoc
End Sub

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.pdf"
      .Filter = "PDF檔案 (*.pdf)|*.pdf"
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir" & EMP_客戶資料, "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir" & EMP_客戶資料, "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir" & EMP_客戶資料, sFile(0)
            txtPath1.Text = sFile(0)
         Else
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
'               For ii = Len(.FileName) To 1 Step -1
'                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir" & EMP_客戶資料, Left(.FileName, InStrRev(.FileName, "\") - 1)
'                     Exit For
'                  End If
'               Next ii
            End If
            'txtPath1.Text = .FileName
            txtPath1.Text = Left(.FileName, InStrRev(.FileName, "\") - 1)
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
Dim i As Integer, j As Integer
   
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
   'cmbPrinter2.Text = cmbPrinter2.List(SeekPrint)
   
   If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir" & EMP_客戶資料, "") <> "" Then
      txtPath1.Text = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir" & EMP_客戶資料, "")
   'Else
   '   txtPath1.Text = PUB_Getdesktop
   End If
   
   'Add By Sindy 2018/10/5 +匯入的檔案類型 (管理部)
   If Left(PUB_GetST03(strUserNum), 1) = "M" And PUB_GetST03(strUserNum) <> "M51" Then
      Combo1.ListIndex = 0 'CASE 客戶資料
      Combo1.Locked = True
      Combo1.Enabled = False
   ElseIf PUB_GetST03(strUserNum) = "M51" Then
      Combo1.ListIndex = 1 'OA   官方來函
   Else '其他單位
      Combo1.Visible = False
      Label1.Visible = False
   End If
   'Add by Amy 2023/01/31 避免同部門同時操作,造成檔案匯入有問題
   'Modify by Amy 2023/02/13 改共用
   'If ChkLock("A") = True Then
   If Pub_ChkLock(1, Me.Name, "A") = True Then
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2023/01/31  避免同部門同時操作,造成檔案匯入有問題
   'Modify by Amy 2023/02/13 改共用
   'Call ChkLock("D")
   'Set RsQ = Nothing
   Call Pub_ChkLock(3, Me.Name, "D")
   'end 2023/01/31
   Set frm010033 = Nothing
End Sub

Private Sub txtPath1_GotFocus()
   InverseTextBox txtPath1
End Sub

Private Function TxtValidate() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim Cancel As Boolean

TxtValidate = False

If IsEmptyText(txtPath1) = True Then
   strTit = "檢核資料"
   strMsg = "請輸入電子檔存放路徑！"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   txtPath1.SetFocus
   Exit Function
End If

TxtValidate = True
End Function

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 5500
End Sub

Sub PrintTitle()
GetPleft
iLine1 = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("匯入電子檔錯誤訊息清單") / 2)
Printer.CurrentY = iLine1 * 300
Printer.Print "匯入電子檔錯誤訊息清單"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine1 = iLine1 + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iLine1 = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print "錯誤訊息"

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print String(148, "-")
iLine1 = iLine1 + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
   For m_j = 1 To 1
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine1 * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine1 = iLine1 + 1
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

'Add by Amy 2023/01/31 取得/釋放資料 p_Status:狀態
'Mark by Amy 2023/02/13 改共用
'Private Function ChkLock(ByVal p_Status As String) As Boolean
'    Dim stQ As String, intQ As Integer
'On Error GoTo ErrHand
'
'    If p_Status = "D" Then
'        stQ = "Delete From LockRec Where LR02='" & strUserNum & "' And LR01='" & Me.Name & "-" & strUserNum & "' "
'        cnnConnection.Execute stQ
'    Else
'        stQ = "Delete From LockRec Where LR02='" & strUserNum & "' And LR01='" & Me.Name & "-" & strUserNum & "' "
'        cnnConnection.Execute stQ
'        stQ = "Select st02 From LockRec,Staff Where LR01 Like '" & Me.Name & "%' And LR02=st01(+) And st03='" & Pub_StrUserSt03 & "' "
'        intQ = 1
'        Set RsQ = ClsLawReadRstMsg(intQ, stQ)
'        If intQ = 1 Then
'          stQ = "" & RsQ.GetString(adClipString, , , ",")
'          MsgBox "【" & Mid(stQ, 1, Len(stQ) - 1) & "】" & vbCrLf & _
'                        "正使用" & Me.Caption & "作業！", vbInformation
'        End If
'        stQ = "Insert Into LockRec(LR01,LR02,LR03) Values ('" & Me.Name & "-" & strUserNum & "','" & strUserNum & "',to_char(sysdate,'YYYYMMDDHH24MISS'))"
'        cnnConnection.Execute stQ
'    End If
'    ChkLock = True
'    Exit Function
'
'ErrHand:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description, vbCritical, "取得/釋放資料異動"
'    End If
'End Function
