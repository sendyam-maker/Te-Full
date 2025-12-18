VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020320 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸商標審定公告及通知續展匯入作業"
   ClientHeight    =   6825
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8955
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "frm020320.frx":0000
      Top             =   4920
      Width           =   8775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00FF0000&
      Height          =   945
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "frm020320.frx":0021
      Top             =   5730
      Width           =   8775
   End
   Begin VB.FileListBox File2 
      Height          =   450
      Left            =   6090
      TabIndex        =   24
      Top             =   540
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "資料範圍"
      Height          =   1005
      Left            =   780
      TabIndex        =   23
      Top             =   1200
      Width           =   6795
      Begin VB.OptionButton Option2 
         Caption         =   "台灣申請人案件 (第二次) .TXT   (無使用)"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   630
         Width           =   4005
      End
      Begin VB.OptionButton Option2 
         Caption         =   "本所大陸代理人全部案件(第一次) .XLS"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   3405
      End
      Begin VB.Label Label9 
         Caption         =   "檔案命名：期數-xx（ 例：1328-1.xls）"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3690
         TabIndex        =   25
         Top             =   300
         Width           =   3045
      End
   End
   Begin VB.TextBox text1 
      Height          =   345
      Index           =   2
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   0
      Top             =   420
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   5385
      Left            =   4680
      TabIndex        =   14
      Top             =   6840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9499
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5385
      Left            =   60
      TabIndex        =   13
      Top             =   6840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9499
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   345
      Left            =   7590
      TabIndex        =   10
      Top             =   3990
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<="
      Height          =   345
      Left            =   7590
      TabIndex        =   5
      Top             =   2340
      Width           =   345
   End
   Begin VB.OptionButton Option1 
      Caption         =   "通知續展 .TXT   (無使用)"
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   8
      Top             =   3660
      Width           =   3405
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Index           =   1
      Left            =   1620
      TabIndex        =   9
      Top             =   3990
      Width           =   5925
   End
   Begin VB.OptionButton Option1 
      Caption         =   "審定公告"
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   1
      Top             =   840
      Width           =   1245
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Index           =   0
      Left            =   1620
      TabIndex        =   4
      Top             =   2340
      Width           =   5925
   End
   Begin VB.TextBox text1 
      Height          =   345
      Index           =   0
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2730
      Width           =   1095
   End
   Begin VB.TextBox text1 
      Height          =   345
      Index           =   1
      Left            =   1620
      MaxLength       =   4
      TabIndex        =   7
      Top             =   3150
      Width           =   645
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7800
      TabIndex        =   12
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6975
      TabIndex        =   11
      Top             =   60
      Width           =   756
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6270
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label11 
      Caption         =   "摩知輪商標數据，欄位為下："
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   90
      TabIndex        =   29
      Top             =   4650
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "檔案內容及欄位的順序，為下："
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   90
      TabIndex        =   27
      Top             =   5460
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   480
      X2              =   8030
      Y1              =   3570
      Y2              =   3570
   End
   Begin MSForms.Label Label8 
      Height          =   285
      Left            =   1620
      TabIndex        =   22
      Top             =   90
      Width           =   1365
      VariousPropertyBits=   27
      Size            =   "2408;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "服務人員："
      Height          =   180
      Left            =   690
      TabIndex        =   21
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "分機："
      Height          =   180
      Left            =   690
      TabIndex        =   20
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "注意：當程式正在執行時，請暫時不要使用Word！"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   4440
      TabIndex        =   19
      Top             =   4440
      Width           =   4005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "期　數："
      Height          =   180
      Left            =   690
      TabIndex        =   18
      Top             =   3210
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "檔案路徑："
      Height          =   180
      Left            =   690
      TabIndex        =   17
      Top             =   4050
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "檔案路徑："
      Height          =   180
      Left            =   690
      TabIndex        =   16
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "公告日：                             (民國年月日)"
      Height          =   180
      Left            =   690
      TabIndex        =   15
      Top             =   2790
      Width           =   3045
   End
End
Attribute VB_Name = "frm020320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 Label8
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim PLeft(1 To 6) As Integer
Dim strTemp(1 To 6) As String
Dim p_Recs1 As Integer, p_Recs2 As Integer
Dim p_Recs1_Y As Integer, p_Recs1_N As Integer
Dim p_Recs1_1102 As Integer
Dim p_Recs2_Y As Integer, p_Recs2_N As Integer
Dim p_Recs2_Recv As Integer, p_Recs2_Close As Integer, p_Recs2_Print As Integer
Dim m_AppAddrZip As String '申請人地址郵遞區號
Dim m_AppAddr As String '申請人地址
Dim m_AppName As String '商標註冊人
Dim m_AppDate As String '申請日期
Dim m_AppNum As String '註冊號
Dim m_TName As String '商標名稱
Dim m_Goods As String '分類
Dim m_DATE As String '公告日期
Dim m_DATEs As String, m_DATEe As String '專用期間
Dim m_Num As String '期數
Dim bolRetry As Boolean '是否已發生錯誤且重試
Dim iLine As Integer
Dim strCP09 As String, strCP10 As String
Dim m_CaseNo As String 'Add By Sindy 2012/10/15
Dim tmp_TName As String, tmp_AppAddrZip As String, tmp_AppAddr As String, tmp_AppName As String 'Add By Sindy 2013/11/25


Private Sub ResetGrid(ByRef p_Grid As MSHFlexGrid, Index As Integer)
   With p_Grid
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      If Index = 0 Then
         .FormatString = "選擇|註冊號|商標名稱|商標註冊人|申請日期|專用期(起)|專用期(止)|國際分類|公告狀態|申請人地址|初審公告|商標類型|代理機構|是否為本所案件|TM01|TM02|TM03|TM04|CP09|CP10|TM10|TM11|TM23|TM20|TM77|ID"
      ElseIf Index = 1 Then
         .FormatString = "序號|註冊號|商標名稱|商標註冊人|申請日期|專用期(起)|專用期(止)|國際分類|公告狀態|申請人地址|是否為本所案件|TM01|TM02|TM03|TM04|TM11|TM22|TM29|NP01|NP06|NP07|已收文|ID"
      End If
   End With
End Sub

Private Sub ChkRecs2_Recv()
Dim strChkDate As String
   '無下一程序時, 檢查是否有中途接案, 若有, 視為已收文
   strChkDate = Format(Val(strSrvDate(1)) - 30000)
   strExc(0) = "SELECT * FROM CaseProgress " & _
                     "WHERE CP01='" & MSHFlexGrid2.TextMatrix(p_Recs2, 11) & "' " & _
                     "and CP02='" & MSHFlexGrid2.TextMatrix(p_Recs2, 12) & "' " & _
                     "and CP03='" & MSHFlexGrid2.TextMatrix(p_Recs2, 13) & "' " & _
                     "and CP04='" & MSHFlexGrid2.TextMatrix(p_Recs2, 14) & "' " & _
                     "and CP10='102' " & _
                     "and CP05>=" & Val(strChkDate)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      p_Recs2_Recv = p_Recs2_Recv + 1
      MSHFlexGrid2.TextMatrix(p_Recs2, 21) = "Y"
   End If
End Sub

Private Function LoadXLS() As Boolean
Dim iRow As Integer
Dim strChkData As String, strWord As String, strTemp As String
Dim i As Integer, j As Integer, intRow As Integer
Dim fs, f
Dim strText As String
Dim intTab As Integer
Dim strErrNot101 As String 'Add By Sindy 2010/10/20
Dim strNameC, strNameE As String 'Add By Sindy 2010/11/12
'Add By Sindy 2012/9/28
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim bolReadExit As Boolean
Dim dblFCnt As Integer
'2012/9/28 End
Dim strErrText As String 'Add By Sindy 2021/12/23

On Error GoTo ErrHnd
   
   p_Recs1 = 0: p_Recs2 = 0
   p_Recs1_Y = 0: p_Recs1_N = 0
   p_Recs1_1102 = 0
   p_Recs2_Y = 0: p_Recs2_N = 0
   p_Recs2_Recv = 0: p_Recs2_Close = 0: p_Recs2_Print = 0
   strErrNot101 = ""
   m_CaseNo = "" 'Add By Sindy 2012/10/15
   intRow = 0
   
'審定公告：
'　選擇，類別，註冊號，商標名稱(中)，商標名稱(英)，商標註冊人，申請日期，專用期(起)，專用期(止)，
'　最後公告，最後流程，申請人地址，初審公告，商標類型，代理機構
'Modify By Sindy 2018/10/2 國方改版-欄位調整
'  0     1       2     3             4             5           6         7           8
'　選擇，註冊號，分類，商標名稱(中)，商標名稱(英)，商標註冊人，申請日期，專用期(起)，專用期(止)，
'  9         10        11        12
'　商標狀態，最後公告，最後流程，申請人地址
   If Option1(0).Value = True Then
      If Dir(txtPath(0)) = "" Then
'         If MsgBox("審定公告檔案不存在，是否繼續！", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
'            txtPath(0).SetFocus
'            Exit Function
'         End If
         MsgBox "審定公告檔案不存在！", vbExclamation
         txtPath(0).SetFocus
         Exit Function
      Else
         'Add By Sindy 2012/9/24
         '全部資料(第一次)
         If Option2(0).Value = True Then
            For i = Len(txtPath(0)) To 1 Step -1
               If Mid(txtPath(0), i, 1) = "\" Then
                  Exit For
               End If
            Next i
            File2.path = Mid(txtPath(0), 1, i)
            File2.Refresh
            
            Screen.MousePointer = vbHourglass
            ResetGrid MSHFlexGrid1, 0
            
            For dblFCnt = 0 To File2.ListCount - 1
               '檔名:期數-
               If InStr(File2.List(dblFCnt), Trim(Text1(1)) & "-") = 0 Then GoTo ReadNextFile
               xlsSalesPoint.Workbooks.Open File2.path & "\" & File2.List(dblFCnt)
               Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
               bolReadExit = False: iRow = 2
               '先檢查標題:
               '商標名稱.D & E=>G
               If (wksaccrpt114.Range("G1").Value = "商?名?" Or wksaccrpt114.Range("G1").Value = "商標名稱") And _
                  Len(wksaccrpt114.Range("G1").Value) = 4 Then
                  MsgBox "格式有問題，G欄位必須是商標名稱！", vbExclamation
                  bolReadExit = True
               End If
               '註冊號.B=>J
               If (wksaccrpt114.Range("J1").Value = "申??" Or wksaccrpt114.Range("J1").Value = "申請號") And _
                  Len(wksaccrpt114.Range("J1").Value) = 3 Then
                  MsgBox "格式有問題，J欄位必須是申請號！", vbExclamation
                  bolReadExit = True
               End If
               '申請日期.G=>P
               If (wksaccrpt114.Range("P1").Value = "申?日期" Or wksaccrpt114.Range("P1").Value = "申請日期") And _
                  Len(wksaccrpt114.Range("P1").Value) = 4 Then
                  MsgBox "格式有問題，P欄位必須是申請日期！", vbExclamation
                  bolReadExit = True
               End If
               Do While bolReadExit = False
                  strErrText = "" 'Add By Sindy 2021/12/23
                  '註冊號空白或商標名稱空白,則代表結束
                  If wksaccrpt114.Range("J" & iRow).Value = "" Or _
                     wksaccrpt114.Range("G" & iRow).Value = "" Then
                     bolReadExit = True
                  Else
                     p_Recs1 = p_Recs1 + 1
                     MSHFlexGrid1.Rows = p_Recs1 + 1
                     'Modify By Sindy 2014/6/25
                     strErrText = File2.path & "\" & File2.List(dblFCnt) & " 內容有誤:註冊號,第 " & iRow & " 筆" 'Add By Sindy 2021/12/23
                     MSHFlexGrid1.TextMatrix(p_Recs1, 1) = Trim(wksaccrpt114.Range("J" & iRow).Value) '註冊號
                     strErrText = File2.path & "\" & File2.List(dblFCnt) & " 內容有誤:申請日期,第 " & iRow & " 筆" 'Add By Sindy 2021/12/23
                     MSHFlexGrid1.TextMatrix(p_Recs1, 4) = ChangeWDateStringToWString(Replace(Trim(wksaccrpt114.Range("P" & iRow).Value), "-", "/")) '申請日期
                     '是否為本所案件
                     strExc(0) = "SELECT * FROM TradeMark WHERE TM12='" & MSHFlexGrid1.TextMatrix(p_Recs1, 1) & "' and TM10='020' and TM28='1'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI <> 1 And Len(Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1))) < 8 Then
                        '補足8碼檢核
                        strTemp = Right("00000000" & Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1)), 8)
                        strExc(0) = "SELECT * FROM TradeMark WHERE TM12='" & strTemp & "' and TM10='020' and TM28='1'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           MSHFlexGrid1.TextMatrix(p_Recs1, 1) = strTemp
                        End If
                     End If
                     If intI = 1 Then
                        p_Recs1_Y = p_Recs1_Y + 1
                        m_CaseNo = m_CaseNo & "　　　　　　　" & RsTemp.Fields("TM01") & "-" & RsTemp.Fields("TM02") & "-" & RsTemp.Fields("TM03") & "-" & RsTemp.Fields("TM04") & "（" & Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1)) & "）" & vbCrLf 'Add By Sindy 2013/1/2
                        MSHFlexGrid1.TextMatrix(p_Recs1, 13) = "Y"
                        MSHFlexGrid1.TextMatrix(p_Recs1, 14) = "" & RsTemp.Fields("TM01")
                        MSHFlexGrid1.TextMatrix(p_Recs1, 15) = "" & RsTemp.Fields("TM02")
                        MSHFlexGrid1.TextMatrix(p_Recs1, 16) = "" & RsTemp.Fields("TM03")
                        MSHFlexGrid1.TextMatrix(p_Recs1, 17) = "" & RsTemp.Fields("TM04")
                        MSHFlexGrid1.TextMatrix(p_Recs1, 20) = "" & RsTemp.Fields("TM10") '申請國家
                        MSHFlexGrid1.TextMatrix(p_Recs1, 21) = "" & RsTemp.Fields("TM11") '申請日
                        MSHFlexGrid1.TextMatrix(p_Recs1, 22) = "" & RsTemp.Fields("TM23") '申請人1
                        MSHFlexGrid1.TextMatrix(p_Recs1, 23) = "" & RsTemp.Fields("TM20") '註冊日
                        MSHFlexGrid1.TextMatrix(p_Recs1, 24) = "" & RsTemp.Fields("TM77") '畫面上定稿語文
                        strExc(0) = "SELECT * FROM CaseProgress " & _
                                          "WHERE CP01='" & MSHFlexGrid1.TextMatrix(p_Recs1, 14) & "' " & _
                                          "and CP02='" & MSHFlexGrid1.TextMatrix(p_Recs1, 15) & "' " & _
                                          "and CP03='" & MSHFlexGrid1.TextMatrix(p_Recs1, 16) & "' " & _
                                          "and CP04='" & MSHFlexGrid1.TextMatrix(p_Recs1, 17) & "' " & _
                                          "and CP10='101'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           MSHFlexGrid1.TextMatrix(p_Recs1, 18) = "" & RsTemp.Fields("CP09")
                           MSHFlexGrid1.TextMatrix(p_Recs1, 19) = "" & RsTemp.Fields("CP10")
                        Else
                           If strErrNot101 <> "" Then strErrNot101 = strErrNot101 & "、"
                           strErrNot101 = strErrNot101 & MSHFlexGrid1.TextMatrix(p_Recs1, 14) & MSHFlexGrid1.TextMatrix(p_Recs1, 15)
                        End If
                     Else
                        p_Recs1_N = p_Recs1_N + 1
                     End If
                  End If
                  strErrText = "" 'Add By Sindy 2021/12/23
                  iRow = iRow + 1
               Loop
               '關閉
               xlsSalesPoint.Workbooks.Close
ReadNextFile:
            Next dblFCnt
            '離開
            xlsSalesPoint.Quit
            Set wksaccrpt114 = Nothing
            Set xlsSalesPoint = Nothing
            Screen.MousePointer = vbDefault
            If strErrNot101 <> "" Then
               MsgBox "下列本所案號(" & strErrNot101 & ")均無申請案資料，請先假收文申請案後，再重新執行該作業。"
               Exit Function
            End If
'            Screen.MousePointer = vbHourglass
'            Set fs = CreateObject("Scripting.FileSystemObject")
'            Set f = fs.OpenTextFile(txtPath(0), ForReading, TristateFalse)
'            ResetGrid MSHFlexGrid1, 0
'            Do While f.AtEndOfLine <> True
'               intRow = intRow + 1
'               strText = f.ReadLine
'               If intRow > 1 And Left(strText, InStr(strText, vbTab) - 1) <> "" Then
'                  p_Recs1 = p_Recs1 + 1
'                  MSHFlexGrid1.Rows = p_Recs1 + 1
'                  For i = 0 To 1
'                     intTab = InStr(strText, vbTab)
'                     If i = 1 Then '最後一個欄位
'                        strTemp = Trim(strText)
'                     Else
'                        strTemp = Trim(Mid(strText, 1, intTab - 1))
'                        strText = Mid(strText, intTab + 1, Len(strText))
'                     End If
'                     If i = 1 Then '日期欄位
'                        strTemp = ChangeWDateStringToWString(strTemp)
'                     End If
'                     If i = 0 Then MSHFlexGrid1.TextMatrix(p_Recs1, 1) = strTemp '註冊號
'                     If i = 1 Then MSHFlexGrid1.TextMatrix(p_Recs1, 4) = strTemp '申請日期
'                  Next i
'                  '是否為本所案件
'                  strExc(0) = "SELECT * FROM TradeMark WHERE TM12='" & MSHFlexGrid1.TextMatrix(p_Recs1, 1) & "' and TM10='020' and TM28='1'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI <> 1 And Len(Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1))) < 8 Then
'                     '補足8碼檢核
'                     strTemp = Right("00000000" & Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1)), 8)
'                     strExc(0) = "SELECT * FROM TradeMark WHERE TM12='" & strTemp & "' and TM10='020' and TM28='1'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        MSHFlexGrid1.TextMatrix(p_Recs1, 1) = strTemp
'                     End If
'                  End If
'                  If intI = 1 Then
'                     p_Recs1_Y = p_Recs1_Y + 1
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 13) = "Y"
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 14) = "" & RsTemp.Fields("TM01")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 15) = "" & RsTemp.Fields("TM02")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 16) = "" & RsTemp.Fields("TM03")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 17) = "" & RsTemp.Fields("TM04")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 20) = "" & RsTemp.Fields("TM10")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 21) = "" & RsTemp.Fields("TM11")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 22) = "" & RsTemp.Fields("TM23")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 23) = "" & RsTemp.Fields("TM20")
'                     MSHFlexGrid1.TextMatrix(p_Recs1, 24) = "" & RsTemp.Fields("TM77")
'                     strExc(0) = "SELECT * FROM CaseProgress " & _
'                                       "WHERE CP01='" & MSHFlexGrid1.TextMatrix(p_Recs1, 14) & "' " & _
'                                       "and CP02='" & MSHFlexGrid1.TextMatrix(p_Recs1, 15) & "' " & _
'                                       "and CP03='" & MSHFlexGrid1.TextMatrix(p_Recs1, 16) & "' " & _
'                                       "and CP04='" & MSHFlexGrid1.TextMatrix(p_Recs1, 17) & "' " & _
'                                       "and CP10='101'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        MSHFlexGrid1.TextMatrix(p_Recs1, 18) = "" & RsTemp.Fields("CP09")
'                        MSHFlexGrid1.TextMatrix(p_Recs1, 19) = "" & RsTemp.Fields("CP10")
'                     Else
'                        If strErrNot101 <> "" Then strErrNot101 = strErrNot101 & "、"
'                        strErrNot101 = strErrNot101 & MSHFlexGrid1.TextMatrix(p_Recs1, 14) & MSHFlexGrid1.TextMatrix(p_Recs1, 15)
'                     End If
'                  Else
'                     p_Recs1_N = p_Recs1_N + 1
'                  End If
'               End If
'            Loop
'            f.Close
'            Screen.MousePointer = vbDefault
'            If strErrNot101 <> "" Then
'               MsgBox "下列本所案號(" & strErrNot101 & ")均無申請案資料，請假收文申請案後，再重新執行該作業。"
'               Exit Function
'            End If
         End If
         '2012/9/24 End
         
         '台灣申請人案件(第二次)
         If Option2(1).Value = True Then
            Screen.MousePointer = vbHourglass
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.OpenTextFile(txtPath(0), ForReading, TristateFalse)
            ResetGrid MSHFlexGrid1, 0
            Do While f.AtEndOfLine <> True
               intRow = intRow + 1
               strText = f.ReadLine
               If intRow > 1 And Left(strText, InStr(strText, vbTab) - 1) <> "" Then
                  p_Recs1 = p_Recs1 + 1
                  MSHFlexGrid1.Rows = p_Recs1 + 1
                  For i = 0 To 12 '14
                     intTab = InStr(strText, vbTab)
                     If i = 13 Then '最後一個欄位 '14
                        strTemp = Trim(strText)
                     Else
                        strTemp = Trim(Mid(strText, 1, intTab - 1))
                        strText = Mid(strText, intTab + 1, Len(strText))
                        If i = 3 Then '商標名稱(中)
                           strNameC = strTemp
                        ElseIf i = 4 Then '商標名稱(英)
                           strNameE = strTemp
                        End If
                     End If
   '                  If i = 1 Then '註冊號
   '                     'Modify By Sindy 2010/9/1
   '                     If bolNewAppNoFormat Then
   '                        strTemp = Right("0000000" & strTemp, 7)
   '                     '2010/9/1 End
   '                     Else
   '                        strTemp = Right("00000000" & strTemp, 8)
   '                     End If
   '                  Else
                     If i = 5 Then '商標註冊人 : 過濾英數字只取得中文字
                        strChkData = strTemp
                        strTemp = ""
                        For j = 1 To Len(strChkData)
                           strWord = Mid(strChkData, j, 1)
                           If Asc(strWord) <> 34 Then 'Add By Sindy 2014/11/28 +if 文字裡頭名稱前面有"符號
                              If IsNumeric(strWord) = True Or (Asc(strWord) >= 0 And Asc(strWord) <= 255) Then
                                 MSHFlexGrid1.TextMatrix(p_Recs1, 25) = Mid(strChkData, j, Len(strChkData) - j + 1) '個人ID
                                 Exit For
                              Else
                                 strTemp = strTemp & strWord
                              End If
                           End If
                        Next j
                        strTemp = CheckStr(strTemp)
                     ElseIf i = 6 Or i = 7 Or i = 8 Then '日期欄位
                        strTemp = ChangeWDateStringToWString(strTemp)
'                     ElseIf i = 9 Then '國際分類
'                        strTemp = Right("00" & Trim(strTemp), 2)
'                        If Asc(Left(strTemp, 1)) = 34 Then
'                           strTemp = Mid(strTemp, 2, Len(strTemp))
'                        End If
'                        If Asc(Right(strTemp, 1)) = 34 Then
'                           strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
'                        End If
'                        If strTemp = "00" Then strTemp = ""
                     ElseIf i = 12 Then '申請人地址 '11
                        strTemp = ToWide(CheckStr(strTemp))
                     End If
                     If i = 0 Then MSHFlexGrid1.TextMatrix(p_Recs1, 0) = strTemp '選擇
                     'If i = 2 Then MSHFlexGrid1.TextMatrix(p_Recs1, 1) = strTemp '註冊號
                     If i = 1 Then MSHFlexGrid1.TextMatrix(p_Recs1, 1) = strTemp '註冊號
                     If i = 4 Then MSHFlexGrid1.TextMatrix(p_Recs1, 2) = strNameC + strNameE '商標名稱
                     If i = 5 Then MSHFlexGrid1.TextMatrix(p_Recs1, 3) = strTemp '商標註冊人
                     If i = 6 Then MSHFlexGrid1.TextMatrix(p_Recs1, 4) = strTemp '申請日期
                     If i = 7 Then MSHFlexGrid1.TextMatrix(p_Recs1, 5) = strTemp '專用期(起)
                     If i = 8 Then MSHFlexGrid1.TextMatrix(p_Recs1, 6) = strTemp '專用期(止)
                     'If i = 1 Then MSHFlexGrid1.TextMatrix(p_Recs1, 7) = strTemp '國際分類
                     If i = 2 Then MSHFlexGrid1.TextMatrix(p_Recs1, 7) = strTemp '國際分類
                     'If i = 9 Then MSHFlexGrid1.TextMatrix(p_Recs1, 8) = strTemp '最後公告
                     If i = 10 Then MSHFlexGrid1.TextMatrix(p_Recs1, 8) = strTemp '最後公告
                     'If i = 11 Then MSHFlexGrid1.TextMatrix(p_Recs1, 9) = strTemp '申請人地址
                     If i = 12 Then MSHFlexGrid1.TextMatrix(p_Recs1, 9) = strTemp '申請人地址
                     'If i = 12 Then MSHFlexGrid1.TextMatrix(p_Recs1, 10) = strTemp '初審公告
                     'If i = 13 Then MSHFlexGrid1.TextMatrix(p_Recs1, 11) = strTemp '商標類型
                     'If i = 14 Then MSHFlexGrid1.TextMatrix(p_Recs1, 12) = strTemp '代理機構
                  Next i
                  '是否為本所案件
                  'Modify By Sindy 2010/12/8 增加TM28='1'卷宗性質作判斷
                  'Modify By Sindy 2011/12/8 註冊號數先用原資料檢核,若Find不到資料,再看是否有須要補足碼數再檢核一次
                  strExc(0) = "SELECT * FROM TradeMark WHERE TM12='" & MSHFlexGrid1.TextMatrix(p_Recs1, 1) & "' and TM10='020' and TM28='1'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI <> 1 And Len(Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1))) < 8 Then
                     '補足8碼檢核
                     strTemp = Right("00000000" & Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1)), 8)
                     strExc(0) = "SELECT * FROM TradeMark WHERE TM12='" & strTemp & "' and TM10='020' and TM28='1'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        MSHFlexGrid1.TextMatrix(p_Recs1, 1) = strTemp
                     End If
                  End If
                  If intI = 1 Then
                     p_Recs1_Y = p_Recs1_Y + 1
                     m_CaseNo = m_CaseNo & "　　　　　　　" & RsTemp.Fields("TM01") & "-" & RsTemp.Fields("TM02") & "-" & RsTemp.Fields("TM03") & "-" & RsTemp.Fields("TM04") & "（" & Trim(MSHFlexGrid1.TextMatrix(p_Recs1, 1)) & "）" & vbCrLf 'Add By Sindy 2012/10/15
                     MSHFlexGrid1.TextMatrix(p_Recs1, 13) = "Y"
                     MSHFlexGrid1.TextMatrix(p_Recs1, 14) = "" & RsTemp.Fields("TM01")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 15) = "" & RsTemp.Fields("TM02")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 16) = "" & RsTemp.Fields("TM03")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 17) = "" & RsTemp.Fields("TM04")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 20) = "" & RsTemp.Fields("TM10")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 21) = "" & RsTemp.Fields("TM11")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 22) = "" & RsTemp.Fields("TM23")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 23) = "" & RsTemp.Fields("TM20")
                     MSHFlexGrid1.TextMatrix(p_Recs1, 24) = "" & RsTemp.Fields("TM77")
                     strExc(0) = "SELECT * FROM CaseProgress " & _
                                       "WHERE CP01='" & MSHFlexGrid1.TextMatrix(p_Recs1, 14) & "' " & _
                                       "and CP02='" & MSHFlexGrid1.TextMatrix(p_Recs1, 15) & "' " & _
                                       "and CP03='" & MSHFlexGrid1.TextMatrix(p_Recs1, 16) & "' " & _
                                       "and CP04='" & MSHFlexGrid1.TextMatrix(p_Recs1, 17) & "' " & _
                                       "and CP10='101'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        MSHFlexGrid1.TextMatrix(p_Recs1, 18) = "" & RsTemp.Fields("CP09")
                        MSHFlexGrid1.TextMatrix(p_Recs1, 19) = "" & RsTemp.Fields("CP10")
                     'Add By Sindy 2010/10/20
                     Else
                        If strErrNot101 <> "" Then strErrNot101 = strErrNot101 & "、"
                        strErrNot101 = strErrNot101 & MSHFlexGrid1.TextMatrix(p_Recs1, 14) & MSHFlexGrid1.TextMatrix(p_Recs1, 15)
                     '2010/10/20 End
                     End If
                  Else
                     p_Recs1_N = p_Recs1_N + 1
                  End If
               End If
            Loop
            f.Close
            Screen.MousePointer = vbDefault
'            'Add By Sindy 2010/10/20
'            If strErrNot101 <> "" Then
'               MsgBox "下列本所案號(" & strErrNot101 & ")均無申請案資料，請假收文申請案後，再重新執行該作業。"
'               Exit Function
'            End If
'            '2010/10/20 End
         End If
      End If
   End If
   
'通知續展:
'　序號 , 註冊號, 商標名稱, 商標註冊人, 申請日期, 專用期(起), 專用期(止),
'　國際分類, 公告狀態，申請人地址
   If Option1(1).Value = True Then
      If Dir(txtPath(1)) = "" Then
'         If MsgBox("通知續展檔案不存在，是否繼續！", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
'            txtPath(1).SetFocus
'            Exit Function
'         End If
         MsgBox "通知續展檔案不存在！", vbExclamation
         txtPath(1).SetFocus
         Exit Function
      Else
         Screen.MousePointer = vbHourglass
         Set fs = CreateObject("Scripting.FileSystemObject")
         Set f = fs.OpenTextFile(txtPath(1), ForReading, TristateFalse)
         ResetGrid MSHFlexGrid2, 1
         Do While f.AtEndOfLine <> True
            intRow = intRow + 1
            strText = f.ReadLine
            If intRow > 1 And Left(strText, InStr(strText, vbTab) - 1) <> "" Then
               p_Recs2 = p_Recs2 + 1
               MSHFlexGrid2.Rows = p_Recs2 + 1
               For i = 0 To 10
                  intTab = InStr(strText, vbTab)
                  If i = 10 Then '最後一個欄位
                     strTemp = ToWide(CheckStr(Trim(strText))) '申請人地址
                  Else
                     strTemp = Trim(Mid(strText, 1, intTab - 1))
                     strText = Mid(strText, intTab + 1, Len(strText))
                     If i = 2 Then '商標名稱(中)
                        strNameC = strTemp
                        If strNameC = "" Then
'                           MsgBox strNameC
                        End If
                     ElseIf i = 3 Then '商標名稱(英)
                        strNameE = strTemp
                     End If
                  End If
'                  If i = 1 Then '註冊號
'                     'Modify By Sindy 2010/9/1
'                     If bolNewAppNoFormat Then
'                        strTemp = Right("0000000" & strTemp, 7)
'                     '2010/9/1 End
'                     Else
'                        strTemp = Right("00000000" & strTemp, 8)
'                     End If
'                  Else
                  If i = 4 Then '商標註冊人 : 過濾英數字只取得中文字
                     strChkData = strTemp
                     strTemp = ""
                     For j = 1 To Len(strChkData)
                        strWord = Mid(strChkData, j, 1)
                        If Asc(strWord) <> 34 Then 'Add By Sindy 2014/11/28 +if 文字裡頭名稱前面有"符號
                           If IsNumeric(strWord) = True Or (Asc(strWord) >= 0 And Asc(strWord) <= 255) Then
                              MSHFlexGrid2.TextMatrix(p_Recs2, 22) = Mid(strChkData, j, Len(strChkData) - j + 1) '個人ID
                              Exit For
                           Else
                              strTemp = strTemp & strWord
                           End If
                        End If
                     Next j
                     strTemp = CheckStr(strTemp)
                  ElseIf i = 5 Or i = 6 Or i = 7 Then '日期欄位
                     strTemp = Replace(strTemp, "年", "/")
                     strTemp = Replace(strTemp, "月", "/")
                     strTemp = Replace(strTemp, "日", "")
                     strTemp = Replace(strTemp, "-", "/") 'Add By Sindy 2013/10/23 ex.2004-10-7
                     strTemp = ChangeWDateStringToWString(strTemp)
                  ElseIf i = 8 Then '國際分類
                     strTemp = Right("00" & Trim(strTemp), 2)
                     If Asc(Left(strTemp, 1)) = 34 Then
                        strTemp = Mid(strTemp, 2, Len(strTemp))
                     End If
                     If Asc(Right(strTemp, 1)) = 34 Then
                        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
                     End If
                     If strTemp = "00" Then strTemp = ""
                  End If
                  If i = 0 Then MSHFlexGrid2.TextMatrix(p_Recs2, 0) = strTemp '序號
                  If i = 1 Then MSHFlexGrid2.TextMatrix(p_Recs2, 1) = strTemp '註冊號
                  If i = 3 Then MSHFlexGrid2.TextMatrix(p_Recs2, 2) = strNameC + strNameE '商標名稱
                  If i = 4 Then MSHFlexGrid2.TextMatrix(p_Recs2, 3) = strTemp '商標註冊人
                  If i = 5 Then MSHFlexGrid2.TextMatrix(p_Recs2, 4) = strTemp '申請日期
                  If i = 6 Then MSHFlexGrid2.TextMatrix(p_Recs2, 5) = strTemp '專用期(起)
                  If i = 7 Then MSHFlexGrid2.TextMatrix(p_Recs2, 6) = strTemp '專用期(止)
                  If i = 8 Then MSHFlexGrid2.TextMatrix(p_Recs2, 7) = strTemp '國際分類
                  If i = 9 Then MSHFlexGrid2.TextMatrix(p_Recs2, 8) = strTemp '最近公告
                  If i = 10 Then MSHFlexGrid2.TextMatrix(p_Recs2, 9) = strTemp '申請人地址
               Next i
               '是否為本所案件
               'Modify By Sindy 2010/12/8 增加TM28='1'卷宗性質作判斷
               'Modify By Sindy 2011/12/8 註冊號數先用原資料檢核,若Find不到資料,再看是否有須要補足碼數再檢核一次
               strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & MSHFlexGrid2.TextMatrix(p_Recs2, 1) & "' and TM10='020' and TM28='1'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI <> 1 And Len(Trim(MSHFlexGrid2.TextMatrix(p_Recs2, 1))) < 8 Then
                  '補足8碼檢核
                  strTemp = Right("00000000" & Trim(MSHFlexGrid2.TextMatrix(p_Recs2, 1)), 8)
                  strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & strTemp & "' and TM10='020' and TM28='1'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     MSHFlexGrid2.TextMatrix(p_Recs2, 1) = strTemp
                  End If
               End If
'               If MSHFlexGrid2.TextMatrix(p_Recs2, 1) = "620048" Then
'                  MsgBox MSHFlexGrid2.TextMatrix(p_Recs2, 1)
'               End If
               If intI = 1 Then
                  p_Recs2_Y = p_Recs2_Y + 1
                  MSHFlexGrid2.TextMatrix(p_Recs2, 10) = "Y"
                  MSHFlexGrid2.TextMatrix(p_Recs2, 11) = "" & RsTemp.Fields("TM01")
                  MSHFlexGrid2.TextMatrix(p_Recs2, 12) = "" & RsTemp.Fields("TM02")
                  MSHFlexGrid2.TextMatrix(p_Recs2, 13) = "" & RsTemp.Fields("TM03")
                  MSHFlexGrid2.TextMatrix(p_Recs2, 14) = "" & RsTemp.Fields("TM04")
                  MSHFlexGrid2.TextMatrix(p_Recs2, 15) = "" & RsTemp.Fields("TM11")
                  MSHFlexGrid2.TextMatrix(p_Recs2, 16) = "" & RsTemp.Fields("TM22")
                  MSHFlexGrid2.TextMatrix(p_Recs2, 17) = "" & RsTemp.Fields("TM29")
                  '專用期止日大於0時
                  If Val(MSHFlexGrid2.TextMatrix(p_Recs2, 16)) > 0 Then
                     '下一程序有延展且法定期限=專用期止日
                     strExc(0) = "SELECT * FROM NextProgress " & _
                                    "WHERE NP02='" & MSHFlexGrid2.TextMatrix(p_Recs2, 11) & "' " & _
                                    "and NP03='" & MSHFlexGrid2.TextMatrix(p_Recs2, 12) & "' " & _
                                    "and NP04='" & MSHFlexGrid2.TextMatrix(p_Recs2, 13) & "' " & _
                                    "and NP05='" & MSHFlexGrid2.TextMatrix(p_Recs2, 14) & "' " & _
                                    "and NP07='102' " & _
                                    "and NP09=" & MSHFlexGrid2.TextMatrix(p_Recs2, 16)
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        MSHFlexGrid2.TextMatrix(p_Recs2, 18) = "" & RsTemp.Fields("NP01")
                        MSHFlexGrid2.TextMatrix(p_Recs2, 19) = "" & RsTemp.Fields("NP06")
                        MSHFlexGrid2.TextMatrix(p_Recs2, 20) = "" & RsTemp.Fields("NP07")
                        'N不辦或Y閉卷=已結案
                        If Trim(MSHFlexGrid2.TextMatrix(p_Recs2, 19)) = "N" Or _
                           Trim(MSHFlexGrid2.TextMatrix(p_Recs2, 17)) = "Y" Then
                           p_Recs2_Close = p_Recs2_Close + 1
                        'Y續辦=已收文
                        ElseIf Trim(MSHFlexGrid2.TextMatrix(p_Recs2, 19)) = "Y" Then
                           p_Recs2_Recv = p_Recs2_Recv + 1
                           MSHFlexGrid2.TextMatrix(p_Recs2, 21) = "Y"
                        End If
                     Else
                        Call ChkRecs2_Recv
                     End If
                  Else
                     Call ChkRecs2_Recv
                  End If
               Else
                  p_Recs2_N = p_Recs2_N + 1
               End If
            End If
         Loop
         f.Close
         Screen.MousePointer = vbDefault
      End If
   End If
   
   LoadXLS = True
   
   Exit Function
   
ErrHnd:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      If strErrText <> "" Then
         '關閉
         xlsSalesPoint.Workbooks.Close
         '離開
         xlsSalesPoint.Quit
         Set wksaccrpt114 = Nothing
         Set xlsSalesPoint = Nothing
      End If
      MsgBox Err.Description & vbCrLf & vbCrLf & _
             strErrText, vbCritical
   End If
End Function

'台灣申請人案件(第二次)
Private Function Process1() As Boolean
Dim iRow As Integer, iRecs As Integer, iXRow As Integer
Dim bPrint As Boolean, i As Integer, bPrintAgain As Boolean
Dim strTempAppAddr As String, strCU01 As String 'Add By Sindy 2010/11/18
Dim rsTmp As New ADODB.Recordset, bolErr2147467259 As Boolean 'Add By Sindy 2011/12/20
   
   On Error GoTo ErrHnd
   
'   Screen.MousePointer = vbHourglass
'   bPrintAgain = False
'PrintAgain3:
'   bPrint = False
'   'Set Printer = Printers(Combo1.ListIndex)
'   'Printer.EndDoc
'   Printer.Orientation = 2 '1.直印 2.橫印
'   For iRow = 1 To MSHFlexGrid1.Rows - 1
'      '本所
'      If MSHFlexGrid1.TextMatrix(iRow, 13) = "Y" Then
'         If MSHFlexGrid1.TextMatrix(iRow, 4) <> MSHFlexGrid1.TextMatrix(iRow, 21) Then
'            '檢查申請日是否相同, 不同者列印清單
'            For i = 1 To 6
'               strTemp(i) = ""
'            Next i
'            strTemp(1) = MSHFlexGrid1.TextMatrix(iRow, 14) & "-" & MSHFlexGrid1.TextMatrix(iRow, 15) & "-" & MSHFlexGrid1.TextMatrix(iRow, 16) & "-" & MSHFlexGrid1.TextMatrix(iRow, 17)
'            strTemp(2) = MSHFlexGrid1.TextMatrix(iRow, 1)
'            strTemp(3) = MSHFlexGrid1.TextMatrix(iRow, 2)
'            strTemp(4) = MSHFlexGrid1.TextMatrix(iRow, 3)
'            strTemp(5) = ChangeWStringToWDateString(MSHFlexGrid1.TextMatrix(iRow, 4))
'            If iLine > 37 Or bPrint = False Then
'               If bPrint <> False Then Printer.NewPage
'               iLine = 1
'               PrintTitle '列印表頭
'            End If
'            PrintDetail
'            bPrint = True
'         Else
'            If bPrintAgain = False Then '重新列印清單時, 資料異動不可再重覆執行
'                  '檢查是否已有(1102核准通知)或(1403改變原處分並且實際結果為核准), 若有, 不出定稿不異動資料, 計算筆數
'                  strSql = "select * " & _
'                              "From caseprogress " & _
'                              "where cp01='" & MSHFlexGrid1.TextMatrix(iRow, 14) & "' " & _
'                              "and cp02='" & MSHFlexGrid1.TextMatrix(iRow, 15) & "' " & _
'                              "and cp03='" & MSHFlexGrid1.TextMatrix(iRow, 16) & "' " & _
'                              "and cp04='" & MSHFlexGrid1.TextMatrix(iRow, 17) & "' " & _
'                              "and (cp10='1102' or (cp10='1403' and cp24='1')) "
'                  intI = 1
'                  Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     p_Recs1_1102 = p_Recs1_1102 + 1
'                  Else
'                     frm02010401_4.m_TM01 = MSHFlexGrid1.TextMatrix(iRow, 14)
'                     frm02010401_4.m_TM02 = MSHFlexGrid1.TextMatrix(iRow, 15)
'                     frm02010401_4.m_TM03 = MSHFlexGrid1.TextMatrix(iRow, 16)
'                     frm02010401_4.m_TM04 = MSHFlexGrid1.TextMatrix(iRow, 17)
'                     frm02010401_4.m_CP09 = MSHFlexGrid1.TextMatrix(iRow, 18)
'                     frm02010401_4.m_CP10 = MSHFlexGrid1.TextMatrix(iRow, 19)
'                     frm02010401_4.m_TM10 = MSHFlexGrid1.TextMatrix(iRow, 20)
'                     frm02010401_4.m_TM11 = MSHFlexGrid1.TextMatrix(iRow, 21)
'                     frm02010401_3.textResult = "1" '核准
'                     frm02010401_4.textCP25 = strSrvDate(2) '核准通知日
'                     frm02010401_4.textTM15 = MSHFlexGrid1.TextMatrix(iRow, 1) '審定號
'                     frm02010401_4.textTM14 = text1(0) '公告日
'                     frm02010401_4.textTMBM07_2 = text1(1) '期數
'                     frm02010401_4.m_CP05 = strSrvDate(2) '來函收文日
'                     frm02010401_4.m_TM23 = MSHFlexGrid1.TextMatrix(iRow, 22)
'                     frm02010401_4.m_TM20 = MSHFlexGrid1.TextMatrix(iRow, 23)
'                     frm02010401_4.textPrint = MSHFlexGrid1.TextMatrix(iRow, 24)
'                     '帶列印定稿預設值
'                     If frm02010401_4.textPrint = "" Then
'                        frm02010401_4.textPrint = GetTWordLng(MSHFlexGrid1.TextMatrix(iRow, 14), MSHFlexGrid1.TextMatrix(iRow, 15), MSHFlexGrid1.TextMatrix(iRow, 16), MSHFlexGrid1.TextMatrix(iRow, 17))
'                     End If
'                     frm02010401_4.m_CP14 = strUserNum
'
'                     '檢查是否已有1001.核准, 若有, 代表重覆執行不異動資料出定稿
'                     strSql = "select * from caseprogress " & _
'                                  "where cp09 in (select cp43 From caseprogress " & _
'                                  "where cp01='" & MSHFlexGrid1.TextMatrix(iRow, 14) & "' " & _
'                                  "and cp02='" & MSHFlexGrid1.TextMatrix(iRow, 15) & "' " & _
'                                  "and cp03='" & MSHFlexGrid1.TextMatrix(iRow, 16) & "' " & _
'                                  "and cp04='" & MSHFlexGrid1.TextMatrix(iRow, 17) & "' " & _
'                                  "and cp10='1001' " & _
'                                  "and cp43 is not null) " & _
'                                  "and cp10='101' "
'                     intI = 1
'                     Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'                     If intI <> 1 Then
'                        If frm02010401_4.OnSaveData = False Then
'                           MsgBox MSHFlexGrid1.TextMatrix(iRow, 1) & "存檔失敗，請洽系統管理員 !", vbCritical
'                           Screen.MousePointer = vbDefault
'                           Exit Function
'                        End If
'                     End If
'                     adoRecordset.Close
'                     '列印定稿
'                     frm02010401_4.PrintLetter
'                     Unload frm02010401_3
'                     Unload frm02010401_4
'                     'Add By Sindy 2012/8/20
'                     Dim strSales As String, strSales_cc As String
'                     Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
'                     strSql = "select cp01,cp02,cp03,cp04,tm12,tm15,tm05,tm09,tm14,DECODE(CU15,'0','台端','1','貴公司','貴單位') as cu15Nm from caseprogress,trademark,customer " & _
'                                  "where cp09 in (select cp43 From caseprogress " & _
'                                  "where cp01='" & MSHFlexGrid1.TextMatrix(iRow, 14) & "' " & _
'                                  "and cp02='" & MSHFlexGrid1.TextMatrix(iRow, 15) & "' " & _
'                                  "and cp03='" & MSHFlexGrid1.TextMatrix(iRow, 16) & "' " & _
'                                  "and cp04='" & MSHFlexGrid1.TextMatrix(iRow, 17) & "' " & _
'                                  "and cp10='1001' " & _
'                                  "and cp43 is not null) " & _
'                                  "and cp10='101' " & _
'                                  "and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
'                                  "and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
'                     intI = 1: strSales = "": strSales_cc = ""
'                     Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 Then
'                        m_CP01 = adoRecordset.Fields("cp01")
'                        m_CP02 = adoRecordset.Fields("cp02")
'                        m_CP03 = adoRecordset.Fields("cp03")
'                        m_CP04 = adoRecordset.Fields("cp04")
'                        '讀取智權人員
'                        strSales = PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)
'                        '若為68096中三杜副總時，檢查客戶若最後收文為在職人員則設為副本收件人
'                        If strSales = "68096" Then
'                           strExc(0) = "select st01 from staff,(select max(cp05||cp13) cp13 from ( " & _
'                              "      Select cp05,cp13 From patent, caseprogress Where pa26='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp13 From trademark, caseprogress Where tm23='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp13 From lawcase, caseprogress Where lc11='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp13 From servicepractice, caseprogress Where sp08='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp13 From hirecase, caseprogress Where hc05='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp09<'B' " & _
'                              ")) aa where substr(aa.cp13,9)=st01(+) and st04='1'"
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                           If intI = 1 Then
'                              strSales_cc = RsTemp.Fields(0).Value
'                           End If
'                        End If
'                        '寄發智權同仁由同仁轉客戶
'                        PUB_SendMail strUserNum, strSales, "", "大陸商標核准通知（本所案號：" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "）", _
'                        "敬啟者：" & vbCrLf & vbCrLf & _
'                        "　　" & adoRecordset.Fields("cu15Nm") & "委託本所辦理之第" & IIf(adoRecordset.Fields("tm15") = "", adoRecordset.Fields("tm12"), adoRecordset.Fields("tm15")) & "號「" & adoRecordset.Fields("tm05") & "」（第" & adoRecordset.Fields("tm09") & "類）大陸商標註冊申請案，業經審查核准，公告於" & Left(adoRecordset.Fields("tm14"), 4) - 1911 & "年" & Mid(adoRecordset.Fields("tm14"), 5, 2) & "月" & Right(adoRecordset.Fields("tm14"), 2) & "日之大陸商標公報，公告資料將另函郵寄予　" & adoRecordset.Fields("cu15Nm") & "，請留意查收。" & vbCrLf & vbCrLf & _
'                        "台一國際專利商標事務所　敬上", "", , , , , strSales_cc
'                     End If
'                     adoRecordset.Close
'                     '2012/8/20 End
'                  End If
'            End If
'         End If
'      End If
'   Next
'   If bPrint = True Then
'      If bPrintAgain = False Then MsgBox "將列印審定公告資料檢核清單，請換一般列印紙!!!"
'      Printer.EndDoc
'      If MsgBox("列印審定公告資料檢核清單，是否已列印成功？成功按「是」，不成功按「否」再列印一次！", vbYesNo + vbDefaultButton2) = vbNo Then
'         bPrintAgain = True
'         GoTo PrintAgain3
'      End If
'   End If
'   Screen.MousePointer = vbDefault
   
   '產生Word檔
   bolRetry = True
   Screen.MousePointer = vbHourglass
'   'Add By Sindy 2014/2/19
'   cnnConnection.BeginTrans
'   cnnConnection.Execute "delete from R020320"
'   '2014/2/19 END
   For iRow = 1 To MSHFlexGrid1.Rows - 1
      '非本所
      'Modify By Sindy 2013/2/25 增加排除不可列印者
      'If MSHFlexGrid1.TextMatrix(iRow, 13) <> "Y" Then
      If MSHFlexGrid1.TextMatrix(iRow, 13) <> "Y" Then
      '2013/2/25 End
         m_AppAddr = MSHFlexGrid1.TextMatrix(iRow, 9) '申請人地址
         m_AppName = MSHFlexGrid1.TextMatrix(iRow, 3) '商標註冊人
         If GetIsNotPrintPer(m_AppName) = False Then 'Add By Sindy 2013/3/6 +if
            'Add By Sindy 2010/11/18
            strTempAppAddr = "": strCU01 = "": m_AppAddrZip = ""
            'Modify By Sindy 2012/6/19 註冊號9727839申請人為吳志忠此案件之故,再增加名稱若小於等於4個字的也是個人
            If Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) <> "" Or Len(Trim(MSHFlexGrid1.TextMatrix(iRow, 3))) <= 4 Then
               '個人, 抓名稱及ID都相同者,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
   '            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' and cu11='" & Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) & "' order by cu01 asc,cu02 asc "
   '            intI = 1
   '            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '            If intI = 1 Then
   '               If RsTemp.Fields("cu02") <> "0" Then
   '                  strCU01 = "" & RsTemp.Fields("cu01")
   '               Else
   '                  m_AppAddrZip = Trim("" & RsTemp.Fields("cu30"))
   '                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
   '               End If
   '            End If
               'Modify By Sindy 2011/12/20
               strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' and cu11='" & Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) & "' order by cu01 asc,cu02 asc "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If bolErr2147467259 = True Then
                  bolErr2147467259 = False
   '               rsTmp.Close
                  strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & " ' and cu11='" & Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) & "' order by cu01 asc,cu02 asc "
                  rsTmp.CursorLocation = adUseClient
                  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               End If
               If rsTmp.RecordCount > 0 Then
                  If rsTmp.Fields("cu02") <> "0" Then
                     strCU01 = "" & rsTmp.Fields("cu01")
                  Else
                     m_AppAddrZip = PUB_ChangeZIPToSir(Trim("" & rsTmp.Fields("cu30")))
                     strTempAppAddr = Trim("" & rsTmp.Fields("cu31"))
                  End If
               End If
               rsTmp.Close
            Else
               '非個人, 抓相同客戶名稱之最小客戶編號,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
   '            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc "
   '            intI = 1
   '            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '            If intI = 1 Then
   '               If RsTemp.Fields("cu02") <> "0" Then
   '                  strCU01 = "" & RsTemp.Fields("cu01")
   '               Else
   '                  m_AppAddrZip = Trim("" & RsTemp.Fields("cu30"))
   '                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
   '               End If
   '            End If
               'Modify By Sindy 2011/12/20
               strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If bolErr2147467259 = True Then
                  bolErr2147467259 = False
   '               rsTmp.Close
                  strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & " ' order by cu01 asc,cu02 asc "
                  rsTmp.CursorLocation = adUseClient
                  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               End If
               If rsTmp.RecordCount > 0 Then
                  If rsTmp.Fields("cu02") <> "0" Then
                     strCU01 = "" & rsTmp.Fields("cu01")
                  Else
                     m_AppAddrZip = PUB_ChangeZIPToSir(Trim("" & rsTmp.Fields("cu30")))
                     strTempAppAddr = Trim("" & rsTmp.Fields("cu31"))
                  End If
               End If
               rsTmp.Close
            End If
            If strCU01 <> "" Then
               strSql = "SELECT * FROM customer WHERE cu01='" & strCU01 & "' and cu02='0' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  m_AppAddrZip = PUB_ChangeZIPToSir(Trim("" & RsTemp.Fields("cu30")))
                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
               End If
            End If
            If strTempAppAddr <> "" Then
               m_AppAddr = strTempAppAddr
            Else
               m_AppAddrZip = ""
            End If
            '2010/11/18 End
            'Add By Sindy 2013/6/3
            If m_AppAddrZip = "" Then
               m_AppAddrZip = PUB_ChangeZIPToSir(Left(PUB_AddrChangeZIPCode(m_AppAddr), 3))
            End If
            '2013/6/3 End
            m_AppDate = MSHFlexGrid1.TextMatrix(iRow, 4) '申請日期
            m_AppDate = Val(Left(m_AppDate, 4)) & "年" & Mid(m_AppDate, 5, 2) & "月" & Right(m_AppDate, 2) & "日"
            m_AppNum = MSHFlexGrid1.TextMatrix(iRow, 1) '註冊號
            m_TName = MSHFlexGrid1.TextMatrix(iRow, 2) '商標名稱
            m_Goods = MSHFlexGrid1.TextMatrix(iRow, 7) '分類
            m_DATE = Val(Text1(0)) + 19110000 '公告日期
            m_DATE = Val(Left(m_DATE, 4)) & "年" & Mid(m_DATE, 5, 2) & "月" & Right(m_DATE, 2) & "日"
            m_Num = Text1(1) '期數
'            'Modify By Sindy 2014/2/19
'            strSql = "insert into R020320(AppZip,AppName,AppAddr,AppNum,TName,Goods,DATEs,DATEe,Num)" & _
'                     " values(" & CNULL(m_AppAddrZip) & "," & CNULL(m_AppName) & "," & CNULL(m_AppAddr) & "," & CNULL(m_AppNum) & _
'                     "," & CNULL(ChgSQL(m_TName)) & "," & CNULL(m_Goods) & "," & CNULL(m_AppDate, True) & "," & CNULL(m_DATE, True) & _
'                     "," & CNULL(m_Num) & ")"
'            cnnConnection.Execute strSql
            ' 列印定稿
            Call WordEdit(1)
'            '2014/2/19 END
         End If
      End If
   Next
'   cnnConnection.CommitTrans 'Add By Sindy 2014/2/19
'
'   'Add By Sindy 2014/2/19 同申請人多案一定稿
'   strSql = "select * from R020320 order by APPZIP,APPADDR,APPNAME asc"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   m_AppAddrZip = "": m_AppAddr = "": m_AppName = "": tmp_TName = ""
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         tmp_AppAddrZip = "" & rsTmp.Fields("APPZIP")
'         tmp_AppAddr = "" & rsTmp.Fields("AppAddr") '申請人地址
'         tmp_AppName = "" & rsTmp.Fields("APPNAME") '商標註冊人
'         m_AppNum = "" & rsTmp.Fields("APPNUM")
'         m_TName = "" & rsTmp.Fields("TNAME")
'         m_Goods = "" & rsTmp.Fields("GOODS")
'         '申請日期
'         m_AppDate = "" & rsTmp.Fields("DATES")
'         m_AppDate = Val(Left(m_AppDate, 4)) & "年" & Mid(m_AppDate, 5, 2) & "月" & Right(m_AppDate, 2) & "日"
'         '公告日期
'         m_DATE = "" & rsTmp.Fields("DATEE")
'         m_DATE = Val(Left(m_DATE, 4)) & "年" & Mid(m_DATE, 5, 2) & "月" & Right(m_DATE, 2) & "日"
'         m_Num = "" & rsTmp.Fields("Num") '期數
'         If Not (m_AppAddrZip = "" And m_AppAddr = "" And m_AppName = "") And _
'            (m_AppAddrZip <> tmp_AppAddrZip Or m_AppAddr <> tmp_AppAddr Or m_AppName <> tmp_AppName) Then
'            ' 列印定稿
'            tmp_TName = Right(tmp_TName, Len(tmp_TName) - 1)
'            Call WordEdit(4)
'            tmp_TName = ""
'         End If
'         m_AppAddrZip = tmp_AppAddrZip
'         m_AppAddr = tmp_AppAddr
'         m_AppName = tmp_AppName
'         tmp_TName = tmp_TName & "、" & m_AppDate & "申請之第" & m_AppNum & "號「" & m_TName & "」(第" & m_Goods & "類)"
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'   If m_AppAddrZip <> "" Or m_AppAddr <> "" Or m_AppName <> "" Then
'      ' 列印定稿
'      tmp_TName = Right(tmp_TName, Len(tmp_TName) - 1)
'      Call WordEdit(4)
'   End If
'   '2014/2/19 END
   
   If bolRetry = False Then
      g_WordAp.Visible = True
      g_WordAp.WindowState = wdWindowStateMaximize
   End If
   Screen.MousePointer = vbDefault
   
   Process1 = True
   Set g_WordAp = Nothing
   Set rsTmp = Nothing 'Add By Sindy 2011/12/20
   Exit Function
   
ErrHnd:
   'Add By Sindy 2011/12/20
   bolErr2147467259 = False
   If Err.Number = -2147467259 Then
      bolErr2147467259 = True
      '接著發生錯誤陳述式的下個陳述式開始執行
      Resume Next
   End If
   Set g_WordAp = Nothing
   Set rsTmp = Nothing
   '2011/12/20 End
   If Err.Number <> 0 Then
      'cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

'Add By Sindy 2012/9/24
'全部資料(第一次)
Private Function Process3() As Boolean
Dim iRow As Integer, iRecs As Integer, iXRow As Integer
Dim bPrint As Boolean, i As Integer, bPrintAgain As Boolean
Dim strTempAppAddr As String, strCU01 As String 'Add By Sindy 2010/11/18
Dim rsTmp As New ADODB.Recordset, bolErr2147467259 As Boolean 'Add By Sindy 2011/12/20
   
   On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   bPrintAgain = False
PrintAgain3:
   bPrint = False
   'Set Printer = Printers(Combo1.ListIndex)
   'Printer.EndDoc
   Printer.Orientation = 2 '1.直印 2.橫印
   For iRow = 1 To MSHFlexGrid1.Rows - 1
      '本所
      If MSHFlexGrid1.TextMatrix(iRow, 13) = "Y" Then
         If MSHFlexGrid1.TextMatrix(iRow, 4) <> MSHFlexGrid1.TextMatrix(iRow, 21) Then
            '檢查申請日是否相同, 不同者列印清單
            For i = 1 To 6
               strTemp(i) = ""
            Next i
            strTemp(1) = MSHFlexGrid1.TextMatrix(iRow, 14) & "-" & MSHFlexGrid1.TextMatrix(iRow, 15) & "-" & MSHFlexGrid1.TextMatrix(iRow, 16) & "-" & MSHFlexGrid1.TextMatrix(iRow, 17)
            strTemp(2) = MSHFlexGrid1.TextMatrix(iRow, 1)
            strTemp(3) = MSHFlexGrid1.TextMatrix(iRow, 2)
            strTemp(4) = MSHFlexGrid1.TextMatrix(iRow, 3)
            strTemp(5) = ChangeWStringToWDateString(MSHFlexGrid1.TextMatrix(iRow, 4))
            If iLine > 37 Or bPrint = False Then
               If bPrint <> False Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            PrintDetail
            bPrint = True
         Else
            If bPrintAgain = False Then '重新列印清單時, 資料異動不可再重覆執行
                  '檢查是否已有(1102核准通知)或(1403改變原處分並且實際結果為核准), 若有, 不出定稿不異動資料, 計算筆數
                  strSql = "select * " & _
                              "From caseprogress " & _
                              "where cp01='" & MSHFlexGrid1.TextMatrix(iRow, 14) & "' " & _
                              "and cp02='" & MSHFlexGrid1.TextMatrix(iRow, 15) & "' " & _
                              "and cp03='" & MSHFlexGrid1.TextMatrix(iRow, 16) & "' " & _
                              "and cp04='" & MSHFlexGrid1.TextMatrix(iRow, 17) & "' " & _
                              "and (cp10='1102' or (cp10='1403' and cp24='1')) "
                  intI = 1
                  Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     p_Recs1_1102 = p_Recs1_1102 + 1
                  Else
                     frm02010401_4.m_TM01 = MSHFlexGrid1.TextMatrix(iRow, 14)
                     frm02010401_4.m_TM02 = MSHFlexGrid1.TextMatrix(iRow, 15)
                     frm02010401_4.m_TM03 = MSHFlexGrid1.TextMatrix(iRow, 16)
                     frm02010401_4.m_TM04 = MSHFlexGrid1.TextMatrix(iRow, 17)
                     frm02010401_4.m_CP09 = MSHFlexGrid1.TextMatrix(iRow, 18)
                     frm02010401_4.m_CP10 = MSHFlexGrid1.TextMatrix(iRow, 19)
                     frm02010401_4.m_TM10 = MSHFlexGrid1.TextMatrix(iRow, 20)
                     frm02010401_4.m_TM11 = MSHFlexGrid1.TextMatrix(iRow, 21)
                     frm02010401_3.textResult = "1" '核准
                     frm02010401_4.textCP25 = strSrvDate(2) '核准通知日
                     frm02010401_4.textTM15 = MSHFlexGrid1.TextMatrix(iRow, 1) '審定號
                     frm02010401_4.textTM14 = Text1(0) '公告日
                     frm02010401_4.textTMBM07_2 = Text1(1) '期數
                     frm02010401_4.m_CP05 = strSrvDate(2) '來函收文日
                     frm02010401_4.m_TM23 = MSHFlexGrid1.TextMatrix(iRow, 22)
                     frm02010401_4.m_TM20 = MSHFlexGrid1.TextMatrix(iRow, 23)
                     frm02010401_4.textPrint = MSHFlexGrid1.TextMatrix(iRow, 24)
                     '帶列印定稿預設值
                     If frm02010401_4.textPrint = "" Then
                        frm02010401_4.textPrint = GetTWordLng(MSHFlexGrid1.TextMatrix(iRow, 14), MSHFlexGrid1.TextMatrix(iRow, 15), MSHFlexGrid1.TextMatrix(iRow, 16), MSHFlexGrid1.TextMatrix(iRow, 17))
                     End If
                     frm02010401_4.m_CP14 = strUserNum
                     
                     '檢查是否已有1001.核准, 若有, 代表重覆執行不異動資料出定稿
                     strSql = "select * from caseprogress " & _
                                  "where cp09 in (select cp43 From caseprogress " & _
                                  "where cp01='" & MSHFlexGrid1.TextMatrix(iRow, 14) & "' " & _
                                  "and cp02='" & MSHFlexGrid1.TextMatrix(iRow, 15) & "' " & _
                                  "and cp03='" & MSHFlexGrid1.TextMatrix(iRow, 16) & "' " & _
                                  "and cp04='" & MSHFlexGrid1.TextMatrix(iRow, 17) & "' " & _
                                  "and cp10='1001' " & _
                                  "and cp43 is not null) " & _
                                  "and cp10='101' "
                     intI = 1
                     Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
                     If intI <> 1 Then
                        If frm02010401_4.OnSaveData = False Then
                           MsgBox MSHFlexGrid1.TextMatrix(iRow, 1) & "存檔失敗，請洽系統管理員 !", vbCritical
                           Screen.MousePointer = vbDefault
                           Exit Function
                        End If
                     'Add By Sindy 2022/1/11 抓進度已存在的(核准-申請)總收文號,後面Run定稿會使用到
                     Else
                        strExc(10) = adoRecordset.Fields("cp09") '申請總收文號
                        adoRecordset.Close
                        strSql = "select * from caseprogress " & _
                                     "where cp01='" & MSHFlexGrid1.TextMatrix(iRow, 14) & "' " & _
                                     "and cp02='" & MSHFlexGrid1.TextMatrix(iRow, 15) & "' " & _
                                     "and cp03='" & MSHFlexGrid1.TextMatrix(iRow, 16) & "' " & _
                                     "and cp04='" & MSHFlexGrid1.TextMatrix(iRow, 17) & "' " & _
                                     "and cp10='1001' " & _
                                     "and cp43='" & strExc(10) & "' "
                        intI = 1
                        Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
                        If intI > 0 Then
                           frm02010401_4.strLD18 = adoRecordset.Fields("cp09") '信函總收文號
                        End If
                     '2022/1/11 END
                     End If
                     adoRecordset.Close
                     '列印定稿
                     frm02010401_4.PrintLetter
                     Unload frm02010401_3
                     Unload frm02010401_4
                     'Add By Sindy 2012/8/20
                     Dim strSales As String, strSales_cc As String
                     Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
                    'modify by sonia 2018/9/13 +申請人名稱cu04
                     strSql = "select cp01,cp02,cp03,cp04,tm12,tm15,tm05,tm09,tm14,DECODE(CU15,'0','台端','1','貴公司','貴單位') as cu15Nm," & _
                              "nvl(nvl(cu04,cu05||decode(cu88,null,null,' '||cu88)||decode(cu89,null,null,' '||cu89)||decode(cu90,null,null,' '||cu90)),cu06) cu04 from caseprogress,trademark,customer " & _
                                  "where cp09 in (select cp43 From caseprogress " & _
                                  "where cp01='" & MSHFlexGrid1.TextMatrix(iRow, 14) & "' " & _
                                  "and cp02='" & MSHFlexGrid1.TextMatrix(iRow, 15) & "' " & _
                                  "and cp03='" & MSHFlexGrid1.TextMatrix(iRow, 16) & "' " & _
                                  "and cp04='" & MSHFlexGrid1.TextMatrix(iRow, 17) & "' " & _
                                  "and cp10='1001' " & _
                                  "and cp43 is not null) " & _
                                  "and cp10='101' " & _
                                  "and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
                                  "and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
                     intI = 1: strSales = "": strSales_cc = ""
                     Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        m_CP01 = adoRecordset.Fields("cp01")
                        m_CP02 = adoRecordset.Fields("cp02")
                        m_CP03 = adoRecordset.Fields("cp03")
                        m_CP04 = adoRecordset.Fields("cp04")
                        '讀取智權人員
                        strSales = PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)
                        '若為68096中三杜副總時，檢查客戶若最後收文為在職人員則設為副本收件人
                        If strSales = "68096" Then
                           strExc(0) = "select st01 from staff,(select max(cp05||cp13) cp13 from ( " & _
                              "      Select cp05,cp13 From patent, caseprogress Where pa26='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and cp09<'B' " & _
                              "union Select cp05,cp13 From trademark, caseprogress Where tm23='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and cp09<'B' " & _
                              "union Select cp05,cp13 From lawcase, caseprogress Where lc11='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and cp09<'B' " & _
                              "union Select cp05,cp13 From servicepractice, caseprogress Where sp08='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and cp09<'B' " & _
                              "union Select cp05,cp13 From hirecase, caseprogress Where hc05='" & GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) & "' and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp09<'B' " & _
                              ")) aa where substr(aa.cp13,9)=st01(+) and st04='1'"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              strSales_cc = RsTemp.Fields(0).Value
                           End If
                        End If
                        '寄發智權同仁由同仁轉客戶
                        'modify by sonia 2018/9/13 內文加申請人
                        'PUB_SendMail strUserNum, strSales, "", "大陸商標核准通知（本所案號：" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "）", _
                        '"敬啟者：" & vbCrLf & vbCrLf & _
                        '"　　" & adoRecordset.Fields("cu15Nm") & "委託本所辦理之第" & IIf(adoRecordset.Fields("tm15") = "", adoRecordset.Fields("tm12"), adoRecordset.Fields("tm15")) & "號「" & adoRecordset.Fields("tm05") & "」（第" & adoRecordset.Fields("tm09") & "類）大陸商標註冊申請案，業經審查核准，公告於" & Left(adoRecordset.Fields("tm14"), 4) - 1911 & "年" & Mid(adoRecordset.Fields("tm14"), 5, 2) & "月" & Right(adoRecordset.Fields("tm14"), 2) & "日之大陸商標公報，公告資料將另函郵寄予　" & adoRecordset.Fields("cu15Nm") & "，請留意查收。" & vbCrLf & vbCrLf & _
                        '"台一國際專利商標事務所　敬上", "", , , , , strSales_cc, , , , , False
                        'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
                        PUB_SendMail strUserNum, strSales, "", "大陸商標核准通知（本所案號：" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "）", _
                        "致　" & adoRecordset.Fields("cu04") & "：" & vbCrLf & vbCrLf & _
                        "　　" & adoRecordset.Fields("cu15Nm") & "委託本所辦理之第" & IIf(adoRecordset.Fields("tm15") = "", adoRecordset.Fields("tm12"), adoRecordset.Fields("tm15")) & "號「" & adoRecordset.Fields("tm05") & "」（第" & adoRecordset.Fields("tm09") & "類）大陸商標註冊申請案，業經審查核准，公告於" & Left(adoRecordset.Fields("tm14"), 4) - 1911 & "年" & Mid(adoRecordset.Fields("tm14"), 5, 2) & "月" & Right(adoRecordset.Fields("tm14"), 2) & "日之大陸商標公報，公告資料將另函郵寄予　" & adoRecordset.Fields("cu15Nm") & "，請留意查收。" & vbCrLf & vbCrLf & _
                        PUB_GetCompName2("1") & "　敬上", "", , , , , strSales_cc, , , , , False
                        'end 2018/9/13
                     End If
                     adoRecordset.Close
                     '2012/8/20 End
                  End If
            End If
         End If
      End If
   Next
   If bPrint = True Then
      If bPrintAgain = False Then MsgBox "將列印審定公告資料檢核清單，請換一般列印紙!!!"
      Printer.EndDoc
      If MsgBox("列印審定公告資料檢核清單，是否已列印成功？成功按「是」，不成功按「否」再列印一次！", vbYesNo + vbDefaultButton2) = vbNo Then
         bPrintAgain = True
         GoTo PrintAgain3
      End If
   End If
   Screen.MousePointer = vbDefault
   
'   '產生Word檔
'   bolRetry = True
'   Screen.MousePointer = vbHourglass
'   For iRow = 1 To MSHFlexGrid1.Rows - 1
'      '非本所
'      If MSHFlexGrid1.TextMatrix(iRow, 13) <> "Y" Then
'         m_AppAddr = MSHFlexGrid1.TextMatrix(iRow, 9) '申請人地址
'         m_AppName = MSHFlexGrid1.TextMatrix(iRow, 3) '商標註冊人
'         'Add By Sindy 2010/11/18
'         strTempAppAddr = "": strCU01 = "": m_AppAddrZip = ""
'         'Modify By Sindy 2012/6/19 註冊號9727839申請人為吳志忠此案件之故,再增加名稱若小於等於4個字的也是個人
'         If Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) <> "" Or Len(Trim(MSHFlexGrid1.TextMatrix(iRow, 3))) <= 4 Then
'            '個人, 抓名稱及ID都相同者,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
''            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' and cu11='" & Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) & "' order by cu01 asc,cu02 asc "
''            intI = 1
''            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
''            If intI = 1 Then
''               If RsTemp.Fields("cu02") <> "0" Then
''                  strCU01 = "" & RsTemp.Fields("cu01")
''               Else
''                  m_AppAddrZip = Trim("" & RsTemp.Fields("cu30"))
''                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
''               End If
''            End If
'            'Modify By Sindy 2011/12/20
'            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' and cu11='" & Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) & "' order by cu01 asc,cu02 asc "
'            rsTmp.CursorLocation = adUseClient
'            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            If bolErr2147467259 = True Then
'               bolErr2147467259 = False
''               rsTmp.Close
'               strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & " ' and cu11='" & Trim(MSHFlexGrid1.TextMatrix(iRow, 25)) & "' order by cu01 asc,cu02 asc "
'               rsTmp.CursorLocation = adUseClient
'               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            End If
'            If rsTmp.RecordCount > 0 Then
'               If rsTmp.Fields("cu02") <> "0" Then
'                  strCU01 = "" & rsTmp.Fields("cu01")
'               Else
'                  m_AppAddrZip = Trim("" & rsTmp.Fields("cu30"))
'                  strTempAppAddr = Trim("" & rsTmp.Fields("cu31"))
'               End If
'            End If
'            rsTmp.Close
'         Else
'            '非個人, 抓相同客戶名稱之最小客戶編號,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
''            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc "
''            intI = 1
''            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
''            If intI = 1 Then
''               If RsTemp.Fields("cu02") <> "0" Then
''                  strCU01 = "" & RsTemp.Fields("cu01")
''               Else
''                  m_AppAddrZip = Trim("" & RsTemp.Fields("cu30"))
''                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
''               End If
''            End If
'            'Modify By Sindy 2011/12/20
'            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc "
'            rsTmp.CursorLocation = adUseClient
'            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            If bolErr2147467259 = True Then
'               bolErr2147467259 = False
''               rsTmp.Close
'               strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & " ' order by cu01 asc,cu02 asc "
'               rsTmp.CursorLocation = adUseClient
'               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            End If
'            If rsTmp.RecordCount > 0 Then
'               If rsTmp.Fields("cu02") <> "0" Then
'                  strCU01 = "" & rsTmp.Fields("cu01")
'               Else
'                  m_AppAddrZip = Trim("" & rsTmp.Fields("cu30"))
'                  strTempAppAddr = Trim("" & rsTmp.Fields("cu31"))
'               End If
'            End If
'            rsTmp.Close
'         End If
'         If strCU01 <> "" Then
'            strSql = "SELECT * FROM customer WHERE cu01='" & strCU01 & "' and cu02='0' "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               m_AppAddrZip = Trim("" & RsTemp.Fields("cu30"))
'               strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
'            End If
'         End If
'         If strTempAppAddr <> "" Then
'            m_AppAddr = strTempAppAddr
'         Else
'            m_AppAddrZip = ""
'         End If
'         '2010/11/18 End
'         m_AppDate = MSHFlexGrid1.TextMatrix(iRow, 4) '申請日期
'         m_AppDate = Val(Left(m_AppDate, 4)) & "年" & Mid(m_AppDate, 5, 2) & "月" & Right(m_AppDate, 2) & "日"
'         m_AppNum = MSHFlexGrid1.TextMatrix(iRow, 1) '註冊號
'         m_TName = MSHFlexGrid1.TextMatrix(iRow, 2) '商標名稱
'         m_Goods = MSHFlexGrid1.TextMatrix(iRow, 7) '分類
'         m_DATE = Val(text1(0)) + 19110000 '公告日期
'         m_DATE = Val(Left(m_DATE, 4)) & "年" & Mid(m_DATE, 5, 2) & "月" & Right(m_DATE, 2) & "日"
'         m_Num = text1(1) '期數
'         ' 列印定稿
'         Call WordEdit(1)
'      End If
'   Next
'   If bolRetry = False Then
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
'   End If
'   Screen.MousePointer = vbDefault
   
   Process3 = True
'   Set g_WordAp = Nothing
'   Set rsTmp = Nothing 'Add By Sindy 2011/12/20
   Exit Function
   
ErrHnd:
'   'Add By Sindy 2011/12/20
'   bolErr2147467259 = False
'   If Err.Number = -2147467259 Then
'      bolErr2147467259 = True
'      '接著發生錯誤陳述式的下個陳述式開始執行
'      Resume Next
'   End If
'   Set g_WordAp = Nothing
'   Set rsTmp = Nothing
'   '2011/12/20 End
   If Err.Number <> 0 Then
      'cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub WordEdit(strKind As String)
   'Add by Morgan 2011/10/26 +信頭
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
   
   'end 2011/10/26
   
On Error GoTo ERRORSECTION1
   
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   With g_WordAp
   
      If bolRetry = True Then
         g_WordAp.Documents.add
         'Add by Morgan 2011/10/26 +信頭
         If PUB_ReadDB2File(stFileName, iPicNo) = True Then
            '切換為整頁模式
            If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
               .ActiveWindow.ActivePane.View.Type = wdPageView
            Else
               .ActiveWindow.View.Type = wdPageView
            End If
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
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
               .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(27)
            End If
            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
            .Selection.EndKey Unit:=wdStory
         End If
         'end 2011/10/26
      End If
   
      If bolRetry = False Then .Selection.InsertBreak Type:=wdPageBreak
      
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      'Modify by Morgan 2008/7/3
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
      'end 2008/7/3
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      'Add by Morgan 2008/7/17 配合新的開窗定稿改固定行高
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
      .Selection.ParagraphFormat.LineSpacing = 15
      'end 2008/7/17
      

      
      .Selection.TypeParagraph 'Add by Morgan 2008/6/11 CFT 信頭比較高
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
'      '設定字型版面(參照定稿)
'      '.Selection.Font.Name = "Times New Roman"
'      .Selection.Font.Name = "標楷體"
'      .Selection.PageSetup.Orientation = wdOrientPortrait
'      .Selection.Orientation = wdTextOrientationHorizontal
'      .Selection.Font.Size = 14
'      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3.175)
'      .Selection.PageSetup.RightMargin = .CentimetersToPoints(3.175)
'      .Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
'      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
'      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      '靠左
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      '置右
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      '置中
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      '不要分散對齊
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
      
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
      '審定公告
      If strKind = "1" Then
         .Selection.TypeText "　　首先恭禧　貴公司／台端於" & m_AppDate & "申請之第" & m_AppNum & "號「" & m_TName & "」(第" & m_Goods & "類)之大陸商標註冊申請案，已獲核准審定，公告於" & m_DATE & "之" & m_Num & "期大陸商標公報，公告三個月期間，若無人提出異議，本件商標即可取得註冊，特此通知。"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "　　台一從事專利及商標代理、保護業務超過三十年，對兩岸及世界各國智慧財產權都能提供最專業的服務。針對　貴公司／台端的相關業務本所可以指派專人注意，若貴公司／台端有任何關於智慧財產權方面的問題及需要請電洽"
         .Selection.Font.Bold = True
         .Selection.TypeText "本所商標處"
         .Selection.Font.Bold = False
         .Selection.TypeText "(聯絡電話：02-25061023轉" & Text1(2) & ")，本所專業人員將協助安排本所最具專業切合貴公司／台端需要的專業人員提供服務。"
      '通知續展
      ElseIf strKind = "2" Then
         .Selection.TypeText "　　台端／貴公司所有註冊第" & m_AppNum & "號「" & m_TName & "」(第" & m_Goods & "類)大陸商標之商標權期限"
         .Selection.Font.Bold = True
         .Selection.TypeText "即將於" & m_DATEe & "屆滿"
         .Selection.Font.Bold = False
         .Selection.TypeText "，　台端／貴公司若有意繼續使用此商標，應儘速辦理商標延展註冊。"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "商標延展註冊所需之資料如下："
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "　一、商標註冊證影本：(商標權期間：" & m_DATEs & "－" & m_DATEe & ")；"
         .Selection.TypeParagraph
         .Selection.TypeText "　二、委託書。"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "　　商標權期限屆滿，未辦理延展註冊者，依大陸商標法規定，可於期滿後6個月內之寬展期內申請延展註冊，惟須再繳納延遲費用。若於寬展期滿仍未提出延展申請者，則商標權失效。"
         .Selection.TypeParagraph
         .Selection.TypeText "　　為維護　台端／貴公司之商標權益，特函提醒　台端／貴公司。請儘速與本所服務人員聯繫，本所竭誠為　台端／貴公司提供服務！"
      'Add By Sindy 2013/11/25
      '通知續展:同申請人多案一定稿
      ElseIf strKind = "3" Then
         .Selection.TypeText "　　台端／貴公司所有註冊" & tmp_TName & "大陸商標之商標權期限"
         .Selection.Font.Bold = True
         .Selection.TypeText "即將於" & m_DATEe & "屆滿"
         .Selection.Font.Bold = False
         .Selection.TypeText "，　台端／貴公司若有意繼續使用此商標，應儘速辦理商標延展註冊。"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "商標延展註冊所需之資料如下："
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "　一、商標註冊證影本。"
         .Selection.TypeParagraph
         .Selection.TypeText "　二、委託書。"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "　　商標權期限屆滿，未辦理延展註冊者，依大陸商標法規定，可於期滿後6個月內之寬展期內申請延展註冊，惟須再繳納延遲費用。若於寬展期滿仍未提出延展申請者，則商標權失效。"
         .Selection.TypeParagraph
         .Selection.TypeText "　　為維護　台端／貴公司之商標權益，特函提醒　台端／貴公司。請儘速與本所服務人員聯繫，本所竭誠為　台端／貴公司提供服務！"
      '2013/11/25 END
'      'Add By Sindy 2014/2/19
'      '審定公告(核准公告定稿):同申請人多案一定稿
'      ElseIf strKind = "4" Then
'         .Selection.TypeText "　　首先恭禧　貴公司／台端於" & tmp_TName & "之大陸商標註冊申請案，已獲核准審定，公告於" & m_DATE & "之" & m_Num & "期大陸商標公報，公告三個月期間，若無人提出異議，本件商標即可取得註冊，特此通知。"
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'         .Selection.TypeText "　　台一從事專利及商標代理、保護業務超過三十年，對兩岸及世界各國智慧財產權都能提供最專業的服務。針對　貴公司／台端的相關業務本所可以指派專人注意，若貴公司／台端有任何關於智慧財產權方面的問題及需要請電洽"
'         .Selection.Font.Bold = True
'         .Selection.TypeText "本所商標處"
'         .Selection.Font.Bold = False
'         .Selection.TypeText "(聯絡電話：02-25061023轉" & Text1(2) & ")，本所專業人員將協助安排本所最具專業切合貴公司／台端需要的專業人員提供服務。"
'      '2014/2/19 END
      End If
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "　　　　耑　此　　敬　頌"
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "商　祺"
      .Selection.TypeParagraph
      'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
      '.Selection.TypeText "　　　　　　　　　　　　　　　台一國際專利商標事務所  敬上"
      .Selection.TypeText "　　　　　　　　　　　　　　　" & PUB_GetCompName2("1") & "  敬上"
      'end 2020/3/30
      .Selection.TypeParagraph
'      If strKind = "1" Then
'         .Selection.TypeText "　　　　　　　　　　　　　　　　　　" & Val(Left(strSrvDate(1), 4)) - 1911 & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Right(strSrvDate(1), 2) & "日"
'      ElseIf strKind = "2" Then
         .Selection.TypeText "　　　　　　　　　　　　　　　服務人員：" & Label8.Caption
         .Selection.TypeParagraph
         .Selection.TypeText "　　　　　　　　　　　　　　　服務專線：02-25061023" & IIf(Text1(2).Text <> "", " 轉 " & Text1(2), "")
'      End If
      .Selection.TypeParagraph
      
'      .Selection.WholeStory
'      ChgWordFormat g_WordAp, .Selection.Text
   End With
   
'   PhaseIndent    '調整首行凸排
'   g_WordAp.Visible = True
'   g_WordAp.WindowState = wdWindowStateMaximize
   bolRetry = False
   
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
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
      End Select
   End If
End Sub

'調整首行凸排
Sub PhaseIndent()
    g_WordAp.Selection.WholeStory
    With g_WordAp.Selection.ParagraphFormat
        .LeftIndent = g_WordAp.CentimetersToPoints(1)
        .RightIndent = g_WordAp.CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 15
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = g_WordAp.CentimetersToPoints(-1)
        .OutlineLevel = wdOutlineLevelBodyText
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = True
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub

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

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         If Text1(2).Text = "" Then
            MsgBox "分機不可空白！", vbExclamation
            Text1(2).SetFocus
            Exit Sub
         End If
         '檢查資料
         If Option1(0).Value = True Then
            If txtPath(0).Text = "" Then
               MsgBox "檔案路徑不可空白！", vbExclamation
               txtPath(0).SetFocus
               Exit Sub
            End If
            If Val(Text1(0).Text) = 0 Then
               MsgBox "公告日不可空白！", vbExclamation
               Text1(0).SetFocus
               Exit Sub
            End If
            If Val(Text1(1).Text) = 0 Then
               MsgBox "期數不可空白！", vbExclamation
               Text1(1).SetFocus
               Exit Sub
            End If
         ElseIf Option1(1).Value = True Then
            If txtPath(1).Text = "" Then
               MsgBox "檔案路徑不可空白！", vbExclamation
               txtPath(1).SetFocus
               Exit Sub
            End If
         End If
         If LoadXLS() = True Then
            If Option1(0).Value = True Then
               'Add By Sindy 2012/9/24
               '全部資料(第一次)
               If Option2(0).Value = True Then
                  If p_Recs1 = 0 Then
                     MsgBox "審定公告檔案內無資料！", vbExclamation
                     Exit Sub
                  Else
                     Process3
                     'If MsgBox("審定公告作業完成！匯入檔案 " & p_Recs1 & " 筆、本所案件數" & p_Recs1_Y & " 筆、非本所案件數" & p_Recs1_N & " 筆。列印此數據，列印後交葉經理!!", vbYesNo + vbDefaultButton2) = vbYes Then
                     'MsgBox "審定公告作業完成！匯入檔案 " & p_Recs1 & " 筆、本所案件數" & p_Recs1_Y & " 筆、非本所案件數" & p_Recs1_N & " 筆、代理人已通知件數" & p_Recs1_1102 & " 筆。列印此數據後交葉經理!!"
                     MsgBox "審定公告作業完成！匯入檔案 " & p_Recs1 & " 筆、本所案件數" & p_Recs1_Y & " 筆、代理人已通知件數" & p_Recs1_1102 & " 筆。列印此數據後交葉經理!!"
                        'Printer.EndDoc
PrintAgain5:
                        Printer.Orientation = 1 '1.直印 2.橫印
                        PLeft(1) = 1000
                        Printer.Font.Size = 12
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 2 * 300
                        Printer.Print "檔案路徑：" & txtPath(0)
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 3 * 300
                        Printer.Print "公告日：" & ChangeTStringToTDateString(Text1(0))
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 4 * 300
                        Printer.Print "期　數：" & Text1(1)
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 5 * 300
                        Printer.Print "審定公告匯入結果："
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 6 * 300
                        Printer.Print "　　　　　匯入檔案 " & Right("            " & p_Recs1, 12) & " 筆"
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 7 * 300
                        Printer.Print "　　　　本所案件數 " & Right("            " & p_Recs1_Y, 12) & " 筆"
'                        Printer.CurrentX = PLeft(1)
'                        Printer.CurrentY = 8 * 300
'                        Printer.Print "　　　非本所案件數 " & Right("            " & p_Recs1_N, 12) & " 筆"
                        Printer.CurrentX = PLeft(1)
                        'Printer.CurrentY = 9 * 300
                        Printer.CurrentY = 8 * 300
                        Printer.Print "　代理人已通知件數 " & Right("            " & p_Recs1_1102, 12) & " 筆"
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 11 * 300
                        Printer.Print "列印人員：" & strUserName
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 12 * 300
                        Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
                        'Add By Sindy 2013/1/2
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 14 * 300
                        Printer.Print "為本所案件有 " & p_Recs1_Y & " 筆，如下：" & vbCrLf & m_CaseNo
                        '2013/1/2 End
                        Printer.EndDoc
                        If MsgBox("列印審定公告匯入結果，是否已列印成功？成功按「是」，不成功按「否」再列印一次！", vbYesNo + vbDefaultButton2) = vbNo Then
                           GoTo PrintAgain5
                        End If
                     'End If
                     Exit Sub
                  End If
               '2012/9/24 End
               ElseIf Option2(1).Value = True Then '台灣申請人案件 (第二次)
                  If p_Recs1 = 0 Then
                     MsgBox "審定公告檔案內無資料！", vbExclamation
                     Exit Sub
                  Else
                     Process1
                     'If MsgBox("審定公告作業完成！匯入檔案 " & p_Recs1 & " 筆、本所案件數" & p_Recs1_Y & " 筆、非本所案件數" & p_Recs1_N & " 筆。列印此數據，列印後交葉經理!!", vbYesNo + vbDefaultButton2) = vbYes Then
                     'MsgBox "審定公告作業完成！匯入檔案 " & p_Recs1 & " 筆、本所案件數" & p_Recs1_Y & " 筆、非本所案件數" & p_Recs1_N & " 筆、代理人已通知件數" & p_Recs1_1102 & " 筆。列印此數據後交葉經理!!"
                     MsgBox "審定公告作業完成！匯入檔案 " & p_Recs1 & " 筆、非本所案件數" & p_Recs1_N & " 筆。列印此數據後交葉經理!!"
                        'Printer.EndDoc
PrintAgain1:
                        Printer.Orientation = 1 '1.直印 2.橫印
                        PLeft(1) = 1000
                        Printer.Font.Size = 12
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 2 * 300
                        Printer.Print "檔案路徑：" & txtPath(0)
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 3 * 300
                        Printer.Print "公告日：" & ChangeTStringToTDateString(Text1(0))
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 4 * 300
                        Printer.Print "期　數：" & Text1(1)
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 5 * 300
                        Printer.Print "台灣申請人：" '"審定公告匯入結果：" Modify By Sindy 2017/5/18
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 6 * 300
                        Printer.Print "　　　　　匯入檔案 " & Right("            " & p_Recs1, 12) & " 筆"
'                        Printer.CurrentX = PLeft(1)
'                        Printer.CurrentY = 7 * 300
'                        Printer.Print "　　　　本所案件數 " & Right("            " & p_Recs1_Y, 12) & " 筆"
                        Printer.CurrentX = PLeft(1)
                        'Printer.CurrentY = 8 * 300
                        Printer.CurrentY = 7 * 300
                        Printer.Print "　　　非本所案件數 " & Right("            " & p_Recs1_N, 12) & " 筆"
'                        Printer.CurrentX = PLeft(1)
'                        Printer.CurrentY = 9 * 300
'                        Printer.Print "　代理人已通知件數 " & Right("            " & p_Recs1_1102, 12) & " 筆"
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 11 * 300
                        Printer.Print "列印人員：" & strUserName
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 12 * 300
                        Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
                        'Add By Sindy 2012/10/15
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = 14 * 300
                        Printer.Print "為本所案件有 " & p_Recs1_Y & " 筆，如下：" & vbCrLf & m_CaseNo
                        '2012/10/15 End
                        Printer.EndDoc
                        If MsgBox("列印審定公告匯入結果，是否已列印成功？成功按「是」，不成功按「否」再列印一次！", vbYesNo + vbDefaultButton2) = vbNo Then
                           GoTo PrintAgain1
                        End If
                     'End If
                     Exit Sub
                  End If
               End If
            ElseIf Option1(1).Value = True Then
               If p_Recs2 = 0 Then
                  MsgBox "通知續展檔案內無資料！", vbExclamation
                  Exit Sub
               Else
                  Process2
                  'If MsgBox("通知續展作業完成！匯入檔案 " & p_Recs2 & " 筆、本所案件數" & p_Recs2_Y & " 筆、已收文" & p_Recs2_Recv & " 筆、已結案" & p_Recs2_Close & " 筆、本所定稿" & p_Recs2_Print & " 筆、非本所案件數" & p_Recs2_N & " 筆。列印此數據，列印後交葉經理!!", vbYesNo + vbDefaultButton2) = vbYes Then
                  MsgBox "通知續展作業完成！匯入檔案 " & p_Recs2 & " 筆、本所案件數" & p_Recs2_Y & " 筆、已收文" & p_Recs2_Recv & " 筆、已結案" & p_Recs2_Close & " 筆、本所定稿" & p_Recs2_Print & " 筆、非本所案件數" & p_Recs2_N & " 筆。列印此數據後交葉經理!!"
                     'Printer.EndDoc
PrintAgain2:
                     Printer.Orientation = 1 '1.直印 2.橫印
                     PLeft(1) = 1000
                     Printer.Font.Size = 12
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 2 * 300
                     Printer.Print "檔案路徑：" & txtPath(1)
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 3 * 300
                     Printer.Print "通知續展匯入結果："
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 4 * 300
                     Printer.Print "　　　匯入檔案 " & Right("            " & p_Recs2, 12) & " 筆"
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 5 * 300
                     Printer.Print "　　本所案件數 " & Right("            " & p_Recs2_Y, 12) & " 筆"
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 6 * 300
                     Printer.Print "　　　　已收文 " & Right("            " & p_Recs2_Recv, 12) & " 筆"
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 7 * 300
                     Printer.Print "　　　　已結案 " & Right("            " & p_Recs2_Close, 12) & " 筆"
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 8 * 300
                     Printer.Print "　　　本所定稿 " & Right("            " & p_Recs2_Print, 12) & " 筆"
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 9 * 300
                     Printer.Print "　非本所案件數 " & Right("            " & p_Recs2_N, 12) & " 筆"
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 11 * 300
                     Printer.Print "列印人員：" & strUserName
                     Printer.CurrentX = PLeft(1)
                     Printer.CurrentY = 12 * 300
                     Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
                     Printer.EndDoc
                     If MsgBox("列印通知續展匯入結果，是否已列印成功？成功按「是」，不成功按「否」再列印一次！", vbYesNo + vbDefaultButton2) = vbNo Then
                        GoTo PrintAgain2
                     End If
                  'End If
                  Exit Sub
               End If
            End If
            MsgBox "作業完成！"
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
Dim stFileName As String
Dim strFile As String 'Add By Sindy 2012/9/28
   
On Error GoTo ErrHnd

   If Option2(0).Value = True Then
      strFile = "xls"
   Else
      strFile = "txt"
   End If
   
   stFileName = "*." & strFile
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      If Option2(0).Value = True Then
         .Filter = "Excel檔案 (*." & strFile & ")|*." & strFile & ""
      Else
         .Filter = "文字檔案 (*." & strFile & ")|*." & strFile & ""
      End If
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtPath(0).Text = .FileName
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
         txtPath(1).Text = .FileName
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
   Option1(0).Value = True
   Option2(0).Value = True 'Add By Sindy 2012/9/24
   Label8.Caption = strUserName
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020320 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   CloseIme
   TextInverse Text1(Index)
End Sub

Private Sub Option1_Click(Index As Integer)
   If Option1(0).Value = True Then
      txtPath(0).Enabled = True
      Text1(0).Enabled = True
      Text1(1).Enabled = True
      Command1.Enabled = True
      txtPath(1).Enabled = False
      Command2.Enabled = False
   ElseIf Option1(1).Value = True Then
      txtPath(0).Enabled = False
      Text1(0).Enabled = False
      Text1(1).Enabled = False
      Command1.Enabled = False
      txtPath(1).Enabled = True
      Command2.Enabled = True
   End If
End Sub
'2011/3/25 add by sonia
' 公告日
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Select Case Index
      Case 0
         If IsEmptyText(Text1(Index)) = False Then
            ' 檢查是否為民國年
            If CheckIsTaiwanDate(Text1(Index), False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的公告日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               Text1_GotFocus Index
            End If
         End If
   End Select
End Sub

Private Sub txtPath_GotFocus(Index As Integer)
   TextInverse txtPath(Index)
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2300
PLeft(3) = 3500
PLeft(4) = 6500
PLeft(5) = 10500
PLeft(6) = 12000
End Sub

Sub PrintTitle()
Dim strItem As String

GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

If Option1(0).Value = True Then
   strItem = "審定公告資料檢核(申請日不同者)"
ElseIf Option1(1).Value = True Then
   strItem = "通知續展資料檢核(申請日或專用期止日不同者)"
End If
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strItem) / 2)
Printer.CurrentY = iLine * 300
Printer.Print strItem

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = 4850
Printer.CurrentY = 900
If Option1(0).Value = True Then
   Printer.Print "檔案路徑：" & txtPath(0)
Else
   Printer.Print "檔案路徑：" & txtPath(1)
End If
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
If Option1(0).Value = True Then
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = 1200
   Printer.Print "公  告  日：" & ChangeTStringToTDateString(Text1(0))
   Printer.CurrentX = 4850
   Printer.CurrentY = 1200
   Printer.Print "期　　數：" & Text1(1)
End If
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iLine = 6
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "註冊號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "商標名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "商標註冊人"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "申請日期"
If Option1(1).Value = True Then
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iLine * 300
   Printer.Print "專用期(止)"
End If

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(205, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
   For m_j = 1 To 6
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub

'通知續展
Private Function Process2() As Boolean
Dim iRow As Integer, iRecs As Integer, iXRow As Integer
Dim bPrint As Boolean, i As Integer, bPrintAgain As Boolean
Dim strTempAppAddr As String, strCU01 As String 'Add By Sindy 2010/11/18
Dim rsTmp As New ADODB.Recordset, bolErr2147467259 As Boolean 'Add By Sindy 2011/12/20
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim boleFileSave As Boolean, m_TM01 As String
'2012/1/13 End
Dim strCP27 As String, strCP64 As String 'Add By Sindy 2013/8/15
   
   On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   bPrintAgain = False
   cnnConnection.BeginTrans
PrintAgain4:
   bPrint = False
   'Set Printer = Printers(Combo1.ListIndex)
   'Printer.EndDoc
   Printer.Orientation = 2 '1.直印 2.橫印
   boleFileSave = False 'Add By Sindy 2012/1/13
   For iRow = 1 To MSHFlexGrid2.Rows - 1
'      If MSHFlexGrid2.TextMatrix(iRow, 1) = "620048" Then
'         MsgBox MSHFlexGrid2.TextMatrix(iRow, 1)
'      End If
      '本所
      If MSHFlexGrid2.TextMatrix(iRow, 10) = "Y" Then
         If (MSHFlexGrid2.TextMatrix(iRow, 4) <> MSHFlexGrid2.TextMatrix(iRow, 15)) Or _
            (MSHFlexGrid2.TextMatrix(iRow, 6) <> MSHFlexGrid2.TextMatrix(iRow, 16)) Then
            '檢查申請日或專用期止日是否相同, 不同者列印清單
            For i = 1 To 6
               strTemp(i) = ""
            Next i
            strTemp(1) = MSHFlexGrid2.TextMatrix(iRow, 11) & "-" & MSHFlexGrid2.TextMatrix(iRow, 12) & "-" & MSHFlexGrid2.TextMatrix(iRow, 13) & "-" & MSHFlexGrid2.TextMatrix(iRow, 14)
            strTemp(2) = MSHFlexGrid2.TextMatrix(iRow, 1)
            strTemp(3) = Left(Trim(MSHFlexGrid2.TextMatrix(iRow, 2)), 16)
            strTemp(4) = Left(Trim(MSHFlexGrid2.TextMatrix(iRow, 3)), 16)
            strTemp(5) = ChangeWStringToWDateString(MSHFlexGrid2.TextMatrix(iRow, 4))
            strTemp(6) = ChangeWStringToWDateString(MSHFlexGrid2.TextMatrix(iRow, 6))
            If iLine > 37 Or bPrint = False Then
               If bPrint <> False Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            PrintDetail
            bPrint = True
         End If
         If bPrintAgain = False Then '重新列印清單時, 資料異動不可再重覆執行
            '本所案件新增一筆進度檔
            ' 收文號
            strCP09 = Empty
            strCP09 = AutoNo("C", 6)
            
            'Add By Sindy 2012/1/13
            ET01 = "12"
            ET02 = strCP09
            bolEdit = False
            m_TM01 = MSHFlexGrid2.TextMatrix(iRow, 11)
            iCopy = 0
            '2012/1/13 End
            
            'Modify By Sindy 2013/12/20 本所案件全部不催延展,不出定稿
            strCP27 = "19221111"
'            'Add By Sindy 2013/8/15 若為FMT案件且為不催延展時,產生CP但發文日上11/11/11,進度備註加註不催延展,不產生定稿
'            strCP27 = strSrvDate(1)
'            strCP64 = ""
'            If Left(GetSalesArea(PUB_GetAKindSalesNo(MSHFlexGrid2.TextMatrix(iRow, 11), _
'                                                     MSHFlexGrid2.TextMatrix(iRow, 12), _
'                                                     MSHFlexGrid2.TextMatrix(iRow, 13), _
'                                                     MSHFlexGrid2.TextMatrix(iRow, 14))), 1) = "F" And _
'               PUB_ChkCaseIsNoticeScale(MSHFlexGrid2.TextMatrix(iRow, 11), _
'                                        MSHFlexGrid2.TextMatrix(iRow, 12), _
'                                        MSHFlexGrid2.TextMatrix(iRow, 13), _
'                                        MSHFlexGrid2.TextMatrix(iRow, 14)) = False Then
'               strCP27 = "19221111"
'               strCP64 = "不催延展,不產生定稿"
'            End If
'            '2013/8/15 END
'            'Add By Sindy 2013/12/20
'            If strCP64 = "" Then
'               strExc(0) = "SELECT * FROM NextProgress " & _
'                              "WHERE NP02='" & MSHFlexGrid2.TextMatrix(iRow, 11) & "' " & _
'                              "and NP03='" & MSHFlexGrid2.TextMatrix(iRow, 12) & "' " & _
'                              "and NP04='" & MSHFlexGrid2.TextMatrix(iRow, 13) & "' " & _
'                              "and NP05='" & MSHFlexGrid2.TextMatrix(iRow, 14) & "' " & _
'                              "and NP07 in('102','109') " & _
'                              "and NP09>=" & strSrvDate(1) & " " & _
'                              "and NP06 is not null"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strCP27 = "19221111"
'                  strCP64 = "不催延展,不產生定稿"
'               End If
'            End If
'            '2013/12/20 END
            
            ' 案件性質
            strCP10 = "1717"
            'Modify By Sindy 2013/8/15 +," & strCP27 & ",'" & strCP64 & "'
            strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP27,CP64) " & _
                 "VALUES ('" & MSHFlexGrid2.TextMatrix(iRow, 11) & "','" & MSHFlexGrid2.TextMatrix(iRow, 12) & "','" & MSHFlexGrid2.TextMatrix(iRow, 13) & "','" & MSHFlexGrid2.TextMatrix(iRow, 14) & "'," & Val(strSrvDate(1)) & "," & _
                         "'','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(MSHFlexGrid2.TextMatrix(iRow, 11), MSHFlexGrid2.TextMatrix(iRow, 12), MSHFlexGrid2.TextMatrix(iRow, 13), MSHFlexGrid2.TextMatrix(iRow, 14))) & "','" & _
                         PUB_GetAKindSalesNo(MSHFlexGrid2.TextMatrix(iRow, 11), MSHFlexGrid2.TextMatrix(iRow, 12), MSHFlexGrid2.TextMatrix(iRow, 13), MSHFlexGrid2.TextMatrix(iRow, 14)) & "','" & strUserNum & "'," & _
                         "'" & "N" & "','" & "N" & "','" & "N" & "','',''," & strCP27 & ",'" & strCP64 & "') "
            cnnConnection.Execute strSql
            
            'Modify By Sindy 2013/12/20 本所案件全部不催延展,不出定稿
'            '未收文列印定稿
'            'Add By Sindy 2013/8/15 不催延展,不產生定稿 (+And InStr(strCP64, "不催延展,不產生定稿") = 0)
'            If Trim(MSHFlexGrid2.TextMatrix(iRow, 19)) = "" And MSHFlexGrid2.TextMatrix(iRow, 21) <> "Y" _
'               And InStr(strCP64, "不催延展,不產生定稿") = 0 Then
'               p_Recs2_Print = p_Recs2_Print + 1
'               ' 清除定稿例外欄位檔原有資料
'               EndLetter "12", strCP09, "01", strUserNum
'               ' 列印定稿
''               NowPrint strCP09, "12", "01", False, strUserNum, 0, , , , 1
'               'Modify By Sindy 2012/1/13
'               ET03 = "01"
'               iCopy = 1
'               If ET03 <> "" Then
'                  bolEmail = PUB_GetEMailFlag(MSHFlexGrid2.TextMatrix(iRow, 11) & MSHFlexGrid2.TextMatrix(iRow, 12) & MSHFlexGrid2.TextMatrix(iRow, 13) & MSHFlexGrid2.TextMatrix(iRow, 14), , , bolPlusPaper)
'                  If bolEmail Then
'                     '判斷是否EMail同時寄紙本
'                     If Not bolPlusPaper Then
'                        iCopy = 1
'                     End If
'                     NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
'                     boleFileSave = True
''                     MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
'                  Else
'                     NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy
'                  End If
'               End If
'               '2012/1/13 End
'            End If
            
         End If
      End If
   Next
   If bPrintAgain = False Then cnnConnection.CommitTrans
   
   'Add By Sindy 2012/1/13
   If boleFileSave = True Then
      MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
   End If
   '2012/1/13 End
   
   If bPrint = True Then
      If bPrintAgain = False Then MsgBox "將列印通知續展資料檢核清單，請換一般列印紙!!!"
      Printer.EndDoc
      If MsgBox("列印通知續展資料檢核清單，是否已列印成功？成功按「是」，不成功按「否」再列印一次！", vbYesNo + vbDefaultButton2) = vbNo Then
         bPrintAgain = True
         GoTo PrintAgain4
      End If
   End If
   Screen.MousePointer = vbDefault
   
   '產生Word檔
   bolRetry = True
   Screen.MousePointer = vbHourglass
   'Add By Sindy 2013/11/22
   cnnConnection.BeginTrans
   cnnConnection.Execute "delete from R020320"
   '2013/11/22 END
   For iRow = 1 To MSHFlexGrid2.Rows - 1
      '非本所
      'Modify By Sindy 2013/2/25 增加排除不可列印者
      'If MSHFlexGrid2.TextMatrix(iRow, 10) <> "Y" Then
      If MSHFlexGrid2.TextMatrix(iRow, 10) <> "Y" Then
      '2013/2/25 End
         m_AppAddr = MSHFlexGrid2.TextMatrix(iRow, 9) '申請人地址
         m_AppName = MSHFlexGrid2.TextMatrix(iRow, 3) '商標註冊人
         If GetIsNotPrintPer(m_AppName) = False Then 'Add By Sindy 2013/3/6 +if
            'Add By Sindy 2010/11/18
            strTempAppAddr = "": strCU01 = "": m_AppAddrZip = ""
            'Modify By Sindy 2012/6/19 註冊號9727839申請人為吳志忠此案件之故,再增加名稱若小於等於4個字的也是個人
            If Trim(MSHFlexGrid2.TextMatrix(iRow, 22)) <> "" Or Len(Trim(MSHFlexGrid2.TextMatrix(iRow, 3))) <= 4 Then
               '個人, 抓名稱及ID都相同者,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
   '            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' and cu11='" & Trim(MSHFlexGrid2.TextMatrix(iRow, 22)) & "' order by cu01 asc,cu02 asc "
   '            intI = 1
   '            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '            If intI = 1 Then
   '               If RsTemp.Fields("cu02") <> "0" Then
   '                  strCU01 = "" & RsTemp.Fields("cu01")
   '               Else
   '                  m_AppAddrZip = Trim("" & RsTemp.Fields("cu30"))
   '                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
   '               End If
   '            End If
               'Modify By Sindy 2011/12/20
               strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' and cu11='" & Trim(MSHFlexGrid2.TextMatrix(iRow, 22)) & "' order by cu01 asc,cu02 asc "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If bolErr2147467259 = True Then
                  bolErr2147467259 = False
   '               rsTmp.Close
                  strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & " ' and cu11='" & Trim(MSHFlexGrid2.TextMatrix(iRow, 22)) & "' order by cu01 asc,cu02 asc "
                  rsTmp.CursorLocation = adUseClient
                  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               End If
               If rsTmp.RecordCount > 0 Then
                  If rsTmp.Fields("cu02") <> "0" Then
                     strCU01 = "" & rsTmp.Fields("cu01")
                  Else
                     m_AppAddrZip = PUB_ChangeZIPToSir(Trim("" & rsTmp.Fields("cu30")))
                     strTempAppAddr = Trim("" & rsTmp.Fields("cu31"))
                  End If
               End If
               rsTmp.Close
            Else
               '非個人, 抓相同客戶名稱之最小客戶編號,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
   '            strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc "
   '            intI = 1
   '            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '            If intI = 1 Then
   '               If RsTemp.Fields("cu02") <> "0" Then
   '                  strCU01 = "" & RsTemp.Fields("cu01")
   '               Else
   '                  m_AppAddrZip = Trim("" & RsTemp.Fields("cu30"))
   '                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
   '               End If
   '            End If
               'Modify By Sindy 2011/12/20
               strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If bolErr2147467259 = True Then
                  bolErr2147467259 = False
   '               rsTmp.Close
                  strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & " ' order by cu01 asc,cu02 asc "
                  rsTmp.CursorLocation = adUseClient
                  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               End If
               If rsTmp.RecordCount > 0 Then
                  If rsTmp.Fields("cu02") <> "0" Then
                     strCU01 = "" & rsTmp.Fields("cu01")
                  Else
                     m_AppAddrZip = PUB_ChangeZIPToSir(Trim("" & rsTmp.Fields("cu30")))
                     strTempAppAddr = Trim("" & rsTmp.Fields("cu31"))
                  End If
               End If
               rsTmp.Close
            End If
            If strCU01 <> "" Then
               strSql = "SELECT * FROM customer WHERE cu01='" & strCU01 & "' and cu02='0' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  m_AppAddrZip = PUB_ChangeZIPToSir(Trim("" & RsTemp.Fields("cu30")))
                  strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
               End If
            End If
            If strTempAppAddr <> "" Then
               m_AppAddr = strTempAppAddr
            Else
               m_AppAddrZip = ""
            End If
            '2010/11/18 End
            'Add By Sindy 2013/6/3
            If m_AppAddrZip = "" Then
               m_AppAddrZip = PUB_ChangeZIPToSir(Left(PUB_AddrChangeZIPCode(m_AppAddr), 3))
            End If
            '2013/6/3 End
            m_AppDate = MSHFlexGrid2.TextMatrix(iRow, 4) '申請日期
            m_AppNum = MSHFlexGrid2.TextMatrix(iRow, 1) '註冊號
            m_TName = MSHFlexGrid2.TextMatrix(iRow, 2) '商標名稱
            m_Goods = MSHFlexGrid2.TextMatrix(iRow, 7) '分類
            m_DATEs = MSHFlexGrid2.TextMatrix(iRow, 5) '專用期(起)
            m_DATEe = MSHFlexGrid2.TextMatrix(iRow, 6) '專用期(止)
            'Modify By Sindy 2013/11/22
            strSql = "insert into R020320(AppZip,AppName,AppAddr,AppNum,TName,Goods,DATEs,DATEe)" & _
                     " values(" & CNULL(m_AppAddrZip) & "," & CNULL(m_AppName) & "," & CNULL(m_AppAddr) & "," & CNULL(m_AppNum) & _
                     "," & CNULL(ChgSQL(m_TName)) & "," & CNULL(m_Goods) & "," & CNULL(m_DATEs, True) & "," & CNULL(m_DATEe, True) & ")"
            cnnConnection.Execute strSql
            ' 列印定稿
'            m_AppDate = Val(Left(m_AppDate, 4)) & "年" & Mid(m_AppDate, 5, 2) & "月" & Right(m_AppDate, 2) & "日"
'            m_DATEs = Val(Left(m_DATEs, 4)) & "年" & Mid(m_DATEs, 5, 2) & "月" & Right(m_DATEs, 2) & "日"
'            m_DATEe = Val(Left(m_DATEe, 4)) & "年" & Mid(m_DATEe, 5, 2) & "月" & Right(m_DATEe, 2) & "日"
'            Call WordEdit(2)
            '2013/11/22 END
         End If
      End If
   Next iRow
   cnnConnection.CommitTrans 'Add By Sindy 2013/11/22
   
   'Add By Sindy 2013/11/25 同申請人多案一定稿
   strSql = "select * from R020320 order by APPZIP,APPADDR,APPNAME asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_AppAddrZip = "": m_AppAddr = "": m_AppName = "": tmp_TName = ""
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         tmp_AppAddrZip = "" & rsTmp.Fields("APPZIP")
         tmp_AppAddr = "" & rsTmp.Fields("AppAddr") '申請人地址
         tmp_AppName = "" & rsTmp.Fields("APPNAME") '商標註冊人
         m_AppNum = "" & rsTmp.Fields("APPNUM")
         m_TName = "" & rsTmp.Fields("TNAME")
         m_Goods = "" & rsTmp.Fields("GOODS")
         m_DATEs = "" & rsTmp.Fields("DATES")
         m_DATEe = "" & rsTmp.Fields("DATEE")
         m_DATEe = Val(Left(m_DATEe, 4)) & "年" & Mid(m_DATEe, 5, 2) & "月"
         If Not (m_AppAddrZip = "" And m_AppAddr = "" And m_AppName = "") And _
            (m_AppAddrZip <> tmp_AppAddrZip Or m_AppAddr <> tmp_AppAddr Or m_AppName <> tmp_AppName) Then
            ' 列印定稿
            tmp_TName = Right(tmp_TName, Len(tmp_TName) - 1)
            Call WordEdit(3)
            tmp_TName = ""
         End If
         m_AppAddrZip = tmp_AppAddrZip
         m_AppAddr = tmp_AppAddr
         m_AppName = tmp_AppName
         tmp_TName = tmp_TName & "、第" & m_AppNum & "號「" & m_TName & "」(第" & m_Goods & "類)"
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   If m_AppAddrZip <> "" Or m_AppAddr <> "" Or m_AppName <> "" Then
      ' 列印定稿
      tmp_TName = Right(tmp_TName, Len(tmp_TName) - 1)
      Call WordEdit(3)
   End If
   '2013/11/25 END
   
   If bolRetry = False Then
      g_WordAp.Visible = True
      g_WordAp.WindowState = wdWindowStateMaximize
   End If
   Screen.MousePointer = vbDefault
   
   Process2 = True
   Set g_WordAp = Nothing
   Set rsTmp = Nothing 'Add By Sindy 2011/12/20
   Exit Function
   
ErrHnd:
   'Add By Sindy 2011/12/20
   bolErr2147467259 = False
   If Err.Number = -2147467259 Then
      bolErr2147467259 = True
      '接著發生錯誤陳述式的下個陳述式開始執行
      Resume Next
   End If
   Set g_WordAp = Nothing
   Set rsTmp = Nothing
   '2011/12/20 End
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

'Add By Sindy 2013/2/25
'檢查此人是否為不可列印者
Private Function GetIsNotPrintPer(strPerName As String) As Boolean
Dim rs As ADODB.Recordset
   
   GetIsNotPrintPer = False
   'Modify By Sindy 2013/3/1 + and tbnp08='T'
   strExc(0) = "select * from tmbulletinnp where ltrim(rtrim(TBNP01))=ltrim(rtrim('" & strPerName & "')) and tbnp08='T'"
   intI = 1
   Set rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetIsNotPrintPer = True
   End If
   rs.Close
   Set rs = Nothing
End Function
