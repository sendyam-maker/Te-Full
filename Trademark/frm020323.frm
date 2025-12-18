VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020323 
   BorderStyle     =   1  '單線固定
   Caption         =   "台灣商標延展開拓函(智慧局)"
   ClientHeight    =   4670
   ClientLeft      =   2800
   ClientTop       =   3950
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4670
   ScaleWidth      =   6360
   Begin VB.CommandButton cmdPic 
      Caption         =   "抓圖檔"
      Height          =   340
      Left            =   5490
      TabIndex        =   28
      Top             =   2250
      Visible         =   0   'False
      Width           =   820
   End
   Begin VB.CommandButton Command3 
      Caption         =   "特定公司不列印者"
      Height          =   400
      Left            =   330
      TabIndex        =   27
      Top             =   90
      Width           =   1620
   End
   Begin VB.CheckBox Check1 
      Caption         =   "寄通知北所、分所信件"
      Height          =   285
      Left            =   690
      TabIndex        =   24
      Top             =   2670
      Value           =   1  '核取
      Width           =   2445
   End
   Begin VB.TextBox txtCntTot 
      Height          =   285
      Left            =   5310
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2670
      Width           =   585
   End
   Begin VB.TextBox txtCnt 
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2670
      Width           =   585
   End
   Begin VB.TextBox text1 
      Height          =   285
      Index           =   0
      Left            =   600
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frm020323.frx":0000
      Top             =   3450
      Width           =   5835
   End
   Begin VB.TextBox text1 
      Height          =   285
      Index           =   2
      Left            =   600
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2790
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.FileListBox File2 
      Height          =   180
      Left            =   630
      TabIndex        =   13
      Top             =   630
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtPath2 
      Height          =   285
      Left            =   1740
      TabIndex        =   2
      Text            =   "\\Sale1\XFER\BaireTrademark"
      Top             =   1890
      Width           =   3795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   345
      Left            =   5340
      TabIndex        =   3
      Top             =   1860
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdWord 
      Cancel          =   -1  'True
      Caption         =   "產生定稿(&W)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3840
      TabIndex        =   7
      Top             =   90
      Width           =   1260
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5265
      Left            =   180
      TabIndex        =   9
      Top             =   4500
      Visible         =   0   'False
      Width           =   11355
      _ExtentX        =   20020
      _ExtentY        =   9296
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<="
      Height          =   345
      Left            =   5550
      TabIndex        =   1
      Top             =   1500
      Width           =   345
   End
   Begin VB.TextBox txtPath1 
      Height          =   285
      Left            =   1740
      TabIndex        =   0
      Text            =   "C:\temp\商標延展\延展資料.xls"
      Top             =   1500
      Width           =   3795
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5160
      TabIndex        =   8
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdImPort 
      Caption         =   "匯入資料(&E)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2520
      TabIndex        =   6
      Top             =   90
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "先校正智權人員及所別資料, 請稍候 . . ."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   440
      Left            =   240
      TabIndex        =   26
      Top             =   990
      Visible         =   0   'False
      Width           =   5900
   End
   Begin VB.Label Label11 
      Caption         =   "開拓函檔位產生位置：C:\temp\商標延展\WordFile"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   495
      TabIndex        =   25
      Top             =   1200
      Width           =   5205
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "匯入前 (貝爾商標圖檔) 要先準備好"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   390
      TabIndex        =   23
      Top             =   780
      Width           =   4590
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "總筆數："
      ForeColor       =   &H00008000&
      Height          =   180
      Left            =   4620
      TabIndex        =   22
      Top             =   2730
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "專用期限止日："
      Height          =   180
      Left            =   -690
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "筆數："
      ForeColor       =   &H00008000&
      Height          =   180
      Left            =   3420
      TabIndex        =   18
      Top             =   2730
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "分機："
      Height          =   180
      Left            =   30
      TabIndex        =   16
      Top             =   2850
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "服務人員："
      Height          =   180
      Left            =   810
      TabIndex        =   15
      Top             =   2250
      Width           =   900
   End
   Begin MSForms.Label Label8 
      Height          =   200
      Left            =   1740
      TabIndex        =   14
      Top             =   2250
      Width           =   1365
      VariousPropertyBits=   27
      Size            =   "2408;353"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標圖檔案路徑："
      Height          =   180
      Left            =   270
      TabIndex        =   12
      Top             =   1950
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "注意：當程式正在執行〔產生定稿〕時，請暫時不要使用Word！"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   180
      TabIndex        =   11
      Top             =   3150
      Width           =   5835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "XLS檔案路徑："
      Height          =   180
      Left            =   495
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frm020323"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0已修改 label8
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Create By Sindy 2012/7/13
Option Explicit

Dim p_Recs1 As Integer
Dim m_WordFilePath As String
Dim m_intFileCnt As Integer, m_iRow As Integer
Dim bolRetry As Boolean '是否已發生錯誤且重試
Dim m_AppName As String '商標註冊人
Dim m_AppAddrZip As String, m_AppAddrZipOld As String '申請人地址郵遞區號
Dim m_AppAddr As String '申請人地址

'加入代表圖用
'Const msoBringInFrontOfText = 4
'Const msoFalse = 0
'Const msoLineSolid = 1
'Const msoLineSingle = 1
Const msoTrue = -1
'Const msoPictureAutomatic = 1
Dim intHeight As Integer ', intCnt As Integer
Dim ff3 As Integer, m_PrintRpt3 As Boolean, m_strFileName3 As String 'Add By Sindy 2013/1/28
Dim m_AttachPath As String
Dim m_ConSql As String
Dim m_ConSql2 As String 'Add By Sindy 2019/5/8
Dim pFtpSrv As String
Dim hConnection As Long
Dim m_ApplSales As String, m_ApplSalesST06 As String 'Add By Sindy 2019/2/1
Dim m_WordAp As Word.Application 'Add By Sindy 2019/3/27

'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean


Private Sub ResetGrid(ByRef p_Grid As MSHFlexGrid, Index As Integer)
   With p_Grid
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      If Index = 0 Then
         .FormatString = "審定號數|商標名稱|專用權人|郵遞區號|專用權人地址|專用期限|是否為本所案件|商標圖檔名|申請案號|申請日|代理人"
      End If
   End With
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdImPort_Click()
Dim i As Integer
Dim strTemp As String
Dim bolExecuteWord As Boolean
Dim intQ As Integer 'Add By Sindy 2018/3/15
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim bolReadExit As Boolean
Dim iRow As Integer
Dim dblFCnt As Integer
Dim ff1 As Integer
Dim rsTmp As New ADODB.Recordset
Dim rsA As New ADODB.Recordset 'Add By Sindy 2023/12/4
Dim strCU01 As String, strCU80 As String, strCU15 As String
Dim bolConn As Boolean
Dim strSubject As String 'Add By Sindy 2025/5/15
   
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
'   If text1(0).Text = "" Then
'      MsgBox "專用期限止日不可空白！", vbExclamation
'      text1(0).SetFocus
'      Exit Sub
'   End If
   
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
'      If InStr(txtPath2, ".") > 0 Then
'         For i = Len(txtPath2) To 1 Step -1
'            If Mid(txtPath2, i, 1) = "\" Then
'               txtPath2 = Mid(txtPath2, 1, i - 1)
'               Exit For
'            End If
'         Next i
'      Else
         If Right(txtPath2, 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
'      End If
      File2.path = txtPath2
      File2.Refresh
      If File2.ListCount = 0 Then
         MsgBox "找不到商標圖檔！"
         Exit Sub
      End If
'      If text1(2).Text = "" Then
'         MsgBox "分機不可空白！", vbExclamation
'         text1(2).SetFocus
'         Exit Sub
'      End If
   ElseIf intQ = vbNo Then '僅匯入作業
   Else '放棄執行
      Exit Sub
   End If
   '2018/3/15 END
   
   p_Recs1 = 0
   m_PrintRpt3 = False 'Add By Sindy 2013/1/28
   
   Screen.MousePointer = vbHourglass
   ResetGrid MSHFlexGrid1, 0
   
   '清除暫存檔資料
   strSql = "delete from BaireTrademark"
   cnnConnection.Execute strSql
   
   txtCnt.Text = "0": txtCntTot.Text = "0" 'Add By Sindy 2024/8/30
   bolReadExit = False: iRow = 2
   '資料夾裡會有多個.xls
   File2.path = Left(txtPath1, InStrRev(txtPath1, "\"))
   File2.Refresh
   For dblFCnt = 0 To File2.ListCount - 1
      If Right(UCase(File2.List(dblFCnt)), 4) <> UCase(".xls") Then GoTo ReadNextFile
      txtPath1 = Left(txtPath1, InStrRev(txtPath1, "\")) & File2.List(dblFCnt)
      DoEvents
      xlsSalesPoint.Workbooks.Open Left(txtPath1, InStrRev(txtPath1, "\")) & File2.List(dblFCnt)
      Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
      
      cnnConnection.BeginTrans: bolConn = True
      Do While bolReadExit = False
         '註冊號空白或商標名稱空白,則代表結束
         If wksaccrpt114.Range("B" & iRow).Value = "" Or _
            wksaccrpt114.Range("C" & iRow).Value = "" Then
            bolReadExit = True
         Else
            p_Recs1 = p_Recs1 + 1
            txtCntTot.Text = p_Recs1: DoEvents
            
            MSHFlexGrid1.Rows = p_Recs1 + 1
            'Modify By Sindy 2021/8/31 W.專用期限止日 => W.專用期間
            'XLS:
            'B        C           F        K      O              W            Q              S
            '申請案號,註冊/審定號,商標名稱,申請日,申請人中文名稱,專用期間,申請人中文地址,代理人
            'MSHFlexGrid1:
            '0        1        2        3        4            5        6              7          8        9      10
            '審定號數|商標名稱|專用權人|郵遞區號|專用權人地址|專用期限|是否為本所案件|商標圖檔名|申請案號|申請日|代理人
            MSHFlexGrid1.TextMatrix(p_Recs1, 8) = Trim(wksaccrpt114.Range("B" & iRow).Value) '申請案號
            MSHFlexGrid1.TextMatrix(p_Recs1, 0) = Trim(wksaccrpt114.Range("C" & iRow).Value) '註冊/審定號
            MSHFlexGrid1.TextMatrix(p_Recs1, 1) = Trim(wksaccrpt114.Range("F" & iRow).Value) '商標名稱
            MSHFlexGrid1.TextMatrix(p_Recs1, 9) = DBDATE(Trim(wksaccrpt114.Range("K" & iRow).Value)) '申請日
            MSHFlexGrid1.TextMatrix(p_Recs1, 2) = Trim(wksaccrpt114.Range("O" & iRow).Value) '申請人中文名稱
            '多人時,抓第一個
            If InStr(MSHFlexGrid1.TextMatrix(p_Recs1, 2), "、") > 0 Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 2) = Left(MSHFlexGrid1.TextMatrix(p_Recs1, 2), InStr(MSHFlexGrid1.TextMatrix(p_Recs1, 2), "、") - 1)
            End If
            'Modify By Sindy 2021/8/31 W.專用期限止日 => W.專用期間
            'MSHFlexGrid1.TextMatrix(p_Recs1, 5) = DBDATE(Trim(wksaccrpt114.Range("W" & iRow).Value)) '專用期限止日
            If InStr(Trim(wksaccrpt114.Range("W" & iRow).Value), "~") > 0 Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 5) = DBDATE(Mid(Trim(wksaccrpt114.Range("W" & iRow).Value), InStr(Trim(wksaccrpt114.Range("W" & iRow).Value), "~") + 1)) '專用期限止日
            Else
               MSHFlexGrid1.TextMatrix(p_Recs1, 5) = DBDATE(Trim(wksaccrpt114.Range("W" & iRow).Value)) '專用期限止日
            End If
            '2021/8/31 END
            
            '申請人中文地址 *********************
            MSHFlexGrid1.TextMatrix(p_Recs1, 4) = Trim(wksaccrpt114.Range("Q" & iRow).Value)
            '多地址時,抓第一個
            If InStr(MSHFlexGrid1.TextMatrix(p_Recs1, 4), "、") > 0 Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 4) = Left(MSHFlexGrid1.TextMatrix(p_Recs1, 4), InStr(MSHFlexGrid1.TextMatrix(p_Recs1, 4), "、") - 1)
            End If
            '去掉中國
            If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 4), 2) = "中國" Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 4) = Mid(MSHFlexGrid1.TextMatrix(p_Recs1, 4), 3)
            End If
            '去掉臺灣
            If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 4), 2) = "臺灣" Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 4) = Mid(MSHFlexGrid1.TextMatrix(p_Recs1, 4), 3)
            End If
            '郵遞區號
            MSHFlexGrid1.TextMatrix(p_Recs1, 3) = Left(PUB_AddrChangeZIPCode(MSHFlexGrid1.TextMatrix(p_Recs1, 4), , False), 3)
            '************************************
            
            MSHFlexGrid1.TextMatrix(p_Recs1, 10) = Trim(wksaccrpt114.Range("S" & iRow).Value) '代理人
            
            '商標圖檔名
            If Dir(txtPath2 & "\" & "T" & Val(MSHFlexGrid1.TextMatrix(p_Recs1, 0)) & ".gif") <> "" Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 7) = "T" & Val(MSHFlexGrid1.TextMatrix(p_Recs1, 0))
            ElseIf Dir(txtPath2 & "\" & "S" & Val(MSHFlexGrid1.TextMatrix(p_Recs1, 0)) & ".gif") <> "" Then
               MSHFlexGrid1.TextMatrix(p_Recs1, 7) = "S" & Val(MSHFlexGrid1.TextMatrix(p_Recs1, 0))
            End If
            
            '檢查是否為本所案件
            '審定號數先用原資料檢核,若Find不到資料,再看是否有須要補足碼數再檢核一次
   '               If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "S" Then '服務商標
   '                  'Modify By Sindy 2013/7/2 原為and TM08 in('4','5','6'),因T-086449及T-085452之故
   '                  strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & MSHFlexGrid1.TextMatrix(p_Recs1, 0) & "' and TM10='000' and TM28='1' and (TM08 in('4','5','6') or instr(TM58,'原為服務標章')>0 or instr(TM58,'原為聯合服務標章')>0) and tm29 is null and tm57 is null"
   '               Else '商標
            'Modify By Sindy 2013/8/13 閉卷不通知,銷卷要通知
            'Modify By Sindy 2013/8/15 銷卷也不通知
            MSHFlexGrid1.TextMatrix(p_Recs1, 6) = "N"
            If MSHFlexGrid1.TextMatrix(p_Recs1, 7) <> "" Then
               If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 7), 1) = "S" Or Left(MSHFlexGrid1.TextMatrix(p_Recs1, 7), 1) = "T" Then
                  strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & MSHFlexGrid1.TextMatrix(p_Recs1, 0) & "' and TM10='000' and TM28='1'" ' and TM08 in('1','2','3')
               Else
                  MsgBox "審定號（" & MSHFlexGrid1.TextMatrix(p_Recs1, 0) & "）之商標種類（" & Left(MSHFlexGrid1.TextMatrix(p_Recs1, 7), 1) & "）有問題，請確認！"
                  Exit Sub
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MSHFlexGrid1.TextMatrix(p_Recs1, 6) = "Y"
               End If
            End If
            '補足8碼檢核
            If MSHFlexGrid1.TextMatrix(p_Recs1, 6) = "N" Then
               strTemp = Right("00000000" & MSHFlexGrid1.TextMatrix(p_Recs1, 0), 8)
   '                  If Left(MSHFlexGrid1.TextMatrix(p_Recs1, 6), 1) = "S" Then '服務商標
   '                     'Modify By Sindy 2013/7/2 原為and TM08 in('4','5','6'),因T-086449及T-085452之故
   '                     strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & strTemp & "' and TM10='000' and TM28='1' and (TM08 in('4','5','6') or instr(TM58,'原為服務標章')>0 or instr(TM58,'原為聯合服務標章')>0) and tm29 is null and tm57 is null"
   '                  Else '商標
               strExc(0) = "SELECT * FROM TradeMark WHERE TM15='" & strTemp & "' and TM10='000' and TM28='1'" ' and TM08 in('1','2','3')
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MSHFlexGrid1.TextMatrix(p_Recs1, 6) = "Y"
               Else
                  MSHFlexGrid1.TextMatrix(p_Recs1, 6) = ""
               End If
            End If
            
            'Add By Sindy 2014/4/18 跨類商標會資料重覆出現,因此過濾重覆出現的審定號數,只收錄一筆
            strExc(0) = "SELECT * FROM BaireTrademark WHERE bt07='" & MSHFlexGrid1.TextMatrix(p_Recs1, 7) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
            '2014/4/18 END
               strSql = "insert into BaireTrademark(bt01,bt02,bt03,bt04,bt05,bt06,bt07,bt08,bt09,bt10,bt11)" & _
                        " values(" & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 0)) & "," & _
                        CNULL(ChgSQL(MSHFlexGrid1.TextMatrix(p_Recs1, 1))) & "," & _
                        CNULL(ChgSQL(MSHFlexGrid1.TextMatrix(p_Recs1, 2))) & "," & _
                        CNULL(ChgSQL(MSHFlexGrid1.TextMatrix(p_Recs1, 4))) & "," & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 5)) & "," & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 6)) & "," & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 7)) & "," & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 3)) & "," & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 8)) & "," & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 9)) & "," & _
                        CNULL(MSHFlexGrid1.TextMatrix(p_Recs1, 10)) & _
                        ")"
               cnnConnection.Execute strSql
            End If
         End If
         iRow = iRow + 1
      Loop
      cnnConnection.CommitTrans: bolConn = False
      
      '關閉
      xlsSalesPoint.Workbooks.Close
      '離開
      xlsSalesPoint.Quit
      Set wksaccrpt114 = Nothing
      
ReadNextFile:
      bolReadExit = False: iRow = 2
   Next dblFCnt
   Set xlsSalesPoint = Nothing
   
   'Add By Sindy 2019/9/24 匯入完畢,直接更新相關資料
   Screen.MousePointer = vbHourglass
   Label12.Visible = True 'Msg Box
   DoEvents
   'Add By Sindy 2019/9/19
   '利用郵遞區號更新所別
   strSql = "UPDATE bairetrademark" & _
            " SET bt13=(SELECT PZD10 FROM postzipdata WHERE substr(PZD01,1,3)=bt08 AND PZD10 IS NOT NULL GROUP BY substr(PZD01,1,3),PZD10)" & _
            " Where bt08 Is Not Null"
   cnnConnection.Execute strSql
   '利用地址更新所別
   strSql = "UPDATE bairetrademark SET bt13=addrgetst06(bt04) WHERE bt13 IS NULL and addrgetst06(bt04) is not null"
   cnnConnection.Execute strSql
   '2019/9/19 END
   'Add by Sindy 2019/5/8 先校正智權人員及所別資料
'1.先依郵遞區號抓所別
'2.再依地址抓所別
'3.檢查申請人名稱為大於4個字且非個人的本所客戶（cu02=0），
'      其負責的智權人員及所別（所別要同地址判斷出來一樣的所別才行，不然不採用）（國外部同仁?）
'4.列印時1.依申請人名稱大於小於4個字的+郵遞區號,其他的只判斷申請人名稱
'5.組多個註冊號數
   '非本所案件
   'strSql = "select * from bairetrademark Where bt06 Is Null order by bt01 asc"
   strSql = "SELECT bt03,bt04,bt08,bt13" & _
            " FROM bairetrademark WHERE bt06 IS NULL" & _
            " GROUP BY bt03,bt04,bt08,bt13" & _
            " order by bt03,bt04,bt08,bt13"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
      .MoveFirst
      Do While Not .EOF
         m_AppName = Trim("" & .Fields("bt03")) '專用權人
         m_AppAddr = "" & .Fields("bt04") '專用權人地址
         
         'Add by Sindy 2019/10/2 排除對造
         '剔除商標案且案件性質為1202(核駁前先行通知)
         '    商標案(CFC/S)案件性質202(申請意見書)及303(延期)
         '    所有專利案件性質404(延期)
         strSql = "select cp40 from caseprogress where cp40 is not null" & _
                  " and not((InStr(cp01,'T')>0 or cp01='CFC' or cp01='S') And (cp10='1202' or cp10='202' or cp10='303'))" & _
                  " and not((InStr(cp01,'P')>0 or cp01='FG') And (cp10='404'))" & _
                  " and cp40='" & m_AppName & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then '有對造 BT13.所別
            'Add By Sindy 2020/9/4
            If "" & .Fields("bt04") = "" Then
               strSql = "UPDATE bairetrademark SET " & _
                        " BT13=null" & _
                        ",BT14='對造'" & _
                        " where bt03=" & CNULL(.Fields("bt03")) & _
                        " and bt04 is null"
            Else
            '2020/9/4 END
               strSql = "UPDATE bairetrademark SET " & _
                        " BT13=null" & _
                        ",BT14='對造'" & _
                        " where bt03=" & CNULL(.Fields("bt03")) & _
                        " and bt04=" & CNULL(.Fields("bt04"))
            End If
            cnnConnection.Execute strSql
         Else
         '2019/10/2 END
            strCU01 = "": strCU80 = "": m_ApplSales = "": m_ApplSalesST06 = ""
            '非個人, 抓相同客戶名稱之最小客戶編號,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
            If Len(m_AppName) > 4 Then
               '增加所別判斷
               'Modify By Sindy 2019/9/23 + and cu02='0' and cu15<>'0'
               'Modify By Sindy 2020/9/3 + 取消 and cu02='0' 因還是有客戶用舊名稱 ex:聯意製作股份有限公司(周哲丞)
               If "" & .Fields("bt13") <> "" Then
                  '名稱+非個人+所別
                  strSql = "SELECT * FROM customer,staff WHERE cu04='" & m_AppName & "'" & _
                           " and cu15<>'0' and cu13=st01(+) and st06='" & .Fields("bt13") & "'" & _
                           " order by cu01||cu02 asc"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     'Add By Sindy 2020/9/3 則抓CU02='0'的資訊
                     If rsA.Fields("cu02") <> "0" Then
                        strSql = "SELECT * FROM customer,staff WHERE cu01='" & rsA.Fields("cu01") & "' and cu02='0'" & _
                                 " and cu15<>'0' and cu13=st01(+) and st06='" & .Fields("bt13") & "'" & _
                                 " order by cu01||cu02 asc"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           strCU01 = RsTemp.Fields("cu01") & RsTemp.Fields("cu02")
                           strCU80 = "" & RsTemp.Fields("cu80")
                           strCU15 = "" & RsTemp.Fields("CU15") 'Add By Sindy 2020/3/10
                           m_ApplSales = "" & RsTemp.Fields("cu13")
                           m_ApplSalesST06 = "" & RsTemp.Fields("st06")
                        End If
                     Else
                     '2020/9/3 END
                        strCU01 = rsA.Fields("cu01") & rsA.Fields("cu02")
                        strCU80 = "" & rsA.Fields("cu80")
                        strCU15 = "" & rsA.Fields("CU15") 'Add By Sindy 2020/3/10
                        m_ApplSales = "" & rsA.Fields("cu13")
                        m_ApplSalesST06 = "" & rsA.Fields("st06")
                     End If
                  End If
               End If
               If strCU01 = "" Then
                  '名稱+非個人:抓相同客戶名稱之最小客戶編號
                  strSql = "SELECT cu01,cu02,cu13,cu80,cu15,st06" & _
                           " FROM customer,staff WHERE cu04='" & m_AppName & "'" & _
                           " and cu15<>'0' and cu13=st01(+)" & _
                           " order by cu01||cu02 asc"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     rsA.MoveFirst
                     Do While Not rsA.EOF = True
                        'If m_ApplSalesST06 = "" Or m_ApplSalesST06 = rsA.Fields("st06") Then
                        'Modify By Sindy 2021/1/13 Mark; 例:勤億蛋品科技股份有限公司(330)郵遞區號檔屬於北所, 但確實是中所的客戶(S23/99031.張力允)
                        'If m_ApplSalesST06 = rsA.Fields("st06") Then
                        '2021/1/13 END
                           'Add By Sindy 2020/9/3 則抓CU02='0'的資訊
                           If rsA.Fields("cu02") <> "0" Then
                              strSql = "SELECT cu01,cu02,cu13,cu80,cu15,st06" & _
                                       " FROM customer,staff WHERE cu01='" & rsA.Fields("cu01") & "' and cu02='0'" & _
                                       " and cu15<>'0' and cu13=st01(+)" & _
                                       " order by cu01||cu02 asc"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                              If intI = 1 Then
                                 strCU01 = RsTemp.Fields("cu01") & RsTemp.Fields("cu02")
                                 strCU80 = "" & RsTemp.Fields("cu80")
                                 strCU15 = "" & RsTemp.Fields("CU15") 'Add By Sindy 2020/3/10
                                 m_ApplSales = "" & RsTemp.Fields("cu13")
                                 m_ApplSalesST06 = "" & RsTemp.Fields("st06")
                                 Exit Do
                              End If
                           Else
                           '2020/9/3 END
                              strCU01 = rsA.Fields("cu01") & rsA.Fields("cu02")
                              strCU80 = "" & rsA.Fields("cu80")
                              strCU15 = "" & rsA.Fields("CU15") 'Add By Sindy 2020/3/10
                              m_ApplSales = "" & rsA.Fields("cu13")
                              m_ApplSalesST06 = "" & rsA.Fields("st06")
                              Exit Do
                           End If
'                        Else
'                           strCU01 = "": strCU80 = "": m_ApplSales = "": m_ApplSalesST06 = ""
'                           Exit Do
                        'End If
                        rsA.MoveNext
                     Loop
                  End If
               End If
            'Add By Sindy 2020/3/10
            '個人, 抓相同客戶名稱之最小客戶編號,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
            Else
               '增加所別判斷
               'Modify By Sindy 2020/9/3 + 取消 and cu02='0' 因還是有客戶用舊名稱 ex:聯意製作股份有限公司(周哲丞)
               If "" & .Fields("bt13") <> "" Then
                  '名稱+個人+郵遞區號+所別 (CU30='５４０')
                  strSql = "SELECT * FROM customer,staff WHERE cu04='" & m_AppName & "'" & _
                           " and cu15='0' and cu30='" & PUB_ChangeZIPToSir("" & .Fields("bt08")) & "'" & _
                           " and cu13=st01(+) and st06='" & .Fields("bt13") & "'" & _
                           " order by cu01||cu02 asc"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     'Add By Sindy 2020/9/3 則抓CU02='0'的資訊
                     If rsA.Fields("cu02") <> "0" Then
                        strSql = "SELECT * FROM customer,staff WHERE cu01='" & rsA.Fields("cu01") & "' and cu02='0'" & _
                                 " and cu15='0' and cu30='" & PUB_ChangeZIPToSir("" & .Fields("bt08")) & "'" & _
                                 " and cu13=st01(+) and st06='" & .Fields("bt13") & "'" & _
                                 " order by cu01||cu02 asc"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           strCU01 = RsTemp.Fields("cu01") & RsTemp.Fields("cu02") '吳耀全
                           strCU80 = "" & RsTemp.Fields("cu80")
                           strCU15 = "" & RsTemp.Fields("CU15")
                           m_ApplSales = "" & RsTemp.Fields("cu13")
                           m_ApplSalesST06 = "" & RsTemp.Fields("st06")
                        End If
                     Else
                     '2020/9/3 END
                        strCU01 = rsA.Fields("cu01") & rsA.Fields("cu02") '吳耀全
                        strCU80 = "" & rsA.Fields("cu80")
                        strCU15 = "" & rsA.Fields("CU15")
                        m_ApplSales = "" & rsA.Fields("cu13")
                        m_ApplSalesST06 = "" & rsA.Fields("st06")
                     End If
                  End If
               End If
               If strCU01 = "" Then
                  '名稱+個人+郵遞區號 (CU30='５４０')
                  strSql = "SELECT * FROM customer,staff WHERE cu04='" & m_AppName & "'" & _
                           " and cu15='0' and cu30='" & PUB_ChangeZIPToSir("" & .Fields("bt08")) & "'" & _
                           " and cu13=st01(+)" & _
                           " order by cu01||cu02 asc"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     rsA.MoveFirst
                     Do While Not rsA.EOF = True
                        'If m_ApplSalesST06 = "" Or m_ApplSalesST06 = rsA.Fields("st06") Then
                        'Modify By Sindy 2021/1/13 Mark; 例:勤億蛋品科技股份有限公司(330)郵遞區號檔屬於北所, 但確實是中所的客戶(S23/99031.張力允)
                        'If m_ApplSalesST06 = rsA.Fields("st06") Then
                        '2021/1/13 END
                           'Add By Sindy 2020/9/3 則抓CU02='0'的資訊
                           If rsA.Fields("cu02") <> "0" Then
                              strSql = "SELECT * FROM customer,staff WHERE cu01='" & rsA.Fields("cu01") & "' and cu02='0'" & _
                                       " and cu15='0' and cu30='" & PUB_ChangeZIPToSir("" & .Fields("bt08")) & "'" & _
                                       " and cu13=st01(+)" & _
                                       " order by cu01||cu02 asc"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                              If intI = 1 Then
                                 strCU01 = RsTemp.Fields("cu01") & RsTemp.Fields("cu02")
                                 strCU80 = "" & RsTemp.Fields("cu80")
                                 strCU15 = "" & RsTemp.Fields("CU15") 'Add By Sindy 2020/3/10
                                 m_ApplSales = "" & RsTemp.Fields("cu13")
                                 m_ApplSalesST06 = "" & RsTemp.Fields("st06")
                                 Exit Do
                              End If
                           Else
                           '2020/9/3 END
                              strCU01 = rsA.Fields("cu01") & rsA.Fields("cu02")
                              strCU80 = "" & rsA.Fields("cu80")
                              strCU15 = "" & rsA.Fields("CU15") 'Add By Sindy 2020/3/10
                              m_ApplSales = "" & rsA.Fields("cu13")
                              m_ApplSalesST06 = "" & rsA.Fields("st06")
                              Exit Do
                           End If
'                        Else
'                           strCU01 = "": strCU80 = "": m_ApplSales = "": m_ApplSalesST06 = ""
'                           Exit Do
                        'End If
                        rsA.MoveNext
                     Loop
                  End If
               End If
               '2020/3/10 END
            End If
            If strCU01 <> "" And m_ApplSales <> "" And m_ApplSalesST06 <> "" Then
               'Modify By Sindy 2019/3/20 '下列客戶狀態採用公報地址
               'Modify By Sindy 2020/3/10 增加"解散","廢止","撤銷","停銷,"停業","死亡"
               'Modify By Sindy 2023/2/24 針對客戶狀態CU80的控制，原程式判斷若為該10項時改採用公報地址，
               '                          請改為CU80非空白且非「解除對造」時一律採用公報地址。
'               If strCU80 = "刪址" Or _
'                  strCU80 = "遷移不明" Or _
'                  strCU80 = "其他" Or _
'                  strCU80 = "業務自行處理" Or _
'                  strCU80 = "解散" Or _
'                  strCU80 = "廢止" Or _
'                  strCU80 = "撤銷" Or _
'                  strCU80 = "停銷" Or _
'                  strCU80 = "停業" Or _
'                  strCU80 = "死亡" Then
               If strCU80 <> "" And strCU80 <> "解除對造" Then
               '2023/2/24 END
                  'Modify By Sindy 2020/3/10
                  'strCU01 = ""
                  strCU01 = strCU15
                  '2020/3/10 END
               End If
               
               'Add By Sindy 2022/4/1 增加檢查員工是否在職, 若離職則換掛ACC090.A0909
               If ChkStaffST04(m_ApplSales, False) = True Then
                  'Added by Lydai 2023/12/26
                  If strSrvDate(1) >= 新部門啟用日 Then
                      strSql = "select decode(a0921,null,a0909,decode(a0924,null,oman,a0924) ) as a0909,st06 " & _
                               "from acc090,staff,acc090new,setspecman where st01='" & m_ApplSales & "' and st15=a0901(+) and st93=a0921(+) and ocode='程式管理人員' "
                  Else
                  'end 2023/12/26
                      strSql = "select A0909,st06 from ACC090,staff where st01='" & m_ApplSales & "' and st15=a0901(+)"
                  End If
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     If "" & RsTemp.Fields("A0909") <> "" Then
                        m_ApplSales = RsTemp.Fields("A0909")
                        m_ApplSalesST06 = RsTemp.Fields("st06")
                     End If
                  End If
               End If
               '2022/4/1 END
               
               'BT12.智權人員 BT13.所別 BT14.客戶編號
               strSql = "UPDATE bairetrademark SET " & _
                        " BT12=" & CNULL(m_ApplSales) & _
                        ",BT13=" & CNULL(m_ApplSalesST06) & _
                        ",BT14=" & CNULL(strCU01) & _
                        " where bt03=" & CNULL(.Fields("bt03")) & _
                        " and bt04=" & CNULL(.Fields("bt04"))
               cnnConnection.Execute strSql
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
   rsTmp.Close
   
   '所別有大於4的,如5,更新為空白
   strSql = "UPDATE bairetrademark SET bt13=null WHERE bt13>'4'"
   cnnConnection.Execute strSql, intI
   'Modify By Sindy 2019/10/4 申請人名稱或地址凡有?號均不出開拓函
   strSql = "UPDATE bairetrademark SET bt13=null WHERE instr(bt03,'?')>0 or instr(bt04,'?')>0"
   cnnConnection.Execute strSql, intI
   Label12.Visible = False 'Msg Box
   DoEvents
   '2019/5/8 END
   
   Screen.MousePointer = vbDefault
   'Add By Sindy 2013/1/28
   If m_PrintRpt3 = True Then
      Close ff3
      MsgBox "資料匯入完畢！新增時有錯誤資料，請至" & Left(txtPath1, InStrRev(txtPath1, "\")) & m_strFileName3 & "查看"
   Else
   '2013/1/28 End
      Call QueryData '重新查詢要產生開拓函的筆數
'      '檢查是否有”沒有商標圖檔名”的資料
'      strExc(0) = "SELECT count(*) FROM BaireTrademark WHERE (bt07 is null or bt07='')"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If RsTemp.Fields(0) > 0 Then
'            'Modify By Sindy 2019/3/29 產生無代表圖的清單
'            If ff1 > 0 Then Close #ff1
'            ff1 = FreeFile
'            Open PUB_Getdesktop & "\" & strSrvDate(1) & "台灣商標延展無代表圖清單.txt" For Output As ff1
'            Print #ff1, "備註：改字型Fixedsys標準11號字以直式上下左右各10MM列印"
'            Print #ff1, "審定號數   商標名稱"
'            Print #ff1, "========== ========================================================"
'            strExc(0) = "SELECT * FROM BaireTrademark WHERE (bt07 is null or bt07='') order by BT01 asc"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               RsTemp.MoveFirst
'               Do While Not RsTemp.EOF
'                  Print #ff1, convForm(CheckStr(RsTemp.Fields("BT01")), 10) & " " & RsTemp.Fields("BT02")
'                  RsTemp.MoveNext
'               Loop
'            End If
'            Close ff1
'            '2019/3/29 END
'            MsgBox "<沒有商標圖檔名> 的資料, 共計 " & RsTemp.RecordCount & " 筆" & vbCrLf & vbCrLf & _
'                   "清單放置:" & PUB_Getdesktop & "\" & strSrvDate(1) & "台灣商標延展無代表圖清單.txt" & vbCrLf & vbCrLf & _
'                   "資料匯入完畢！", vbExclamation, "資料有問題"
'            Exit Sub
'         End If
'      End If
      
      'Add By Sindy 2024/3/22
      PUB_SendMail strUserNum, "97038", "", "台灣商標延展開拓函(智慧局),資料匯入完畢", "將 bairetrademark 匯出至Excel保留資料, 以供有問題時可以查看！", , , , , , , , , , True, False
      '2024/3/22 END
      
      'Modify By Sindy 2025/5/15 Move到此處發信
      strSubject = Me.Caption & "，電子檔已產生完畢！"
      PUB_SendMail strUserNum, strUserNum, "", strSubject, strSubject, , , , , , , , , , , False
      If Check1.Visible = True And Check1.Value = 1 Then
'            varTmp = Split(Pub_GetSpecMan("台灣商標延展開拓函分所收受者"), ";")
'            For i = 0 To UBound(varTmp)
'               If GetPrjSalesNM_2(GetPrjSalesNM(CStr(varTmp(i))), , strST06) <> "" Then
'                  If strST06 = "1" Then
'                     strEmp1 = strEmp1 & ";" & varTmp(i)
'                  Else
'                     strEmp2 = strEmp2 & ";" & varTmp(i)
'                  End If
'               End If
'            Next i
'            If strEmp1 <> "" Then
'               strEmp1 = Mid(strEmp1, 2) '北所
'               PUB_SendMail strUserNum, strEmp1, "", strSubject, "Dear Sirs," & vbCrLf & vbCrLf & "　　商標處已將延展資料上傳完畢，請至 Server 做後續處理。" & vbCrLf & _
'                  "請至 " & Replace(m_WordFilePath, "c:", "\\" & PUB_ReadHostName) & " 資料夾中列印開拓函。", , , , , , , , , , , False
'            End If
'            If strEmp2 <> "" Then
'               strEmp2 = Mid(strEmp2, 2) '分所
'               PUB_SendMail strUserNum, strEmp2, "", strSubject, "Dear Sirs," & vbCrLf & vbCrLf & "　　北所已將延展資料上傳完畢，各分所可以進行其作業。", _
'                            , , , , , , , , , , False
'            End If
         
         '北所
'         PUB_SendMail strUserNum, Pub_GetSpecMan("台灣商標延展開拓函北所收受者"), "", strSubject, "Dear Sirs," & vbCrLf & vbCrLf & "　　商標處已將延展資料上傳完畢，請至 Server 做後續處理。" & vbCrLf & _
'                        "請至 " & Replace(m_WordFilePath, "c:", "\\" & PUB_ReadHostName) & " 資料夾中列印開拓函。", , , , , , , , , , , False
         
         '分所
         PUB_SendMail strUserNum, Pub_GetSpecMan("台灣商標延展開拓函分所收受者") & ";" & Pub_GetSpecMan("台灣商標延展開拓函北所收受者"), "", _
            strSubject, "Dear Sirs," & vbCrLf & vbCrLf & "　　" & IIf(PUB_GetST06(strUserNum) = "1", "北所", IIf(PUB_GetST06(strUserNum) = "2", "中所", IIf(PUB_GetST06(strUserNum) = "3", "南所", "高所"))) & _
            "已將延展資料上傳完畢，各分所可以進行其作業。", _
                        , , , , , , , , , , False
      End If
      
      If bolExecuteWord = True Then
         Call cmdWord_Click
      Else
         MsgBox "資料匯入完畢！"
      End If
   End If
   
   Set rsTmp = Nothing
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
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      MsgBox "第" & Left(txtPath1, InStrRev(txtPath1, "\")) & File2.List(dblFCnt) & "檔," & vbCrLf & vbCrLf & Err.Description & vbCrLf & strSql, vbCritical
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
      Open Left(txtPath1, InStrRev(txtPath1, "\")) & m_strFileName3 For Output As ff3
   End If
   Print #ff3, strSql
End Sub

'Add By Sindy 2025/5/9
Private Sub CmdPic_Click()
Dim dblFCnt As Double

   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   strExc(9) = Dir(txtPath2 & "\*1773536*.gif", vbNormal) 'Find單檔名稱OK,多檔會隨機抓一個檔名
   
   File2.path = txtPath2.Text
   File2.Refresh
   'strTotRow = File2.ListCount
   For dblFCnt = 0 To File2.ListCount - 1
      strExc(10) = File2.List(dblFCnt)
   Next dblFCnt
End Sub

Private Sub Command1_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.xls"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "Excel 檔案 (*.xls)|*.xls"
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         SaveSetting "TAIE", "Frm020323", UCase(Me.Name) & "Dir", .FileName
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

'Add By Sindy 2021/11/2
Private Sub Command3_Click()
   If CheckUse("frm030617", strExec) = True Then
      frm030617.Hide
      frm030617.SSTab1.TabVisible(1) = False
      frm030617.Caption = "特定公司不列印者"
      frm030617.Show
   End If
End Sub

Private Sub Form_Load()
   'Add By Sindy 2021/10/7
   m_bInsert = IsUserHasRightOfFunction("frm020323", strAdd, False)
'   m_bUpdate = IsUserHasRightOfFunction("frm020323", strEdit, False)
'   m_bDelete = IsUserHasRightOfFunction("frm020323", strDel, False)
'   m_bQuery = IsUserHasRightOfFunction("frm020323", strFind, False)
   '2021/10/7 END
   
   MoveFormToCenter Me
   Label8.Caption = strUserName
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   'Modify By Sindy 2019/7/31
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
      m_WordFilePath = PUB_Getdesktop & "\商標延展\WordFile"
      If Dir(PUB_Getdesktop & "\商標延展", vbDirectory) = "" Then
         MkDir PUB_Getdesktop & "\商標延展"
      End If
      If Dir(PUB_Getdesktop & "\商標延展\WordFile", vbDirectory) = "" Then
         MkDir PUB_Getdesktop & "\商標延展\WordFile"
      End If
   Else
   '2019/7/31 END
      m_WordFilePath = "c:\temp\商標延展\WordFile"
      If Dir("c:\temp", vbDirectory) = "" Then
         MkDir "c:\temp"
      End If
      If Dir("c:\temp\商標延展", vbDirectory) = "" Then
         MkDir "c:\temp\商標延展"
      End If
      If Dir("c:\temp\商標延展\WordFile", vbDirectory) = "" Then
         MkDir "c:\temp\商標延展\WordFile"
      End If
   End If
   ChDir App.path 'Add By Sindy 2020/3/10 釋放資料夾權限
   
   txtPath2.Text = Pub_GetSpecMan("內商開拓資料存放路徑") & "\BaireTrademark" 'Modify By Sindy 2023/8/1
   txtPath2.Enabled = False 'Add By Sindy 2024/8/30
   
   '商標處程序才能轉檔
   'Modify By Sindy 2021/10/7 非本所客戶商標延展開拓函-開放予南所鄭鈺華及蘇嫄媛兩位可以匯入資料。
   '                          秀玲設了A5005鄭鈺華及N1等級蘇嫄媛
   'If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "P22" Then
   If m_bInsert = True Or Pub_StrUserSt03 = "M51" Then
   '2021/10/7 END
      cmdImport.Visible = True
      Check1.Visible = True
      Label10.Visible = True
      Label9.Visible = True
      txtCntTot.Visible = True
   Else
      cmdImport.Visible = False
      Check1.Visible = False
      Label10.Visible = False
      Label9.Visible = False
      txtCntTot.Visible = False
   End If
      
'   If Me.Tag = "北所" Then
'      Label9.Visible = True
'      txtCntTot.Visible = True
'   Else
'      Label9.Visible = False
'      txtCntTot.Visible = False
'   End If
   
   'Add By Sindy 2021/11/2
   Command3.Visible = False
   If CheckUse("frm030617", strExec, False) = True Then
      Command3.Visible = True
   End If
   '2021/11/2 END
   
   Me.Tag = "北所"
'   If PUB_GetST06(strUserNum) = "1" Then '北所
'      Me.Tag = "北所"
'      cmdImPort.Visible = True
'   Else
   If PUB_GetST06(strUserNum) = "2" Then '中所
      Me.Tag = "中所"
   ElseIf PUB_GetST06(strUserNum) = "3" Then '南所
      Me.Tag = "南所"
   ElseIf PUB_GetST06(strUserNum) = "4" Then '高所
      Me.Tag = "高所"
   End If
   
   If QueryData = False Then
      CmdWord.Enabled = False
   Else
      CmdWord.Enabled = True
   End If
   
On Error GoTo ErrHnd 'Add By Sindy 2021/1/6

   '讀取前次設定路徑
   txtPath1 = GetSetting("TAIE", "Frm020323", UCase(Me.Name) & "Dir", "")
   If txtPath1 <> "" Then
      strExc(1) = Left(txtPath1, InStrRev(txtPath1, "\"))
      strExc(0) = Dir(strExc(1) & "*.Xls")
      If Dir(txtPath1) <> "" Then
         txtPath1 = strExc(1) & strExc(0)
      End If
   End If
   
   'Add By Sindy 2025/5/9
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      cmdPic.Visible = True
   End If
   '2025/5/9 END
   
   Exit Sub 'Add By Sindy 2021/1/6

'Add By Sindy 2021/1/6
ErrHnd:
   If Err.Number <> 0 Then
      txtPath1 = ""
      SaveSetting "TAIE", "Frm020323", UCase(Me.Name) & "Dir", txtPath1.Text
   End If
'2021/1/6 END
End Sub

Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   QueryData = False
   txtCnt.Text = "0": txtCntTot.Text = "0"
   
   strExc(0) = "SELECT count(*) FROM BaireTrademark"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         txtCntTot.Text = RsTemp.Fields(0)
      Else
         txtCntTot.Text = ""
      End If
   Else
      txtCntTot.Text = ""
   End If
   
'   'Modify By Sindy 2019/5/8
'   '沒掛智權同仁,中文名稱無?號,中文地址無?號,無郵遞區號
'   m_ConSql = " and bt13 is null" & _
'              " and instr(bt03,'?')=0" & _
'              " and instr(bt04,'?')=0" & _
'              " and bt08 is not null"
'   m_ConSql2 = ""
'   If Me.Tag = "北所" Then
'      m_ConSql = m_ConSql & " and (PZD10='1' or (bt08 is null and addrgetst06(bt04)='1')) "
'      m_ConSql2 = " and (addrgetst06(bt04)='' or bt13='1'" & _
'                         " or ((PZD10='1' or (bt08 is null and addrgetst06(bt04)='1')) and (instr(bt03,'?')>0 or instr(bt04,'?')>0 or bt08 is null))" & _
'                       ")"
'   ElseIf Me.Tag = "中所" Then
'      m_ConSql = m_ConSql & " and (PZD10='2' or (bt08 is null and addrgetst06(bt04)='2')) "
'      m_ConSql2 = " and (bt13='2'" & _
'                         " or ((PZD10='2' or (bt08 is null and addrgetst06(bt04)='2')) and (instr(bt03,'?')>0 or instr(bt04,'?')>0 or bt08 is null))" & _
'                       ")"
'   ElseIf Me.Tag = "南所" Then
'      m_ConSql = m_ConSql & " and (PZD10='3' or (bt08 is null and addrgetst06(bt04)='3')) "
'      m_ConSql2 = " and (bt13='3'" & _
'                         " or ((PZD10='3' or (bt08 is null and addrgetst06(bt04)='3')) and (instr(bt03,'?')>0 or instr(bt04,'?')>0 or bt08 is null))" & _
'                       ")"
'   ElseIf Me.Tag = "高所" Then
'      m_ConSql = m_ConSql & " and (PZD10='4' or (bt08 is null and addrgetst06(bt04)='4')) "
'      m_ConSql2 = " and (bt13='4'" & _
'                         " or ((PZD10='4' or (bt08 is null and addrgetst06(bt04)='4')) and (instr(bt03,'?')>0 or instr(bt04,'?')>0 or bt08 is null))" & _
'                       ")"
'   End If
   
   'Modify By Sindy 2019/5/8
   '沒掛智權同仁,中文名稱無?號,中文地址無?號,無郵遞區號
   'Modify By Sindy 2019/9/19 +  or bt13 is not null)
'   m_ConSql = " and bt12 is null" & _
'              " and instr(bt03,'?')=0" & _
'              " and instr(bt04,'?')=0"
'   m_ConSql2 = " and not (bt12 is null" & _
'              " and instr(bt03,'?')=0" & _
'              " and instr(bt04,'?')=0)"
   'Modify By Sindy 2019/10/4 申請人名稱或地址凡有?號均不出開拓函
   m_ConSql = " and bt12 is null" & _
              " and instr(bt03,'?')=0" & _
              " and instr(bt04,'?')=0"
   m_ConSql2 = " and bt12 is not null" & _
              " and instr(bt03,'?')=0" & _
              " and instr(bt04,'?')=0"
   If Me.Tag = "北所" Then
      m_ConSql = m_ConSql & " and (BT14 IS NULL OR BT14<>'對造') and bt13='1'"
      'and bt08 is not null => 瑞典,日本...
      m_ConSql2 = m_ConSql2 & " and (BT14 IS NULL OR BT14<>'對造') and (bt13 is null or bt13='1') and bt08 is not null"
   ElseIf Me.Tag = "中所" Then
      m_ConSql = m_ConSql & " and (BT14 IS NULL OR BT14<>'對造') and bt13='2'"
      m_ConSql2 = m_ConSql2 & " and (BT14 IS NULL OR BT14<>'對造') and bt13='2'"
   ElseIf Me.Tag = "南所" Then
      m_ConSql = m_ConSql & " and (BT14 IS NULL OR BT14<>'對造') and bt13='3'"
      m_ConSql2 = m_ConSql2 & " and (BT14 IS NULL OR BT14<>'對造') and bt13='3'"
   ElseIf Me.Tag = "高所" Then
      m_ConSql = m_ConSql & " and (BT14 IS NULL OR BT14<>'對造') and bt13='4'"
      m_ConSql2 = m_ConSql2 & " and (BT14 IS NULL OR BT14<>'對造') and bt13='4'"
   End If
   
   '非本所案件，排除國內商標公報特定公司不列印者，排除國外客戶
'   strSql = "select count(distinct bt03) as tt from bairetrademark,tmbulletinnp,(select substr(PZD01,1,3) PZD01,PZD10 from postzipdata group by substr(PZD01,1,3),PZD10) Z" & _
'            " Where bt06 Is Null and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
'            " and bt08=PZD01(+) " & m_ConSql
'   strSql = strSql & " union all" & _
'            " select count(distinct bt03) as tt from bairetrademark,tmbulletinnp,(select substr(PZD01,1,3) PZD01,PZD10 from postzipdata group by substr(PZD01,1,3),PZD10) Z" & _
'            " Where bt06 Is Null and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
'            " and bt08=PZD01(+) " & m_ConSql2
'   strSql = "select sum(tt) from (" & strSql & ")"
   
   strSql = "select bt03,bt08,bt12 from bairetrademark,tmbulletinnp" & _
            " Where bt06 Is Null and not(bt07 is null and bt07='') and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
            m_ConSql & _
            " group by bt03,bt08,bt12"
   strSql = strSql & " union all " & _
            "select bt03,decode(cu30,NULL,bt08,cu30),bt12 from bairetrademark,tmbulletinnp,customer" & _
            " Where bt06 Is Null and not(bt07 is null and bt07='') and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
            " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) " & m_ConSql2 & _
            " group by bt03,decode(cu30,NULL,bt08,cu30),bt12"
   strSql = "select count(bt03) from (" & strSql & ")"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields(0)) Then
         If Val(rsTmp.Fields(0)) > 0 Then
            QueryData = True
            txtCnt.Text = rsTmp.Fields(0)
         End If
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
   If QueryData = False Or Val(txtCnt) = 0 Then
      MsgBox "無資料！", vbOKOnly, Me.Caption & "列印"
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm020323 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   CloseIme
   TextInverse Text1(Index)
End Sub

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
               strMsg = "請輸入正確的日期"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               Text1_GotFocus Index
            End If
         End If
   End Select
End Sub

Private Sub txtPath1_GotFocus()
   TextInverse txtPath1
End Sub

Private Sub txtPath2_GotFocus()
   TextInverse txtPath2
End Sub

Private Sub cmdWord_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strTime As String, strSubject As String
Dim fs As Object
Dim i As Integer, i_Emp As Integer
'Dim strCU01 As String, strCU80 As String, strTempAppAddr As String, strTempAppAddrZip As String
Dim intCount As Integer, intRunCnt As Integer, strEmp As String
Dim varTmp
'Dim strST06 As String
'Dim strEmp1 As String, strEmp2 As String
Dim strFiles As String
Dim kk As Integer
Dim strCon As String
Dim ff1 As Integer
   
On Error GoTo ErrHnd
   
   strTime = time()
   
   'Add By Sindy 2019/5/9
   If Val(txtCnt) = 0 Then
      MsgBox "無開拓資料！", vbExclamation
      Exit Sub
   End If
   
   If txtPath2.Text = "" Then
      MsgBox "商標圖檔案路徑不可空白！", vbExclamation
      txtPath2.SetFocus
      Exit Sub
   End If
'   If InStr(txtPath2, ".") > 0 Then
'      For i = Len(txtPath2) To 1 Step -1
'         If Mid(txtPath2, i, 1) = "\" Then
'            txtPath2 = Mid(txtPath2, 1, i - 1)
'            Exit For
'         End If
'      Next i
'   Else
      If Right(txtPath2, 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
'   End If
   File2.Refresh
   If File2.ListCount = 0 Then
      MsgBox "找不到商標圖檔！"
      Exit Sub
   End If
   
   ChDir App.path 'Add By Sindy 2020/2/18 釋放資料夾權限
'   If text1(2).Text = "" Then
'      MsgBox "分機不可空白！", vbExclamation
'      text1(2).SetFocus
'      Exit Sub
'   End If
   
   'Modify By Sindy 2019/3/29
   '檢查是否有”沒有商標圖檔名”的資料
   If Me.Tag = "北所" Then
      strCon = " and (bt13 is null or bt13='1')"
   ElseIf Me.Tag = "中所" Then
      strCon = " and bt13='2'"
   ElseIf Me.Tag = "南所" Then
      strCon = " and bt13='3'"
   ElseIf Me.Tag = "高所" Then
      strCon = " and bt13='4'"
   End If
   strExc(0) = "SELECT * FROM bairetrademark,tmbulletinnp" & _
               " WHERE bt06 Is Null and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
               " and (bt07 is null or bt07='') and not(bt08 is null and bt13 is null)" & _
               " and (BT14 IS NULL OR BT14<>'對造') and instr(bt03,'?')=0" & _
               " and instr(bt04,'?')=0" & strCon
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         'Modify By Sindy 2019/3/29 產生無代表圖的清單
         If ff1 > 0 Then Close #ff1
         ff1 = FreeFile
         Open PUB_Getdesktop & "\" & strSrvDate(1) & "(" & Me.Tag & ")台灣商標延展無代表圖清單.txt" For Output As ff1
         Print #ff1, "備註：改字型Fixedsys標準11號字以直式上下左右各10MM列印"
         Print #ff1, "審定號數   商標名稱"
         Print #ff1, "========== ========================================================"
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            Print #ff1, convForm(CheckStr(RsTemp.Fields("BT01")), 10) & " " & RsTemp.Fields("BT02")
            RsTemp.MoveNext
         Loop
         Close ff1
         '2019/3/29 END
         If MsgBox("<沒有商標圖檔名> 的資料,共計 " & RsTemp.RecordCount & " 筆！" & vbCrLf & vbCrLf & _
                   "清單放置:" & PUB_Getdesktop & "\" & strSrvDate(1) & "(" & Me.Tag & ")台灣商標延展無代表圖清單.txt" & vbCrLf & vbCrLf & _
                   "是否還要繼續產生開拓函？", vbYesNo + vbQuestion + vbDefaultButton2, "資料有問題") = vbNo Then
            Exit Sub
         End If
      End If
   End If
      
   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.DeleteFolder m_WordFilePath, True
NotFolder76:
   fs.CreateFolder m_WordFilePath
   
   pFtpSrv = Pub_GetSpecMan("FTP_TM31")
   hConnection = PUB_GetFtpConnect(, , , pFtpSrv)
'   '代表圖 Test
'   '判斷word是否已開啟
'   If Not m_WordAp Is Nothing Then Set m_WordAp = Nothing
'   Set m_WordAp = New Word.Application
'   m_WordAp.Visible = True 'False
'   m_WordAp.Documents.Open "C:\20203-1.doc"
   
   m_intFileCnt = 0 'Add By Sindy 2019/10/4
   For kk = 1 To 3 '2
      intCount = 0
      intRunCnt = 1 '人數
      
      '非本所案件，排除國內商標公報特定公司不列印者，排除國外客戶
      'Modify By Sindy 2019/9/23 + 列印時, 依申請人名稱+郵遞區號
      If kk = 1 Then
'         strSql = "select distinct bt03 from bairetrademark,tmbulletinnp,(select substr(PZD01,1,3) PZD01,PZD10 from postzipdata group by substr(PZD01,1,3),PZD10) Z" & _
'                  " Where bt06 Is Null and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
'                  " and bt08=PZD01(+) " & m_ConSql & _
'                  " order by bt03 asc"
         strSql = "select bt03,bt08 from bairetrademark,tmbulletinnp" & _
                  " Where bt06 Is Null and not(bt07 is null and bt07='') and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
                  m_ConSql & _
                  " group by bt03,bt08" & _
                  " order by bt03 asc"
      ElseIf kk = 2 Then
         '要寄開拓函的客戶 無郵遞區號 或 為所內智權人員的客戶-非個人
'         strSql = "select distinct bt03,bt12 from bairetrademark,tmbulletinnp,(select substr(PZD01,1,3) PZD01,PZD10 from postzipdata group by substr(PZD01,1,3),PZD10) Z" & _
'                  " Where bt06 Is Null and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
'                  " and bt08=PZD01(+) " & m_ConSql2 & _
'                  " order by bt12 asc, bt03 asc"
         strSql = "select bt03,decode(cu30,NULL,bt08,cu30) bt08,bt12 from bairetrademark,tmbulletinnp,customer" & _
                  " Where bt06 Is Null and not(bt07 is null and bt07='') and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
                  " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) and cu15<>'0'" & m_ConSql2 & _
                  " group by bt03,decode(cu30,NULL,bt08,cu30),bt12" & _
                  " union all select bt03,decode(cu30,NULL,bt08,cu30) bt08,bt12 from bairetrademark,tmbulletinnp,customer" & _
                  " Where bt06 Is Null and not(bt07 is null and bt07='') and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
                  " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) AND cu01 is null AND bt14<>'0'" & m_ConSql2 & _
                  " group by bt03,decode(cu30,NULL,bt08,cu30),bt12" & _
                  " order by bt12 asc,bt03 asc"
      'Add By Sindy 2020/3/10
      ElseIf kk = 3 Then
         '要寄開拓函的客戶 無郵遞區號 或 為所內智權人員的客戶-個人
         strSql = "select bt03,decode(cu30,NULL,bt08,cu30) bt08,bt12 from bairetrademark,tmbulletinnp,customer" & _
                  " Where bt06 Is Null and not(bt07 is null and bt07='') and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
                  " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) and cu15='0'" & m_ConSql2 & _
                  " group by bt03,decode(cu30,NULL,bt08,cu30),bt12" & _
                  " union all select bt03,decode(cu30,NULL,bt08,cu30) bt08,bt12 from bairetrademark,tmbulletinnp,customer" & _
                  " Where bt06 Is Null and not(bt07 is null and bt07='') and ltrim(rtrim(bt03))=tbnp01(+) and tbnp08(+)='T' and tbnp01 is null" & _
                  " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) AND cu01 is null AND bt14='0'" & m_ConSql2 & _
                  " group by bt03,decode(cu30,NULL,bt08,cu30),bt12" & _
                  " order by bt12 asc,bt03 asc"
      '2020/3/10 END
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         If kk = 1 Then
            If Me.Tag = "北所" Then
               varTmp = Split(Pub_GetSpecMan("台灣延展北所開拓人員"), ";")
               intRunCnt = UBound(varTmp) + 1 '人數
               intCount = Round(Val(rsTmp.RecordCount) / intRunCnt) '分配到的筆數(四捨五入)
            End If
         End If
         
         '產生Word檔
         Screen.MousePointer = vbHourglass
         rsTmp.MoveFirst
         'm_intFileCnt = 0 'Modify By Sindy 2019/10/4 Mark
         bolRetry = True
         If intCount > 0 Then
            m_intFileCnt = 0 'Modify By Sindy 2019/10/4
            i_Emp = 0
            strEmp = varTmp(i_Emp) '北所開拓人員
         Else
            strEmp = "" '分所
         End If
         For m_iRow = 1 To rsTmp.RecordCount '筆數
            '一舜科技股 / " & rsTmp.Fields(0) & " / 七寶旅行社股份有限公司
            'Modify By Sindy 改以專用權人的審定號數小到大排序,抓最小號數的資料
            'strSql = "select * from bairetrademark Where bt06 Is Null and bt03='七寶旅行社股份有限公司' order by bt01 asc"
            'Modify By Sindy 2012/10/2 龔說要以專用權人+商標種類(商標前服務標章後)+審定號數小到大排序,抓最小號數的資料
            strSql = "select bairetrademark.*,substr(bt07,1,1) as T1,cu30,cu31" & _
                     " from bairetrademark,customer" & _
                     " Where bt06 Is Null and bt03='" & rsTmp.Fields(0) & "'" & _
                     " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) " & _
                     " AND (decode(cu30,NULL,bt08,cu30)='" & "" & rsTmp.Fields("bt08") & "' or decode(cu30,NULL,bt08,cu30)='" & PUB_ChangeZIPToSir("" & rsTmp.Fields("bt08")) & "')" & _
                     " order by BT05 asc,BT01 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               m_AppName = "" & RsTemp.Fields("bt03") '專用權人
               'Modify By Sindy 2020/3/10 + And kk < 3 因為客戶檔-個人都抓郵遞區號和公報地址
               If "" & RsTemp.Fields("cu31") <> "" And kk < 3 Then '有聯絡地址
                  m_AppAddr = "" & RsTemp.Fields("cu31") '專用權人地址
                  m_AppAddrZip = PUB_ChangeZIPToSir("" & RsTemp.Fields("cu30")) '郵遞區號
                  m_AppAddrZipOld = "" & RsTemp.Fields("cu30")
               Else
                  m_AppAddr = "" & RsTemp.Fields("bt04") '專用權人地址
                  m_AppAddrZip = PUB_ChangeZIPToSir("" & RsTemp.Fields("bt08")) '郵遞區號
                  m_AppAddrZipOld = "" & RsTemp.Fields("bt08")
               End If
               m_ApplSales = "" & RsTemp.Fields("bt12") '智權人員
               m_ApplSalesST06 = "" & RsTemp.Fields("bt13") '所別
               
'               '非個人, 抓相同客戶名稱之最小客戶編號,但抓出之編號若CU02<>'0',則抓CU02='0'的聯絡地址
'               strCU01 = "": strTempAppAddr = "": strTempAppAddrZip = ""
'               m_ApplSales = "": m_ApplSalesST06 = "" 'Add By Sindy 2019/2/1
'               If Len(m_AppName) > 4 Then
'                  strSql = "SELECT * FROM customer WHERE cu04='" & m_AppName & "' order by cu01 asc,cu02 asc"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     If RsTemp.RecordCount > 0 Then
'                        RsTemp.MoveFirst
'                        If RsTemp.Fields("cu02") <> "0" Then
'                           strCU01 = "" & RsTemp.Fields("cu01")
'                        Else
'                           'Modify By Sindy 2019/3/20
'                           strCU80 = Trim("" & RsTemp.Fields("CU80"))
'                           If strCU80 <> "" Then
'                              If strCU80 <> "刪址" And _
'                                 strCU80 <> "遷移不明" And _
'                                 strCU80 <> "其他" And _
'                                 strCU80 <> "業務自行處理" Then
'                                 GoTo ReadNext '上列客戶狀態用公報地址,其他的不產生開拓函
'                              End If
'                           Else
'                           '2019/3/20 END
'                              strTempAppAddrZip = Trim("" & RsTemp.Fields("cu30"))
'                              strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
'                              m_ApplSales = GetPrjSalesNM_2("" & RsTemp.Fields("cu13"), , m_ApplSalesST06, 1) 'Add By Sindy 2019/2/1
'                           End If
'                        End If
'                     End If
'                  End If
'                  If strCU01 <> "" Then
'                     strSql = "SELECT * FROM customer WHERE cu01='" & strCU01 & "' and cu02='0'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 Then
'                        'Modify By Sindy 2019/3/20
'                        strCU80 = Trim("" & RsTemp.Fields("CU80"))
'                        If strCU80 <> "" Then
'                           If strCU80 <> "刪址" And _
'                              strCU80 <> "遷移不明" And _
'                              strCU80 <> "其他" And _
'                              strCU80 <> "業務自行處理" Then
'                              GoTo ReadNext '上列客戶狀態用公報地址,其他的不產生開拓函
'                           End If
'                        Else
'                        '2019/3/20 END
'                           strTempAppAddrZip = Trim("" & RsTemp.Fields("cu30"))
'                           strTempAppAddr = Trim("" & RsTemp.Fields("cu31"))
'                           m_ApplSales = GetPrjSalesNM_2("" & RsTemp.Fields("cu13"), , m_ApplSalesST06, 1) 'Add By Sindy 2019/2/1
'                        End If
'                     End If
'                  End If
'                  If strTempAppAddr <> "" Then
'                     m_AppAddrZip = PUB_ChangeZIPToSir(strTempAppAddrZip)
'                     m_AppAddr = strTempAppAddr
'                  End If
'               End If
               
               'Add By Sindy 2013/6/3
               If m_AppAddrZip = "" Then
                  m_AppAddrZip = PUB_ChangeZIPToSir(Left(PUB_AddrChangeZIPCode(m_AppAddr, , False), 3))
               End If
               '2013/6/3 End
               
               '列印定稿
               Forms(0).tmrConnect.Tag = 0 '所外透過VPN,進所內操作之故。不斷線 Add By Sindy 2020/4/14
               'Modify By Sindy 2020/3/10
               If kk = 3 Then '台一客戶-個人
                  If WordEdit2() = False Then
                     GoTo ErrHnd
                  End If
               Else
               '2020/3/10 END
                  If WordEdit(strEmp) = False Then
                     GoTo ErrHnd
                  End If
               End If
            End If
ReadNext:
            If intCount > 0 Then
               '北所,依件數平均分配給開拓的人員,超過100一樣切一個檔案
               If ((m_iRow - (i_Emp * intCount)) Mod 100) = 0 Or _
                  m_iRow = ((i_Emp + 1) * intCount) Or _
                  m_iRow = rsTmp.RecordCount Then
                  
                  g_WordAp.Documents.Save
                  g_WordAp.Documents.Close
                  bolRetry = True
                  'Modify By Sindy 2019/2/18 + And i_Emp < UBound(varTmp)
                  If m_iRow = ((i_Emp + 1) * intCount) And i_Emp < UBound(varTmp) Then  '換下一位開拓人員
                     m_intFileCnt = 0
                     i_Emp = i_Emp + 1
                     strEmp = varTmp(i_Emp) '北所開拓人員
                  End If
               End If
            Else
               '分所,100個申請人切一個檔案
               If (m_iRow Mod 100) = 0 Or m_iRow = rsTmp.RecordCount Then
                  g_WordAp.Documents.Save
                  g_WordAp.Documents.Close
                  bolRetry = True
               End If
            End If
            
            rsTmp.MoveNext
         Next m_iRow
         If bolRetry = False Then
            g_WordAp.Documents.Save
            g_WordAp.Documents.Close
      '      g_WordAp.Visible = True
      '      g_WordAp.WindowState = wdWindowStateMaximize
         End If
         Set g_WordAp = Nothing
         rsTmp.Close
   '      '代表圖 Test
   '      m_WordAp.Documents.Close
   '      Set m_WordAp = Nothing
      End If
   Next kk
   
   '檢查附件區
   File2.path = m_WordFilePath
   File2.Refresh
   If File2.ListCount > 0 Then
      For i = 0 To File2.ListCount - 1
         strFiles = strFiles & "*" & m_WordFilePath & "\" & File2.List(i)
      Next i
   End If
   If strFiles <> "" Then strFiles = Mid(strFiles, 2)
   '通知開拓函電子檔已產生完畢
   If strFiles <> "" Then
      Screen.MousePointer = vbDefault
      MsgBox "作業完成！請至 " & m_WordFilePath & " 資料夾中列印開拓函。（花費時間：" & strTime & "  " & time() & "）"
   End If
   
   Screen.MousePointer = vbDefault
   
   ChDir App.path 'Add By Sindy 2020/3/10 釋放資料夾權限
   Set rsTmp = Nothing
   Exit Sub
   
ErrHnd:
   If Err.Number = 76 Then
      GoTo NotFolder76
   ElseIf Err.Number = 70 Then
      MsgBox Err.Description & vbCrLf & _
         "（請將 Word 和 檔案總管 關閉，再執行！）", vbCritical
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
      '通知程序開拓函電子檔產生有誤
'      If strP22 <> "" Then
         strSubject = Me.Caption & "，電子檔產生有誤！"
         'PUB_SendMail strUserNum, strUserNum, "", strSubject, strSubject, , , , , , , , , , , False
'      End If
   End If
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Set g_WordAp = Nothing
   'Resume
End Sub

Private Function WordEdit(strEmp As String) As Boolean
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
Dim i As Integer, k As Integer
Dim strFAX As String
Dim strNo As String, strNote As String 'Add By Sindy 2019/3/21
Dim strManyTData As String 'Add By Sindy 2019/9/24
   
On Error GoTo ERRORSECTION1
   
   WordEdit = True
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize 'wdWindowStateMinimize  wdWindowStateMaximize
   With g_WordAp
   
      If bolRetry = True Then
         m_intFileCnt = m_intFileCnt + 1
         g_WordAp.Documents.add.SaveAs m_WordFilePath & "\台灣商標延展開拓函(智慧局)-" & IIf(strEmp <> "", GetPrjSalesNM(strEmp), "") & Format(m_intFileCnt, "00") & ".doc"
      End If
      
      If bolRetry = False Then .Selection.InsertBreak Type:=wdPageBreak
      
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5) '2
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(0.75) '2.5
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
'         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
'            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'            oShape.ZOrder 4
'            oShape.LockAnchor = True
'            oShape.LockAspectRatio = -1
'            oShape.Width = .CentimetersToPoints(21)
'            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'            oShape.Left = .CentimetersToPoints(0)
'            'oShape.Top = .CentimetersToPoints(27.3)
'            oShape.Top = .CentimetersToPoints(27)
'            oShape.WrapFormat.Type = wdWrapSquare
'            oShape.WrapFormat.Side = wdWrapBoth
'         End If
         .Selection.EndKey Unit:=wdStory
      End If
      
      '配合新的開窗定稿改固定行高
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行高
      .Selection.ParagraphFormat.LineSpacing = 15 '行高
      
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
      If m_AppAddrZip = "" Then
         .Selection.TypeParagraph
      End If
      .Selection.TypeText getAddrData
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "致：" & m_AppName
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
      .Selection.TypeText "敬啟者："
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      'Modify By Sindy 改以專用權人的審定號數小到大排序,抓最小號數的資料
      'Modify By Sindy 2012/10/2 龔說要以專用權人+商標種類(商標前服務標章後)+審定號數小到大排序,抓最小號數的資料
      'strSql = "select bairetrademark.*,substr(bt07,1,1) as T1 from bairetrademark Where bt06 Is Null and bt03='" & m_AppName & "' order by T1 desc,bt01 asc"
      strSql = "select bairetrademark.*,substr(bt07,1,1) as T1,cu30,cu31" & _
               " from bairetrademark,customer" & _
               " Where bt06 Is Null and bt03='" & m_AppName & "'" & _
               " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) " & _
               " AND decode(cu30,NULL,bt08,cu30)='" & m_AppAddrZipOld & "'" & _
               " order by BT05 asc,BT01 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         intHeight = 0 ': intCnt = 0
         For i = 1 To rsTmp.RecordCount
            If i = 1 Then
               'Add By Sindy 2019/9/24 串多個註冊號資訊
               rsTmp.MoveFirst
               strManyTData = ""
               Do While Not rsTmp.EOF
                  If strManyTData = "" Then
                     strManyTData = "第" & rsTmp.Fields("bt01") & "號「" & rsTmp.Fields("bt02") & "」"
                  Else
                     strManyTData = strManyTData & "、第" & rsTmp.Fields("bt01") & "號「" & rsTmp.Fields("bt02") & "」"
                  End If
                  rsTmp.MoveNext
               Loop
               rsTmp.MoveFirst
               
               strNo = rsTmp.Fields("bt01") 'Add By Sindy 2019/3/21
               '.Selection.TypeText "　　貴公司／台 端所註冊第" & rsTmp.Fields("bt01") & "號「" & rsTmp.Fields("bt02") & "」商標之專用期限將於民國" & Left(ChangeWStringToTString(rsTmp.Fields("bt05")), 3) & "年" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 4, 2) & "月" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 6, 2) & "日屆滿（商標資訊如下），若逾期未辦理延展，商標權將當然消滅，　貴公司／台 端若有意繼續使用前揭商標，即應辦理商標延展註冊，故請儘速與本所服務人員聯繫！"
               .Selection.TypeText "　　貴公司／台 端所註冊" & strManyTData & "商標之專用期限將於民國" & Left(ChangeWStringToTString(rsTmp.Fields("bt05")), 3) & "年" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 4, 2) & "月" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 6, 2) & "日屆滿（商標資訊如下），若逾期未辦理延展，商標權將當然消滅，　貴公司／台 端若有意繼續使用前揭商標，即應辦理商標延展註冊，故請儘速與本所服務人員聯繫！"
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
               '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　台一國際專利商標事務所"
               .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　" & PUB_GetCompName2("1")
               'end 2020/3/30
               .Selection.TypeParagraph
               'Modify By Sindy 2019/2/1
               If m_ApplSales <> "" And m_ApplSalesST06 <> "" Then
                  strNote = GetPrjSalesNM(m_ApplSales) 'Add By Sindy 2019/3/21
                  If m_ApplSalesST06 = "1" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台北所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　電話:(02)2506-1023 分機" & Pub_GetStaffExtn(m_ApplSales)
                  ElseIf m_ApplSalesST06 = "2" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台中所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　電話:(04)2327-0288 分機" & Pub_GetStaffExtn(m_ApplSales)
                  ElseIf m_ApplSalesST06 = "3" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台南所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　電話:(06)2743-866 分機" & Pub_GetStaffExtn(m_ApplSales)
                  ElseIf m_ApplSalesST06 = "4" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　高雄所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　電話:(07)2363-602 分機" & Pub_GetStaffExtn(m_ApplSales)
                  End If
               Else
               '2019/2/1 END
                  If Me.Tag = "北所" Then
                     strNote = GetPrjSalesNM(strEmp) 'Add By Sindy 2019/3/21
                     .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台北所　" & strNote & "敬上"
                  ElseIf Me.Tag = "中所" Then
                     'strNote = "陳家欣" 'Add By Sindy 2019/3/21
                     'strNote = GetPrjSalesNM("A7016")
                     strNote = GetPrjSalesNM(Pub_GetSpecMan("台灣商標開拓中所人員")) 'Modify By Sindy 2020/9/30
                     .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台中所　" & strNote & "敬上"
                  ElseIf Me.Tag = "南所" Then
                     'strNote = "鄭鈺華" 'Add By Sindy 2019/3/21
                     'strNote = GetPrjSalesNM("A5005")
                     'Modify By Sindy 2025/5/6
                     strNote = GetPrjSalesNM(Pub_GetSpecMan("台灣商標開拓南所人員")) 'Modify By Sindy 2020/9/30
                     If strNote <> "" Then
                        .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台南所　" & strNote & "敬上"
                     Else
                        .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　" & Pub_GetSpecMan("台灣商標開拓南所人員") & "　敬上"
                     End If
                     '2025/5/6 END
                  ElseIf Me.Tag = "高所" Then
                     'strNote = "謝秀珠" 'Add By Sindy 2019/3/21
                     'strNote = GetPrjSalesNM("89047")
                     strNote = GetPrjSalesNM(Pub_GetSpecMan("台灣商標開拓高所人員")) 'Modify By Sindy 2020/9/30
                     .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　高雄所　" & strNote & "敬上"
                  End If
                  .Selection.TypeParagraph
                  If Me.Tag = "北所" Then
                     .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　電話:(02)2506-1023 分機" & Pub_GetStaffExtn(strEmp)
                  ElseIf Me.Tag = "中所" Then
                     .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　電話:(04)2327-0288 分機" & Pub_GetStaffExtn(strNote)
                  ElseIf Me.Tag = "南所" Then
                     .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　電話:(06)2743-866"
                  ElseIf Me.Tag = "高所" Then
                     .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　電話:(07)2363-602 分機" & Pub_GetStaffExtn(strNote)
                  End If
               End If
               .Selection.ParagraphFormat.SpaceAfter = 6 '與後段距離
               .Selection.TypeParagraph
               '插入表格(無框線)
               Dim oTable
               Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=rsTmp.RecordCount, NumColumns:=1)
               oTable.AllowAutoFit = True
'               .ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=3, NumColumns:= _
'        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
'        wdAutoFitFixed
               For k = 1 To rsTmp.RecordCount
                  .Selection.TypeText "***代表圖***"
                  If k = rsTmp.RecordCount Then
                     .Selection.MoveDown Unit:=wdLine, Count:=1
                  Else
                     .Selection.MoveRight Unit:=wdCell
                  End If
               Next k
               '.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
               '.Selection.TypeParagraph
            End If
            
''代表圖 Test
'Call WordFindText2(m_WordAp, "6206846")
'Call WordFindText(g_WordAp, "***代表圖***", "複製圖片")
            
            Call WordFindText(g_WordAp, "***代表圖***")
            '配合新的開窗定稿改固定行高
            '.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行高
            .Selection.ParagraphFormat.LineSpacing = 5 '行高
            .Selection.ParagraphFormat.SpaceAfter = 0 '與後段距離
            If "" & rsTmp.Fields("bt07") <> "" Then
               AddInPicToWordR g_WordAp, rsTmp.Fields("bt07") & ".gif", i '插入圖檔
            End If
            rsTmp.MoveNext
         Next i
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.ParagraphFormat.SpaceAfter = 0 '與後段距離
         '.Selection.TypeParagraph
         .Selection.Font.Size = 12
         .Selection.TypeText "說明："
         .Selection.TypeParagraph
         .Selection.TypeText "　　1.依商標法第三十四條規定：商標權之延展，應於商標權屆滿前六個月內提出申請，並繳納延展註冊費；其於商標權屆滿後六個月內提出申請者，應繳納二倍延展註冊費。"
         .Selection.TypeParagraph
         .Selection.TypeText "　　2.所需文件：委任書(本所準備)。"
         .Selection.TypeParagraph
         .Selection.TypeText "==============================《回覆單》=================================="
         .Selection.TypeParagraph
         .Selection.TypeText "□同意　貴所派員聯繫，共商本案後續處理事宜。"
         .Selection.TypeParagraph
         .Selection.TypeText "□本人／本公司自行處理本案之後續作業，請 貴所無須對本案進行後續追蹤及通知。"
         .Selection.TypeParagraph
         .Selection.TypeText "□放棄延展。"
         .Selection.TypeParagraph
         'Add By Sindy 2019/3/21
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
         .Selection.TypeText "註冊號：" & strNo & "　　" & strNote
         .Selection.TypeParagraph
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
         '2019/3/21 END
         .Selection.TypeText "回 覆 者：_______________________(簽章)   回覆日期：    年    月    日"
         .Selection.TypeParagraph
         .Selection.TypeText "聯 絡 方 式：電話____________________    傳真：____________________"
         .Selection.TypeParagraph
         .Selection.Font.Size = 11
         strFAX = "02-25011666" '北所
         If (m_ApplSales <> "" And m_ApplSalesST06 = "2") Or Me.Tag = "中所" Then
            strFAX = "04-23227483"
         ElseIf (m_ApplSales <> "" And m_ApplSalesST06 = "3") Or Me.Tag = "南所" Then
            strFAX = "06-2744030"
         ElseIf (m_ApplSales <> "" And m_ApplSalesST06 = "4") Or Me.Tag = "高所" Then
            strFAX = "07-2364360"
         End If
         .Selection.TypeText "※請於框內勾選填妥後，傳真通知本所(傳真號碼：" & strFAX & ")或來電告知，感謝您的合作!"
         .Selection.TypeParagraph
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

'Add By Sindy 2020/3/10 + 台一客戶-個人
Private Function WordEdit2() As Boolean
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
Dim i As Integer, k As Integer
Dim strFAX As String
Dim strNo As String, strNote As String 'Add By Sindy 2019/3/21
Dim strManyTData As String 'Add By Sindy 2019/9/24
Dim intCnt As Integer
   
On Error GoTo ERRORSECTION1
   
   WordEdit2 = True
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize 'wdWindowStateMinimize  wdWindowStateMaximize
   With g_WordAp
   
      If bolRetry = True Then
         m_intFileCnt = m_intFileCnt + 1
         'g_WordAp.Documents.add.SaveAs m_WordFilePath & "\台灣商標延展開拓函(智慧局)-個人" & Format(m_intFileCnt, "00") & ".doc"
         g_WordAp.Documents.add.SaveAs m_WordFilePath & "\台灣商標延展開拓函(智慧局)-個人.doc"
      End If
      
      'Modify By Sindy 改以專用權人的審定號數小到大排序,抓最小號數的資料
      'Modify By Sindy 2012/10/2 龔說要以專用權人+商標種類(商標前服務標章後)+審定號數小到大排序,抓最小號數的資料
      'strSql = "select bairetrademark.*,substr(bt07,1,1) as T1 from bairetrademark Where bt06 Is Null and bt03='" & m_AppName & "' order by T1 desc,bt01 asc"
      strSql = "select bairetrademark.*,substr(bt07,1,1) as T1,cu30,cu31" & _
               " from bairetrademark,customer" & _
               " Where bt06 Is Null and bt03='" & m_AppName & "'" & _
               " AND substr(bt14,1,8)=cu01(+) and substr(bt14,9,1)=cu02(+) " & _
               " AND decode(cu30,NULL,bt08,cu30)='" & m_AppAddrZipOld & "'" & _
               " order by BT05 asc,BT01 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         intHeight = 0 ': intCnt = 0
'         For i = 1 To rsTmp.RecordCount
'            If i = 1 Then
         'rsTmp.MoveFirst
         Do While Not rsTmp.EOF
               If bolRetry = False Then .Selection.InsertBreak Type:=wdPageBreak
               
               .Selection.Font.Name = "標楷體"
               .Selection.PageSetup.Orientation = wdOrientPortrait
               .Selection.Orientation = wdTextOrientationHorizontal
               .Selection.Font.Size = 14
               .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
               .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5) '2
               .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
               .Selection.PageSetup.BottomMargin = .CentimetersToPoints(0.75) '2.5
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
         '         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
         '            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         '            oShape.ZOrder 4
         '            oShape.LockAnchor = True
         '            oShape.LockAspectRatio = -1
         '            oShape.Width = .CentimetersToPoints(21)
         '            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         '            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         '            oShape.Left = .CentimetersToPoints(0)
         '            'oShape.Top = .CentimetersToPoints(27.3)
         '            oShape.Top = .CentimetersToPoints(27)
         '            oShape.WrapFormat.Type = wdWrapSquare
         '            oShape.WrapFormat.Side = wdWrapBoth
         '         End If
                  .Selection.EndKey Unit:=wdStory
               End If
               
               '配合新的開窗定稿改固定行高
               .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行高
               .Selection.ParagraphFormat.LineSpacing = 15 '行高
               
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
               If m_AppAddrZip = "" Then
                  .Selection.TypeParagraph
               End If
               .Selection.TypeText getAddrData
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
         '      .Selection.TypeParagraph
         '      .Selection.TypeText "致：" & m_AppName
         '      .Selection.TypeParagraph
         '      .Selection.TypeParagraph
               .Selection.TypeText "敬啟者："
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               
               strNo = rsTmp.Fields("bt01") 'Add By Sindy 2019/3/21
               '.Selection.TypeText "　　貴公司／台 端所註冊第" & rsTmp.Fields("bt01") & "號「" & rsTmp.Fields("bt02") & "」商標之專用期限將於民國" & Left(ChangeWStringToTString(rsTmp.Fields("bt05")), 3) & "年" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 4, 2) & "月" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 6, 2) & "日屆滿（商標資訊如下），若逾期未辦理延展，商標權將當然消滅，　貴公司／台 端若有意繼續使用前揭商標，即應辦理商標延展註冊，故請儘速與本所服務人員聯繫！"
               .Selection.TypeText "　　貴公司／台 端所註冊第" & rsTmp.Fields("bt01") & "號「" & rsTmp.Fields("bt02") & "」商標之專用期限將於民國" & Left(ChangeWStringToTString(rsTmp.Fields("bt05")), 3) & "年" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 4, 2) & "月" & Mid(ChangeWStringToTString(rsTmp.Fields("bt05")), 6, 2) & "日屆滿（商標資訊如下），若逾期未辦理延展，商標權將當然消滅，　貴公司／台 端若有意繼續使用前揭商標，即應辦理商標延展註冊，故請儘速與本所服務人員聯繫！"
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
               '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　台一國際專利商標事務所"
               .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　" & PUB_GetCompName2("1")
               'end 2020/3/30
               .Selection.TypeParagraph
               
                  strNote = GetPrjSalesNM(m_ApplSales) 'Add By Sindy 2019/3/21
                  If m_ApplSalesST06 = "1" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台北所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　電話:(02)2506-1023 分機" & Pub_GetStaffExtn(m_ApplSales)
                  ElseIf m_ApplSalesST06 = "2" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台中所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　電話:(04)2327-0288 分機" & Pub_GetStaffExtn(m_ApplSales)
                  ElseIf m_ApplSalesST06 = "3" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　台南所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　電話:(06)2743-866 分機" & Pub_GetStaffExtn(m_ApplSales)
                  ElseIf m_ApplSalesST06 = "4" Then
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　高雄所　" & GetPrjSalesNM(m_ApplSales) & "敬上"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　電話:(07)2363-602 分機" & Pub_GetStaffExtn(m_ApplSales)
                  End If
               
               .Selection.ParagraphFormat.SpaceAfter = 6 '與後段距離
               .Selection.TypeParagraph
               '插入表格(無框線)
               Dim oTable
               Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=1)
               oTable.AllowAutoFit = True
'               .ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=3, NumColumns:= _
'        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
'        wdAutoFitFixed
'               For k = 1 To rsTmp.RecordCount
                  .Selection.TypeText "***代表圖***"
'                  If k = rsTmp.RecordCount Then
                     .Selection.MoveDown Unit:=wdLine, Count:=1
'                  Else
'                     .Selection.MoveRight Unit:=wdCell
'                  End If
'               Next k
               '.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
               '.Selection.TypeParagraph
            
''代表圖 Test
'Call WordFindText2(m_WordAp, "6206846")
'Call WordFindText(g_WordAp, "***代表圖***", "複製圖片")
            
            Call WordFindText(g_WordAp, "***代表圖***")
            '配合新的開窗定稿改固定行高
            '.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行高
            .Selection.ParagraphFormat.LineSpacing = 5 '行高
            .Selection.ParagraphFormat.SpaceAfter = 0 '與後段距離
            If "" & rsTmp.Fields("bt07") <> "" Then
               AddInPicToWordR g_WordAp, rsTmp.Fields("bt07") & ".gif", i '插入圖檔
            End If
'            rsTmp.MoveNext
'         Next i
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.ParagraphFormat.SpaceAfter = 0 '與後段距離
         '.Selection.TypeParagraph
         .Selection.Font.Size = 12
         .Selection.TypeText "說明："
            .Selection.TypeParagraph
            .Selection.TypeText "　　1.依商標法第三十四條規定：商標權之延展，應於商標權屆滿前六個月內提出申請，並繳納延展註冊費；其於商標權屆滿後六個月內提出申請者，應繳納二倍延展註冊費。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　2.所需文件：委任書(本所準備)。"
            .Selection.TypeParagraph
            .Selection.TypeText "==============================《回覆單》=================================="
            .Selection.TypeParagraph
            .Selection.TypeText "□同意　貴所派員聯繫，共商本案後續處理事宜。"
            .Selection.TypeParagraph
            .Selection.TypeText "□本人／本公司自行處理本案之後續作業，請 貴所無須對本案進行後續追蹤及通知。"
            .Selection.TypeParagraph
            .Selection.TypeText "□放棄延展。"
            .Selection.TypeParagraph
            'Add By Sindy 2019/3/21
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
            .Selection.TypeText "註冊號：" & strNo & "　　" & strNote
            .Selection.TypeParagraph
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
            '2019/3/21 END
            .Selection.TypeText "回 覆 者：_______________________(簽章)   回覆日期：    年    月    日"
            .Selection.TypeParagraph
            .Selection.TypeText "聯 絡 方 式：電話____________________    傳真：____________________"
            .Selection.TypeParagraph
            .Selection.Font.Size = 11
            strFAX = "02-25011666" '北所
            If (m_ApplSales <> "" And m_ApplSalesST06 = "2") Or Me.Tag = "中所" Then
               strFAX = "04-23227483"
            ElseIf (m_ApplSales <> "" And m_ApplSalesST06 = "3") Or Me.Tag = "南所" Then
               strFAX = "06-2744030"
            ElseIf (m_ApplSales <> "" And m_ApplSalesST06 = "4") Or Me.Tag = "高所" Then
               strFAX = "07-2364360"
            End If
            .Selection.TypeText "※請於框內勾選填妥後，傳真通知本所(傳真號碼：" & strFAX & ")或來電告知，感謝您的合作!"
            .Selection.TypeParagraph
            
            rsTmp.MoveNext
         Loop
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
            WordEdit2 = False
      End Select
   End If
End Function

'尋找Word檔中文字:Find/清除/置換文字 或 貼上
Private Sub WordFindText(g_WordAp As Word.Application, strFindText As String, Optional strReplaceText As String = "")
Dim bolResult As Boolean
   
   If Trim(strFindText) = "" Then Exit Sub
   With g_WordAp
'      .Selection.WholeStory
'      .Selection.Copy
      .Selection.GoTo what:=wdGoToPage, which:=wdGoToPrevious, Count:=3
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = strFindText
      .Selection.Find.Replacement.Text = ""
      .Selection.Find.Forward = True
      .Selection.Find.Wrap = wdFindContinue
      .Selection.Find.Format = False
      .Selection.Find.MatchCase = False
      .Selection.Find.MatchWholeWord = False
      .Selection.Find.MatchWildcards = False
      .Selection.Find.MatchSoundsLike = False
      .Selection.Find.MatchAllWordForms = False
      .Selection.Find.MatchByte = True
      bolResult = .Selection.Find.Execute
      If bolResult = True Then
         .Selection.Delete
         If strReplaceText = "複製圖片" Then
            .Selection.Paste 'Format '(wdSingleCellText)
         Else
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strReplaceText
         End If
      End If
   End With
End Sub

'尋找Word檔中文字:Find/清除/複製
Private Sub WordFindText2(g_WordAp As Word.Application, strFindText As String)
Dim bolResult As Boolean
   
   If Trim(strFindText) = "" Then Exit Sub
   With g_WordAp
'      .Selection.WholeStory
'      .Selection.Copy
      .Selection.GoTo what:=wdGoToPage, which:=wdGoToPrevious, Count:=3
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = strFindText
      .Selection.Find.Replacement.Text = ""
      .Selection.Find.Forward = True
      .Selection.Find.Wrap = wdFindContinue
      .Selection.Find.Format = False
      .Selection.Find.MatchCase = False
      .Selection.Find.MatchWholeWord = False
      .Selection.Find.MatchWildcards = False
      .Selection.Find.MatchSoundsLike = False
      .Selection.Find.MatchAllWordForms = False
      .Selection.Find.MatchByte = True
      bolResult = .Selection.Find.Execute
      If bolResult = True Then
         .Selection.MoveRight Unit:=wdCell
         .Selection.MoveRight Unit:=wdCell
         .Selection.Copy
      End If
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

Private Sub AddInPicToWordR(ByRef oWord As Word.Application, strFileName As String, intFileCnt As Integer)
Dim oShape

   With oWord
'      '筆數為3的倍數時,接下一頁
'      'If (intFileCnt Mod 3) = 1 Then
'      If intHeight > 453 Or intHeight = 0 Then
'         intCnt = 1
'         intHeight = 5
'         .Selection.InsertBreak
'      Else
'         intCnt = intCnt + 1
'      End If
      
Dim strFromPathFile As String
Dim strToPathFile As String
Dim pData As WIN32_FIND_DATA
Dim hFind As Long
      
      '不可以含網路名稱
      'XFER為大寫才讀的到資料夾
      '  範例路徑: //XFER/BAIRETRADEMARK/T1726814.gif
      'Modified by Lydia 2024/07/22 改成變數
      'strFromPathFile = Replace("\\" & Replace(UCase(txtPath2), "\\SALE1\", "") & "\" & strFileName, "\", "/")
      strFromPathFile = Replace("\\" & Replace(UCase(txtPath2), "\\" & UCase(strSale1Path) & "\", "") & "\" & strFileName, "\", "/")
      strToPathFile = m_AttachPath & "\" & strFileName
'      pData.cFileName = String(MAX_PATH, 0)
'      hFind = FtpFindFirstFile(hConnection, strFromPathFile, pData, 0, 0)
'      If hFind <> 0 Then
         If PUB_FtpGetFile(strFromPathFile, strToPathFile, pFtpSrv) = False Then Exit Sub
         DoEvents 'Add By Sindy 2019/10/3
'      Else
'         Exit Sub
'      End If
      
      '插入圖片檔案
      '.ChangeFileOpenDirectory m_AttachPath & "\"
      'Add By Sindy 2012/10/17 檢查檔案是否存在
      If FileExists(Replace(strToPathFile, "/", "\")) = False Then Exit Sub
      DoEvents 'Add By Sindy 2019/10/3
      
      '插入圖片檔案
      'Modified by Lydia 2016/09/29 用舊寫法會造成Word2010出錯
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=strFileName, LinkToFile:=False, SaveWithDocument:=True
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:= _
      'strFileName, LinkToFile:= _
      'False, SaveWithDocument:=True
      '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
      Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=strToPathFile, LinkToFile:=False, SaveWithDocument:=True)
      oShape.Select
      DoEvents 'Add By Sindy 2019/10/3
      
      '定義大小
      '鎖定最高 圖區
      '圖大小
      'Modified by Lydia 2016/09/29
      '.Selection.ShapeRange.LockAspectRatio = msoTrue
      '.Selection.ShapeRange.Line.Visible = True '加框線
      'msoTrue：如果指定的圖案會保留原來調整。
      'msoFalse：如果您在您調整可以變更高度及寬度互不圖形。
      oShape.LockAspectRatio = 0 'msoTrue:-1
      oShape.Line.Visible = True '加框線
      oShape.WrapFormat.Type = wdWrapSquare
      oShape.WrapFormat.Side = wdWrapBoth
      
      '移到指定位置
      '3個圖檔高度:210
      '4個圖檔高度:155
      'Modified by Lydia 2016/09/29
      'If Selection.ShapeRange.Height > 200 Then
      '   Selection.ShapeRange.Height = 200 '210
      If oShape.Height > 200 Then
         oShape.Height = 200
      End If
'      If intFileCnt = 1 Then
'         intHeight = 5
'      Else
'         intHeight = intHeight + 6
'      End If
      'Modified by Lydia 2016/09/29
      '.Selection.ShapeRange.Top = intHeight
      '.Selection.ShapeRange.Left = 5 'Add By Sindy 2012/9/3
      'intHeight = intHeight + Selection.ShapeRange.Height
      oShape.Top = 0 'intHeight
      oShape.Left = 4
      'intHeight = intHeight + oShape.Height
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
      '其圖片配置為【文繞圖方式：(與文字排列),水平對齊方式(不可選)】
'      .Selection.ShapeRange.WrapFormat.Type = wdWrapSquare 'wdWrapNone.圖蓋文 wdWrapSquare.文字繞圖
'      .Selection.ShapeRange.WrapFormat.Side = wdWrapBoth
'      .Selection.ShapeRange.WrapFormat.DistanceTop = InchesToPoints(0)
'      .Selection.ShapeRange.WrapFormat.DistanceBottom = InchesToPoints(0.1)
'      .Selection.ShapeRange.WrapFormat.DistanceLeft = InchesToPoints(0.1)
'      .Selection.ShapeRange.WrapFormat.DistanceRight = InchesToPoints(0.1)
      
      .Selection.EndKey Unit:=wdStory
   End With
   Exit Sub
   
ErrHnd:
   Err.Raise Err.Number
End Sub
