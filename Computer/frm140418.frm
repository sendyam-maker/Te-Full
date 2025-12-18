VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140418 
   BorderStyle     =   1  '單線固定
   Caption         =   "整批匯入至往來記錄"
   ClientHeight    =   6020
   ClientLeft      =   40
   ClientTop       =   280
   ClientWidth     =   9080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6020
   ScaleWidth      =   9080
   Begin VB.Frame FrmExcel 
      Caption         =   "匯入Excel檔, 必須欄位"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1365
      Left            =   30
      TabIndex        =   25
      Top             =   30
      Visible         =   0   'False
      Width           =   6075
      Begin VB.ComboBox cboPlace 
         Height          =   300
         ItemData        =   "frm140418.frx":0000
         Left            =   720
         List            =   "frm140418.frx":0002
         TabIndex        =   29
         Top             =   960
         Width           =   5295
      End
      Begin VB.ComboBox cboSort 
         Height          =   300
         ItemData        =   "frm140418.frx":0004
         Left            =   1110
         List            =   "frm140418.frx":0006
         Style           =   2  '單純下拉式
         TabIndex        =   27
         Top             =   280
         Width           =   4920
      End
      Begin MSForms.TextBox txtCR06 
         Height          =   315
         Left            =   720
         TabIndex        =   28
         Top             =   630
         Width           =   5265
         VariousPropertyBits=   671105051
         Size            =   "9287;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "場合："
         Height          =   225
         Left            =   180
         TabIndex        =   31
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "主旨："
         Height          =   225
         Left            =   180
         TabIndex        =   30
         Top             =   690
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "往來類別："
         Height          =   225
         Left            =   180
         TabIndex        =   26
         Top             =   330
         Width           =   945
      End
   End
   Begin VB.CheckBox ChkExcel 
      Caption         =   "Check1"
      Height          =   180
      Left            =   6210
      TabIndex        =   32
      Top             =   150
      Width           =   255
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯入Excel檔(&E)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   345
      Left            =   6180
      TabIndex        =   24
      Top             =   60
      Width           =   1725
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5505
      Left            =   30
      TabIndex        =   8
      Top             =   450
      Width           =   9015
      _ExtentX        =   15893
      _ExtentY        =   9719
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "匯入資料"
      TabPicture(0)   =   "frm140418.frx":0008
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPath1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdImport"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrint"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "查詢資料"
      TabPicture(1)   =   "frm140418.frx":0024
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "GrdDataList"
      Tab(1).Control(3)=   "cmdQuery"
      Tab(1).Control(4)=   "txt1"
      Tab(1).Control(5)=   "cmdOK(9)"
      Tab(1).Control(6)=   "cmdOK(8)"
      Tab(1).Control(7)=   "cmdOK(0)"
      Tab(1).Control(8)=   "cmdOK(2)"
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdPrint 
         Caption         =   "列印(&P)"
         Height          =   345
         Left            =   7680
         TabIndex        =   33
         Top             =   750
         Width           =   885
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "關係企業"
         Height          =   345
         Index           =   2
         Left            =   -67320
         Style           =   1  '圖片外觀
         TabIndex        =   22
         Top             =   540
         Visible         =   0   'False
         Width           =   920
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "申請人"
         Height          =   345
         Index           =   0
         Left            =   -70260
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   930
         Width           =   740
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "往來記錄"
         Height          =   345
         Index           =   8
         Left            =   -67860
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   930
         Width           =   1170
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "接洽人/聯絡人"
         Height          =   345
         Index           =   9
         Left            =   -69465
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   930
         Width           =   1530
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Left            =   -73530
         TabIndex        =   1
         Top             =   420
         Width           =   2205
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   345
         Left            =   -71280
         TabIndex        =   2
         Top             =   390
         Width           =   885
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "匯入(&T)"
         Height          =   345
         Left            =   6660
         TabIndex        =   16
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox txtPath1 
         Height          =   315
         Left            =   1770
         TabIndex        =   0
         Top             =   420
         Width           =   6825
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<="
         Height          =   315
         Left            =   8610
         TabIndex        =   14
         Top             =   420
         Width           =   345
      End
      Begin VB.Frame Frame1 
         Height          =   465
         Left            =   60
         TabIndex        =   12
         Top             =   4980
         Width           =   8895
         Begin VB.TextBox TextCnt 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00FF0000&
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   30
            TabIndex        =   13
            Top             =   120
            Width           =   8820
         End
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
         Height          =   3855
         Left            =   90
         TabIndex        =   10
         Top             =   1020
         Width           =   8865
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  '沒有框線
            Height          =   210
            Left            =   1560
            TabIndex        =   23
            Text            =   "Text3"
            Top             =   0
            Width           =   975
         End
         Begin VB.ListBox List1 
            Height          =   3280
            ItemData        =   "frm140418.frx":0040
            Left            =   90
            List            =   "frm140418.frx":0042
            TabIndex        =   11
            Top             =   240
            Width           =   8685
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   1065
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frm140418.frx":0044
         Top             =   4050
         Visible         =   0   'False
         Width           =   8865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
         Height          =   4110
         Left            =   -74940
         TabIndex        =   3
         Top             =   1320
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   7267
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   15
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
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
         _Band(0).Cols   =   15
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "(模糊比對)"
         Height          =   180
         Left            =   -73530
         TabIndex        =   21
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "名片臨時編號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   17
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "寄發信函存放路徑："
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1620
      End
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   30
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   705
      Left            =   2010
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1358
      _ExtentY        =   1252
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "檔案名稱                                                             "
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.FileListBox File1 
      Height          =   240
      Left            =   1410
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   8070
      TabIndex        =   4
      Top             =   60
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   810
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm140418"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'程式使用的(範本格式)電子檔存放在:
'位置: \\LINUX\PolyCOM\TaieNew\RptSample
'電子檔名: (匯入系統) 日程表暨紀錄(2019 APAA) [整批匯入至往來記錄]-範本.xls
'Memo By Sindy 2022/2/17 Form2.0已修改
'Create By Sindy 2018/6/6
Option Explicit

Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iLine1 As Integer
Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
Public cmdState As Integer
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim i As Integer, j As Integer
Dim m_bolPrintRight As Boolean
Dim iNowLine As Integer 'Add Sindy 2021/1/28
Dim iRowHeight As Integer 'Add Sindy 2021/1/28


'Add By Sindy 2019/7/30
Private Sub cboSort_Click()
Dim iPos As Integer
   
   iPos = InStr(cboSort.Text, Chr(1))
   If iPos > 0 Then
      cboSort.Text = Left(cboSort.Text, iPos - 1)
   End If
   cboSort.Tag = Left(Trim(cboSort.Text), 3)
End Sub

'Add By Sindy 2019/7/30
Private Sub cboSort_GotFocus()
   If cboSort.Locked = False Then
      CloseIme
      SendMessage cboSort.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

'Add By Sindy 2019/7/30
Private Sub ChkExcel_Click()
   cmdExcel.Caption = "匯入Excel檔(&E)"
   If ChkExcel.Value = 1 Then
      cmdExcel.Caption = "選擇Excel檔案"
      cmdExcel.Enabled = True
      FrmExcel.Visible = True
      FrmExcel.ZOrder '移至頂層
      cmdImPort.Visible = False
      cboSort.SetFocus
      SSTab1.Tab = 0
      SSTab1.TabVisible(1) = False
      txtPath1.Visible = False
      Command2.Visible = False
   Else
      cmdExcel.Enabled = False
      FrmExcel.Visible = False
      cmdImPort.Visible = True
      SSTab1.TabVisible(1) = True
      txtPath1.Visible = True
      Command2.Visible = True
   End If
End Sub

'Add By Sindy 2019/7/26
Private Sub cmdExcel_Click()
Dim xlsSalesPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim stFileName As String, strTempFileName As String
Dim strTitleField As String
Dim ii As Integer, jj As Integer
Dim intMaxRow As Integer 'Excel資料列
Dim intField As Integer
Dim strTitNM As String
Dim strNo As String '編號
Dim strContact As String, strContactNm As String '聯絡人
Dim strContactTit As String
Dim strEmpArr(1 To 50) As String, intEmpCnt As Integer, strEmpTit As String '本所人員
Dim strCR08Tit As String '內容
Dim bolExists As Boolean, bolConn As Boolean, iRow As Integer
Dim strCR01 As String, StrCR05 As String, strCR06 As String, strCR07 As String, strCR09 As String
Dim strCR08 As String, strCR19 As String
Dim varTmp As Variant, varTmp2 As Variant
Dim dblMaxWidth As Double
Dim strShowText As String 'Add By Sindy 2023/11/8
   
On Error GoTo flgErr
   
   'Add By Sindy 2022/1/28
   strTempFileName = App.path & "\$$" & UCase(Me.Name) & ".xls"
   If Dir(strTempFileName) <> "" Then
      Kill strTempFileName
   End If
   '2022/1/28 END
   
   If Trim(cboSort.Text) = "" Then
      ShowMsg "往來類別不可為空白，若不在選項內請自行輸入 !"
      cboSort.SetFocus
      Exit Sub
   End If
   If Trim(txtCR06.Text) = "" Then
      ShowMsg "主旨不可為空白，請輸入 !"
      txtCR06.SetFocus
      Exit Sub
   End If
   If Trim(cboPlace.Text) = "" Then
      ShowMsg "場合不可為空白，若不在選項內請自行輸入 !"
      cboPlace.SetFocus
      Exit Sub
   End If
   
   stFileName = ""
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "Excel檔案 (*.xls 或 *.xlsx)|*.xls;*.xlsx"
      '.Filter = "Excel檔案 (*.xls 或 *.xlsx)"
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
'            sFile = Split(.FileName, ChrW$(0))
'            '記錄路徑
'            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
'            For ii = 1 To UBound(sFile)
'               If InStr(sFile(ii), "\") > 0 Then
'                  stFileName = sFile(ii)
'               Else
'                  stFileName = sFile(0) & "\" & sFile(ii)
'               End If
'            Next ii
         Else
            'stFileName = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            stFileName = .FileName
         End If
      End If
   End With
   If stFileName = "" Then
       MsgBox "請選擇一個Excel檔案！"
       Exit Sub
   End If
   
   '開檔
   Screen.MousePointer = vbHourglass
   xlsSalesPoint.Workbooks.Open stFileName
   'xlsSalesPoint.Visible = True
   Set wksrpt = xlsSalesPoint.Worksheets(1)
   '把Excel的警告訊息關掉
   xlsSalesPoint.DisplayAlerts = False
   
   Load frmpic002
   frmpic002.Label1.Caption = "檢查資料中...請稍候..."
   frmpic002.Show
   frmpic002.ZOrder 0
   
   'xlsAgentPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   'xlsAgentPoint.Workbooks.add
   '先檢查Excel的內容:
   List1.Clear
   intMaxRow = 4
   '固定欄位
   'A.日期
   strTitNM = Trim(wksrpt.Range("A" & 4).Value)
   If strTitNM = "日期" Then
      Do While Trim(wksrpt.Range("A" & (intMaxRow + 1)).Value) <> ""
         intMaxRow = intMaxRow + 1 '取得資料區最後列數
         strShowText = "檢查A欄位(日期) 第 " & intMaxRow & " 筆" 'Add By Sindy 2023/11/8
         '檢查是否為日期資料
         'Modify By Sindy 2022/1/25
         'If ChkDate(DBDATE(wksrpt.Range("A" & intMaxRow).Value), False) = False Then
         'Add By Sindy 2023/7/18
         If IsDate(Format(wksrpt.Range("A" & intMaxRow).Value, "yyyy/mm/dd")) = False Then
            List1.AddItem "A 欄第 " & intMaxRow & " 列並非日期, 請確認"
         Else
         '2023/7/18 END
            If ChkDate(ChangeWDateStringToWString(Format(wksrpt.Range("A" & intMaxRow).Value, "yyyy/mm/dd")), False) = False Then
            '2022/1/25 END
               List1.AddItem "A 欄第 " & intMaxRow & " 列並非日期, 請確認"
            End If
         End If
      Loop
   Else
      List1.AddItem "A 欄並非[日期], 有誤"
   End If
   'B.編號
   strTitNM = Trim(wksrpt.Range("B" & 4).Value)
   If strTitNM = "編號" Then
      For jj = 5 To intMaxRow
         strShowText = "檢查B欄位(編號) 第 " & jj & " 筆" 'Add By Sindy 2023/11/8
         '檢查編號是否存在
         If Trim(wksrpt.Range("B" & jj).Value) <> "" Then
            If InStr(Trim(wksrpt.Range("B" & jj).Value), ",") = 0 Then
               strNo = ChangeCustomerL(Trim(wksrpt.Range("B" & jj).Value))
               strExc(0) = "select pcu01 from potcustomer where pcu01(+)='" & Left(strNo, 8) & "' and pcu02(+)='0'" & _
                           " union " & _
                           "select cu01 from customer where cu01(+)='" & Left(strNo, 8) & "' and cu02(+)='0'" & _
                           " union " & _
                           "select fa01 from fagent where fa01(+)='" & Left(strNo, 8) & "' and fa02(+)='0'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  List1.AddItem "B 欄第 " & jj & " 列編號(" & strNo & ")不存在, 請確認"
               End If
            Else
               List1.AddItem "B 欄第 " & jj & " 列編號(" & Trim(wksrpt.Range("B" & jj).Value) & ")不可放多個, 請確認"
            End If
         Else
            List1.AddItem "B 欄第 " & jj & " 列編號不可空白, 請確認"
         End If
      Next jj
   Else
      List1.AddItem "B 欄並非[編號], 有誤"
   End If
   'C.名稱
   strTitNM = Trim(wksrpt.Range("C" & 4).Value)
   If strTitNM = "名稱" Then
      For jj = 5 To intMaxRow
         strShowText = "檢查C欄位(名稱) 第 " & jj & " 筆" 'Add By Sindy 2023/11/8
         '檢查名稱不可空白
         If Trim(wksrpt.Range("C" & jj).Value) = "" Then
            List1.AddItem "C 欄第 " & jj & " 列名稱不可空白, 請確認"
         End If
      Next jj
   Else
      List1.AddItem "C 欄並非[名稱], 有誤"
   End If
   'Add By Sindy 2023/8/10
   'D.財務處告知有產生國外交際餐費
   strTitNM = Trim(wksrpt.Range("D" & 4).Value)
   If strTitNM = "財務處告知有產生國外交際餐費" Then
   Else
      List1.AddItem "D 欄並非[財務處告知有產生國外交際餐費], 有誤"
   End If
   '2023/8/10 END
   '會面人
   intField = 4: strContactTit = ""
   For ii = intField To 99
      strTitleField = GetFieldStr(ii, 65) '65.A~90.Z
      If InStr(Trim(wksrpt.Range(strTitleField & 4).Value), "會面人") > 0 Then
         intField = ii + 1
         strContactTit = strContactTit & "," & strTitleField
         For jj = 5 To intMaxRow
            strShowText = "檢查" & strTitleField & "欄位(會面人) 第 " & jj & " 筆" 'Add By Sindy 2023/11/8
            strNo = ChangeCustomerL(Trim(wksrpt.Range("B" & jj).Value))
            strContact = UCase(Trim(wksrpt.Range(strTitleField & jj).Value))
            '檢查若有輸入聯絡人必須存在
            If strContact <> "" _
               And strNo <> "" And InStr(strNo, ",") = 0 Then
               '聯絡人編號
               strExc(0) = "select pcc01,pcc02 from PotcustCont where pcc01='" & Left(strNo, 8) & "' and pcc02='" & strContact & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  'Add By Sindy 2021/1/28 + 聯絡人名稱
                  strContact = UCase(strContact)
                  strExc(0) = "select pcc01,pcc02 from PotcustCont where pcc01='" & Left(strNo, 8) & "'" & _
                              " and (upper(pcc03)='" & strContact & "' or instr('" & strContact & "',replace(pcc04,' ',''))>0 or pcc05='" & strContact & "')"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then
                  '2021/1/28 END
                     wksrpt.Range(strTitleField & jj).Value = "@" & Trim(wksrpt.Range(strTitleField & jj).Value) 'Add By Sindy 2021/1/28
                     'List1.AddItem strTitleField & " 欄第 " & jj & " 列" & strNo & "聯絡人(" & strContact & ")不存在, 請確認"
                  Else
                     'Modify By Sindy 2023/2/4 後面會使用到會面人名稱
                     'wksrpt.Range(strTitleField & jj).Value = RsTemp.Fields("pcc02") 'Add By Sindy 2021/1/28
                     wksrpt.Range(strTitleField & jj).Value = RsTemp.Fields("pcc02") & "|" & Trim(wksrpt.Range(strTitleField & jj).Value)
                  End If
               Else
                  'Modify By Sindy 2023/2/4 後面會使用到會面人名稱
                  'wksrpt.Range(strTitleField & jj).Value = RsTemp.Fields("pcc02") 'Add By Sindy 2021/1/28
                  wksrpt.Range(strTitleField & jj).Value = RsTemp.Fields("pcc02") & "|" & Trim(wksrpt.Range(strTitleField & jj).Value)
               End If
            End If
         Next jj
      Else
         Exit For
      End If
   Next ii
   If strContactTit <> "" Then strContactTit = Mid(strContactTit, 2)
   '本所人員
   intEmpCnt = 0: strEmpTit = ""
   For ii = intField To 99
      strTitleField = GetFieldStr(ii, 65) '65.A~90.Z
      If (InStr(Trim(wksrpt.Range(strTitleField & 3).Value), "本所人員") > 0 And intEmpCnt = 0) Or _
         (Trim(wksrpt.Range(strTitleField & 4).Value) <> "" And intEmpCnt >= 1) Then
         '檢查是否已存在本所人員名單中
         bolExists = False
         For jj = 1 To intEmpCnt
            If InStr(strEmpArr(jj), Trim(wksrpt.Range(strTitleField & 4).Value)) > 0 Then
               bolExists = True
               'Modify By Sindy 2021/1/28
               'Exit For
               GoTo ExitEmp
               '2021/1/28 END
            End If
         Next jj
         If bolExists = False Then
            intField = ii + 1
            strEmpTit = strEmpTit & "," & strTitleField
            strShowText = "檢查" & strTitleField & "欄位(本所人員) " & strEmpTit 'Add By Sindy 2023/11/8
            '檢查是否為本所人員
            If Trim(wksrpt.Range(strTitleField & 4).Value) <> "" Then
               strExc(0) = "select st01,st02 from staff" & _
                           " where (st01='" & Trim(wksrpt.Range(strTitleField & 4).Value) & "'" & _
                           " or st02='" & Trim(wksrpt.Range(strTitleField & 4).Value) & "')" & _
                           " and substr(st01,1,1)<>'F' and substr(st01,4,1)<>'9'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp.RecordCount = 1 Then
                     intEmpCnt = intEmpCnt + 1
                     strEmpArr(intEmpCnt) = RsTemp.Fields("st02") & "," & RsTemp.Fields("st01")
                     '內容只能輸入V
                     For jj = 5 To intMaxRow
                        strShowText = "檢查" & strTitleField & "欄位(本所人員) 第 " & jj & " 筆" 'Add By Sindy 2023/11/8
                        If UCase(Trim(wksrpt.Range(strTitleField & jj).Value)) <> "V" _
                           And Trim(wksrpt.Range(strTitleField & jj).Value) <> "" Then
                           List1.AddItem strTitleField & " 欄第 " & jj & " 列資料不是空白只能輸入V, 請確認"
                        End If
                     Next jj
                  Else
                     List1.AddItem strTitleField & " 欄人員(" & Trim(wksrpt.Range(strTitleField & 4).Value) & ")查出多筆資料, 請確認"
                  End If
               Else
                  List1.AddItem strTitleField & " 欄人員(" & Trim(wksrpt.Range(strTitleField & 4).Value) & ")不存在, 請確認"
               End If
            Else
               List1.AddItem strTitleField & " 欄人員空白, 請輸入"
            End If
         End If
      Else
         Exit For
      End If
   Next ii
ExitEmp: 'Add By Sindy 2021/1/28
   If strEmpTit <> "" Then strEmpTit = Mid(strEmpTit, 2)
   '記錄內容
   strCR08Tit = ""
   For ii = intField To 99
      strTitleField = GetFieldStr(ii, 65) '65.A~90.Z
      If Trim(wksrpt.Range(strTitleField & 4).Value) = "" Then
         Exit For
         '代表結束,不須再往後檢查資料
      Else
         intField = ii + 1
         strCR08Tit = strCR08Tit & "," & strTitleField
         '必須存在於本所人員名單中
         bolExists = False
         For jj = 1 To intEmpCnt
            If InStr(strEmpArr(jj), Trim(wksrpt.Range(strTitleField & 4).Value)) > 0 Then
               bolExists = True
               Exit For
            End If
         Next jj
         If bolExists = False Then
            List1.AddItem strTitleField & " 欄人員(" & Trim(wksrpt.Range(strTitleField & 4).Value) & ")在本所人員清單中不存在, 請確認"
         End If
         '內容不能只輸入V
         For jj = 5 To intMaxRow
            strShowText = "檢查" & strTitleField & "欄位(記錄內容) 第 " & jj & " 筆" 'Add By Sindy 2023/11/8
            If UCase(Trim(wksrpt.Range(strTitleField & jj).Value)) = "V" Then
               List1.AddItem strTitleField & " 欄第 " & jj & " 列資料內容不能只輸入V, 請確認"
            End If
         Next jj
      End If
   Next ii
   If strCR08Tit <> "" Then strCR08Tit = Mid(strCR08Tit, 2)
   strShowText = "" 'Add By Sindy 2023/11/8
   Unload frmpic002
   
   '無錯誤,即可匯入資料
   If List1.ListCount = 0 Then
      'Modify By Sindy 2023/2/4
      If MsgBox("資料檢查完畢！確定要執行匯入資料了嗎？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         GoTo RunClose
      End If
      '2023/2/4 END
      
      dblMaxWidth = 8820
      TextCnt.Width = 0
      cnnConnection.BeginTrans: bolConn = True
      For jj = 5 To intMaxRow
         TextCnt.Width = dblMaxWidth / intMaxRow * jj
         strCR01 = AutoNo("K", 6) '往來記錄編號
         strDate = DBDATE(Trim(wksrpt.Range("A" & jj).Value))
         strNo = Left(ChangeCustomerL(Trim(wksrpt.Range("B" & jj).Value)), 8) & "0"
         '聯絡人=會面人
         strContact = "": strContactNm = ""
         strCR08 = ""
         varTmp = Split(strContactTit, ",")
         For ii = 0 To UBound(varTmp)
            If Trim(wksrpt.Range(varTmp(ii) & jj).Value) <> "" Then
               'Modify By Sindy 2023/2/4 內容裡要增加會面人名稱
               If Left(Trim(wksrpt.Range(varTmp(ii) & jj).Value), 1) <> "@" Then '第1碼是@代表沒有聯絡人代碼
                  varTmp2 = Split(Trim(wksrpt.Range(varTmp(ii) & jj).Value), "|") '有聯絡人代碼
                  strContact = strContact & "," & Format(varTmp2(0), "00")
                  strContactNm = varTmp2(1)
               Else
                  strContactNm = Mid(Trim(wksrpt.Range(varTmp(ii) & jj).Value), 2)
               End If
               If strCR08 <> "" Then strCR08 = strCR08 & vbCrLf
               strCR08 = strCR08 & "會面人" & ii + 1 & ":" & strContactNm
               '2023/2/4 END
            End If
         Next ii
         If strContact <> "" Then strContact = Mid(strContact, 2)
         StrCR05 = cboSort.Tag '往來類別
         strCR06 = txtCR06 '主旨
         strCR07 = cboPlace.Text '場合
         strCR09 = UCase(Trim(wksrpt.Range("D" & jj).Value)) '財務處告知有產生國外交際餐費 Add By Sindy 2023/8/10
         '內容
         varTmp = Split(strCR08Tit, ",")
         For ii = 0 To UBound(varTmp)
            If Trim(wksrpt.Range(varTmp(ii) & jj).Value) <> "" Then
               '若輸入員編改抓名字
               strExc(10) = GetPrjSalesNM(Trim(wksrpt.Range(varTmp(ii) & 4).Value))
               If Trim(strExc(10)) = "" Then
                  strExc(10) = Trim(wksrpt.Range(varTmp(ii) & 4).Value)
               End If
               If strCR08 <> "" Then strCR08 = strCR08 & vbCrLf & vbCrLf
               strCR08 = strCR08 & strExc(10) & ":" & Trim(wksrpt.Range(varTmp(ii) & jj).Value)
            End If
         Next ii
         '接洽同仁
         strCR19 = ""
         varTmp = Split(strEmpTit, ",")
         For ii = 0 To UBound(varTmp)
            If UCase(Trim(wksrpt.Range(varTmp(ii) & jj).Value)) = "V" Then
               varTmp2 = Split(strEmpArr(ii + 1), ",")
               If UBound(varTmp2) >= 0 Then
                  strCR19 = strCR19 & "," & varTmp2(1)
               End If
            End If
         Next ii
         If strCR19 <> "" Then strCR19 = Mid(strCR19, 2)
         '新增
         'Modify By Sindy 2023/8/10 + ,CR09
         strSql = "INSERT INTO ContactRecord(CR01,CR02,CR03,CR04,CR05" & _
                  ",CR06,CR07,CR08,CR09,CR19)" & _
                  " VALUES('" & strCR01 & "'," & strDate & "," & CNULL(strNo) & "," & CNULL(strContact) & ",'" & StrCR05 & "'" & _
                  "," & CNULL(strCR06) & "," & CNULL(strCR07) & ",'" & ChgSQL(strCR08) & "'," & CNULL(strCR09) & "," & CNULL(strCR19) & ")"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      Next jj
      
      cnnConnection.CommitTrans: bolConn = False
      
      TextCnt.Width = dblMaxWidth: DoEvents
      
      MsgBox "資料匯入完畢！" & vbCrLf & vbCrLf & "【注意：請一定要檢查「往來記錄」的內容是否有匯入正確!!!】"
   Else
      MsgBox "資料有誤, 請修改後再重新匯入！"
   End If
   
RunClose:
   '另存
   xlsSalesPoint.Workbooks(1).SaveAs FileName:=strTempFileName, FileFormat:=56
   '關閉
   xlsSalesPoint.Workbooks.Close
   '離開
   xlsSalesPoint.Quit
   Set wksrpt = Nothing
   Set xlsSalesPoint = Nothing
   Screen.MousePointer = vbDefault

   Exit Sub
   
flgErr:
   'Add By Sindy 2023/11/8
   If strShowText <> "" Then
      List1.AddItem strShowText
   End If
   '2023/11/8 END
   If TypeName(frmpic002) <> "Nothing" Then Unload frmpic002
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   
   If stFileName <> "" Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strTempFileName, FileFormat:=56
      xlsSalesPoint.Workbooks.Close
      xlsSalesPoint.Quit
   End If
   Set wksrpt = Nothing
   Set xlsSalesPoint = Nothing
   
   If Err.Number <> 0 Then
       MsgBox iRow & " 筆 : " & Err.Description
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'匯入
Private Sub cmdImPort_Click()
Dim fs
Dim dblFCnt As Double
Dim dblMaxWidth As Double
Dim bolErr As Boolean
Dim intRow As Integer
   
On Error GoTo ErrHand
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
      If MsgBox("目前是測試資料庫，確定要匯入嗎？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   If Right(Trim(txtPath1), 1) <> "\" Then txtPath1 = Trim(txtPath1) & "\"
   
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
      
   'Add By Sindy 2018/10/30
   If MsgBox("確定要匯入" & txtPath1.Text & " 此資料夾( " & File1.ListCount & " 筆 )電子檔嗎？", vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass

   dblMaxWidth = 8820
   TextCnt.Width = 0
   List1.Clear
   Text3.Text = ""
   intRow = 0
   For dblFCnt = File1.ListCount - 1 To 0 Step -1
      intRow = intRow + 1
      TextCnt.Width = dblMaxWidth / Val(File1.ListCount) * intRow
      Text3.Text = intRow & " / " & Val(File1.ListCount)
      DoEvents
      bolErr = False
      
      '檢查檔案是否正在使用中
      If PUB_ChkFileOpening(txtPath1.Text & Trim(File1.List(dblFCnt))) = True Then
         bolErr = True
         List1.AddItem Trim(File1.List(dblFCnt)) & " : 檔案正在使用中，請關閉才可執行匯入！", 0: SetListScroll List1
      End If
      '檔名後4碼為.MSG者才須匯入
      If UCase(Right(Trim(File1.List(dblFCnt)), 4)) <> ".MSG" Then
         bolErr = True
         List1.AddItem Trim(File1.List(dblFCnt)) & " : 檔名後4碼必須為.MSG者才能匯入！", 0: SetListScroll List1
      End If
      
      '國際會議郵件
      If bolErr = False Then
         If PUB_IPDeptISDMail(Me, "", txtPath1, txtPath1, Trim(File1.List(dblFCnt)), , List1) = True Then
            DoEvents
         Else
            SetListScroll List1
         End If
      End If
   Next dblFCnt

   TextCnt.Width = dblMaxWidth: DoEvents
   
   Screen.MousePointer = vbDefault
   
   MsgBox "匯入完畢！"
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

'Add By Sinidy 2021/1/28
Private Sub cmdPrint_Click()
   Dim ii As Integer, jj As Integer
   Dim strFontName As String
   
'   If Check1.Value <> vbChecked And Check2.Value <> vbChecked Then
'      MsgBox "請勾選要列印的內容！", vbInformation
'      Exit Sub
'   End If
   
   strFontName = Printer.FontName
   Printer.FontName = "細明體"
   
'   '待匯入案件
'   If Check1.Value = 1 Then
'      GetPleft 1
'      PrintTitle 1
'      For jj = 1 To MSHFlexGrid1.Rows - 1
'         For ii = 1 To 4
'            strTemp(ii) = "" & MSHFlexGrid1.TextMatrix(jj, ii)
'         Next ii
'         If (iNowLine + 2) * iRowHeight > Printer.ScaleHeight Then
'            Printer.NewPage
'            PrintTitle 1  '列印表頭
'         End If
'         PrintDetail '列印明細
'      Next jj
'      Printer.EndDoc
'   End If
   
   '匯入結果
'   If Check2.Value = 1 Then
      GetPleft 2
      PrintTitle 2
      For jj = List1.ListCount - 1 To 0 Step -1
         strTemp(1) = List1.List(jj)
         If (iNowLine + 2) * iRowHeight > Printer.ScaleHeight Then
            Printer.NewPage
            PrintTitle 2 '列印表頭
         End If
         
         PrintDetail '列印明細
         
      Next jj
      Printer.EndDoc
'   End If
   
   Printer.FontName = strFontName
End Sub

Sub GetPleft(Optional pIndex As Integer = 1)
   iRowHeight = 300
'   If pIndex = 1 Then
'      ReDim PLeft(1 To 5)
'      ReDim strTemp(1 To 5)
'
'      PLeft(1) = 500
'      PLeft(2) = 2500
'      PLeft(3) = 6000
'      PLeft(4) = 7500
'      PLeft(5) = 9000
'   Else
'      ReDim PLeft(1 To 1)
'      ReDim strTemp(1 To 1)
      
      PLeft(1) = 500
'   End If
End Sub

Sub PrintTitle(Optional pIndex As Integer = 1)
Dim strTitle As String

iNowLine = 0
'If pIndex = 1 Then
'   strTitle = "待匯入案件"
'Else
   strTitle = "匯入結果"
'End If

iNowLine = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(Me.Caption) / 2)
Printer.CurrentY = iNowLine * iRowHeight
Printer.Print Me.Caption

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
'If pIndex = 1 Then
'   Printer.CurrentX = PLeft(1)
'   Printer.CurrentY = iNowLine * iRowHeight
'   Printer.Print MSHFlexGrid1.TextMatrix(0, 1)
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iNowLine * iRowHeight
'   Printer.Print MSHFlexGrid1.TextMatrix(0, 2)
'   Printer.CurrentX = PLeft(3)
'   Printer.CurrentY = iNowLine * iRowHeight
'   Printer.Print MSHFlexGrid1.TextMatrix(0, 3)
'   Printer.CurrentX = PLeft(4)
'   Printer.CurrentY = iNowLine * iRowHeight
'   Printer.Print MSHFlexGrid1.TextMatrix(0, 4)
'Else
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print strTitle
'End If
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

Private Sub cmdQuery_Click()
Dim s As Integer
Dim strSQLCon As String
   
On Error GoTo ErrHnd
   
   If Txt1 = "" Then
      s = MsgBox("名片臨時編號不可空白", , "輸入條件錯誤")
      Text1.SetFocus
      Exit Sub
   Else
      Txt1 = Trim(Txt1)
      strSQLCon = " and instr(upper(pcc25),upper('" & Txt1 & "'))>0"
   End If
   
   Screen.MousePointer = vbHourglass
   grdDataList.Clear
   grdDataList.Rows = 2
   SetDataListWidth
   
   strSql = "select ' ' AS V,cu01||cu02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號" & _
            ",nvl(cu04,decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 名稱,pcc02||'.'||nvl(PCC03,decode(PCC05,null,PCC04,PCC05)) as 聯絡人,pcc25 as 名片臨時編號" & _
            ",na03 as 國籍,st02 as 智權人員,decode(cu142,null,cu80,getdizhang(cu142,'Y')) as 狀態,cu79 as 備註" & _
            " from potcustcont,customer,NATION,STAFF where pcc25 is not null" & _
            " and pcc01=cu01(+) and cu02='0'" & _
            " and cu10=na01(+) and cu13=st01(+)" & strSQLCon
   strSql = strSql & " Union All " & _
            "select ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號" & _
            ",nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,pcc02||'.'||nvl(PCC03,decode(PCC05,null,PCC04,PCC05)) as 聯絡人,pcc25 as 名片臨時編號" & _
            ",NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註" & _
            " from potcustcont,fagent,NATION where pcc25 is not null" & _
            " and pcc01=fa01(+) and fa02='0'" & _
            " and fa10=NA01(+)" & strSQLCon
   strSql = strSql & " Union All " & _
            "select ' ' as v ,pcu01||pcu02||decode(pcu02,'0','','＊') as 編號" & _
            ",nvl(pcu08,decode(pcu03,null,pcu07,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06))) as 名稱,pcc02||'.'||nvl(PCC03,decode(PCC05,null,PCC04,PCC05)) as 聯絡人,pcc25 as 名片臨時編號" & _
            ",NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註" & _
            " from potcustcont,potcustomer,nation,staff where pcc25 is not null" & _
            " and pcc01=pcu01(+) and pcu02='0'" & _
            " and pcu09=na01(+) and substr(ltrim(pcu38),1,5)=st01(+)" & strSQLCon
   strSql = strSql & " Union All " & _
            "select ' ' AS V ,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號" & _
            ",NVL(PoC03,Decode(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS 名稱,pcc02||'.'||nvl(pcc03,decode(pcc05,null,pcc04,pcc05)) as 聯絡人,pcc25 as 名片臨時編號" & _
            ",NA03 AS 國籍,ST02 AS 智權人員,PoC14 AS 狀態,PoC15 AS 備註" & _
            " from potcustcont,potcustomer1,nation,staff where pcc25 is not null" & _
            " and pcc01=poc01(+) and poc02='0'" & _
            " and PoC04=NA01(+) and poc13=ST01(+)" & strSQLCon
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
      Set grdDataList.Recordset = adoRecordset
      If Not cmdOK(0).Enabled Then cmdOK(0).Enabled = True
      If Not cmdOK(2).Enabled Then cmdOK(2).Enabled = True
      If Not cmdOK(8).Enabled Then cmdOK(8).Enabled = True
      If Not cmdOK(9).Enabled Then cmdOK(9).Enabled = True
      Me.grdDataList.Visible = False
      '若只有一筆資料, 則直接設定為點選此筆資料
      With Me.grdDataList
         If .Rows = 2 Then
            .row = 1
            .col = 1
            If .Text <> "" Then
              .Visible = False
              .row = 1
              .col = 0
              .Text = "V"
              .Visible = True
            End If
         End If
      End With
      Me.grdDataList.Visible = True
   Else
      ShowNoData
      cmdOK(0).Enabled = False
      cmdOK(2).Enabled = False
      cmdOK(8).Enabled = False
      cmdOK(9).Enabled = False
      grdDataList.Clear
      SetDataListWidth
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHnd:
   If Err.Number = -2147217900 Then
      MsgBox "輸入的文字無法查詢,請電腦中心協助！"
   Else
      MsgBox Err.Description
   End If
   Screen.MousePointer = vbDefault
End Sub

''列印
'Private Sub cmdPrint_Click()
'Dim i As Integer, j As Integer
'
'   iLine1 = 0
'   For j = List1.ListCount - 1 To 0 Step -1
'      For i = 1 To 1
'         strTemp(i) = ""
'      Next i
'      strTemp(1) = List1.List(j)
'      If iLine1 > 52 Or iLine1 = 0 Then
'         If iLine1 > 0 Then Printer.NewPage: iLine1 = 0
'         PrintTitle '列印表頭
'      End If
'      PrintDetail '列印明細
'   Next j
'   Printer.EndDoc
'End Sub

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.msg"
      .Filter = "Msg檔案 (*.msg)|*.msg"
'      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "DirMsg", "") <> "" Then
'         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "DirMsg", "")
'      Else
         .InitDir = txtPath1.Text
'      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
'            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "DirMsg", sFile(0)
            txtPath1.Text = sFile(0)
         Else
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
'               For ii = Len(.FileName) To 1 Step -1
'                  If Mid(Trim(.FileName), ii, 1) = "\" Then
'                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "DirMsg", Left(.FileName, InStrRev(.FileName, "\") - 1)
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
   
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
      txtPath1.Text = PUB_Getdesktop
   Else
      txtPath1.Text = Pub_GetSpecMan("國外部開拓分信電子檔存放路徑")
   End If
'   If GetSetting("TAIE", "FCP", UCase(Me.Name) & "DirMsg", "") <> "" Then
'      txtPath1.Text = GetSetting("TAIE", "FCP", UCase(Me.Name) & "DirMsg", "")
'   'Else
'   '   txtPath1.Text = PUB_Getdesktop
'   End If
   SSTab1.Tab = 0
   m_bolPrintRight = IsUserHasRightOfFunction("frm100102_1", strPrint, False)
   cmdState = -1
   Text3.Text = ""
   
   'Add By Sindy 2019/7/30
   '往來類別
   cboSort.Clear
   strExc(0) = "select ac02,ac03 from allcode where ac01='11' order by ac01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         cboSort.AddItem RsTemp.Fields("ac02") & " " & RsTemp.Fields("ac03")
         RsTemp.MoveNext
      Loop
   End If
   '場合
   cboPlace.Clear
   cboPlace.AddItem "線上會議", 0 'Add By Sindy 2022/1/20
   cboPlace.AddItem "Email", 0
   cboPlace.AddItem "會議場合", 0
   cboPlace.AddItem "彼所/公司", 0
   cboPlace.AddItem "台一", 0
   '2019/7/30 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140418 = Nothing
End Sub

Private Sub GrdDataList_Click()
Dim strCopyTxt As String '複製編號文字
Dim LongRow As Long
   
   grdDataList.row = grdDataList.MouseRow
   LongRow = grdDataList.MouseRow
   
   '選到編號欄=複製
   grdDataList.col = grdDataList.MouseCol
   If grdDataList.col = 1 Then
      strCopyTxt = grdDataList.TextMatrix(grdDataList.row, grdDataList.col)
      If strCopyTxt <> "" Then
         '複製編號至剪貼簿
         Clipboard.SetText strCopyTxt
         grdDataList.CellBackColor = QBColor(7)
         MsgBox "編號已複製", , MsgText(21)
      End If
      Exit Sub
   End If
   grdDataList.Visible = False
   grdDataList.col = 0
   'If grdDataList.row <> 0 Then
   If LongRow <> 0 Then
      If grdDataList.Text = "V" Then
         grdDataList.Text = ""
         For i = 0 To grdDataList.Cols - 1
            If i <> 1 Then
               grdDataList.col = i
               grdDataList.CellBackColor = QBColor(15)
            End If
         Next i
      Else
         grdDataList.Text = "V"
         For i = 0 To grdDataList.Cols - 1
            If i <> 1 Then
               grdDataList.col = i
               grdDataList.CellBackColor = &HFFC0C0
            End If
         Next i
      End If
   End If
   grdDataList.Visible = True
End Sub

Private Sub SetDataListWidth()
   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 1: grdDataList.Text = "編號"
   grdDataList.ColWidth(1) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "名稱"
   grdDataList.ColWidth(2) = 2500
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "聯絡人"
   grdDataList.ColWidth(3) = 1500
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "名片臨時編號"
   grdDataList.ColWidth(4) = 1400
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "國籍"
   grdDataList.ColWidth(5) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "智權人員"
   grdDataList.ColWidth(6) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 7: grdDataList.Text = "狀態"
   grdDataList.ColWidth(7) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 8: grdDataList.Text = "備註"
   grdDataList.ColWidth(8) = 2000
   grdDataList.CellAlignment = flexAlignLeftCenter
End Sub

Public Sub PubShowNextData()
Dim blnPrintAdd As Boolean
Dim ii As Integer
Dim j As Integer
Dim strTmp As String

   Select Case cmdState
      Case 0 '申請人資料
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               grdDataList.col = 1
               For j = 0 To grdDataList.Cols - 1
                  If j <> 1 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               grdDataList.col = 1
               Screen.MousePointer = vbHourglass
               '加判斷第一碼切不同畫面
               strTmp = Pub_RplStr(grdDataList.Text)
               Select Case Left(strTmp, 1)
                  Case "X"
                     If Mid(strTmp, 10, 1) = "-" Then
                        strTmp = Left(strTmp, 9)
                     End If
                     frm100101_11.Show
                     frm100101_11.Tag = strTmp
                     frm100101_11.StrMenu
                  Case "Y" '代理人
                     '+判斷有權限的才能查代理人的案件資料
                     If bolFNation = True Then
                        If Mid(strTmp, 10, 1) = "-" Then
                           strTmp = Left(strTmp, 9)
                        End If
                        frm100101_10.Show
                        frm100101_10.Tag = strTmp
                        frm100101_10.StrMenu
                     Else
                        Me.Show
                        MsgBox "您無查詢國外代理人資料權限！", vbInformation
                     End If
                  Case "R"
                     '判斷是國外或是國內潛在客戶
                     strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strTmp, 8) & "' and pcu02(+)='" & Mid(strTmp, 9, 1) & "' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     strExc(2) = ""
                     If intI = 1 Then
                        strExc(2) = "" & RsTemp.Fields(0)
                     End If
                     If strExc(2) <> "" Then '國外
                        frm100101_14.Show
                        frm100101_14.Tag = strTmp
                        frm100101_14.StrMenu
                     Else '國內
                        frm100101_21.Show
                        frm100101_21.Tag = strTmp
                        frm100101_21.StrMenu
                     End If
               End Select
               Screen.MousePointer = vbDefault
               grdDataList.col = 0
               grdDataList.Text = ""
               grdDataList.col = 1
               For j = 0 To grdDataList.Cols - 1
                  If j <> 1 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               Me.Enabled = True
               Exit Sub
            End If
            Next i
            Me.Enabled = True
      Case 2 '關係企業
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
              grdDataList.col = 0
              grdDataList.row = i
              If Trim(grdDataList.Text) = "V" Then
                  grdDataList.col = 0
                  grdDataList.Text = ""
                  grdDataList.col = 1
                  For j = 0 To grdDataList.Cols - 1
                     If j <> 1 Then
                         grdDataList.col = j
                         grdDataList.CellBackColor = QBColor(15)
                     End If
                  Next j
                  grdDataList.col = 1
                  '檢查國內外權限
                  If CheckSR12(Pub_RplStr(grdDataList.Text)) = True Then
                     Screen.MousePointer = vbHourglass
                     cmdOK(2).Enabled = False
                     Screen.MousePointer = vbDefault
                  End If
              End If
            Next i
            Me.Enabled = True
      Case 8 '往來記錄
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               grdDataList.col = 1
               For j = 0 To grdDataList.Cols - 1
                  If j <> 1 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               grdDataList.col = 1
               Screen.MousePointer = vbHourglass
               strTmp = Pub_RplStr(grdDataList.Text)
               
               '判斷是國外或是國內潛在客戶
               '客戶檔
               strExc(3) = "select cu12,cu13 from customer where cu01(+)='" & Left(strTmp, 8) & "' and cu02(+)='" & Mid(strTmp, 9, 1) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
               strExc(4) = ""
               If intI = 1 Then
                  strExc(4) = "" & RsTemp.Fields("cu12")
               End If
               '潛在客戶檔
               strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strTmp, 8) & "' and pcu02(+)='" & Mid(strTmp, 9, 1) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               strExc(2) = ""
               If intI = 1 Then
                  strExc(2) = "" & RsTemp.Fields(0)
               End If
'               If strExc(2) <> "" Or Left(Trim(strTmp), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Then '國外
                  frm100101_15.Show
                  frm100101_15.Tag = strTmp
                  'Modify By Sindy 2020/5/19
                  If strExc(2) <> "" Or Left(Trim(strTmp), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Then '國外
                     frm100101_15.m_quyDataKind = 0
                     frm100101_15.StrMenu
                  Else
                     frm100101_15.m_quyDataKind = 1
                     frm100101_15.StrMenu2
                  End If
                  '2020/5/19 END
'               Else '國內
'                  frm100101_20.Show
'                  frm100101_20.Tag = strTmp
'                  frm100101_20.StrMenu
'               End If
               
               Screen.MousePointer = vbDefault
               grdDataList.col = 0
               grdDataList.Text = ""
               grdDataList.col = 1
               For j = 0 To grdDataList.Cols - 1
                  If j <> 1 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               Me.Enabled = True
               Exit Sub
            End If
            Next i
            Me.Enabled = True
      Case 9 '聯絡人
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               grdDataList.col = 1
               For j = 0 To grdDataList.Cols - 1
                  If j <> 1 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               grdDataList.col = 1
               Screen.MousePointer = vbHourglass
               strTmp = Pub_RplStr(grdDataList.Text)
               '國內外客戶跑不同畫面
               Select Case Left(strTmp, 1)
                  '潛在客戶跑申請人資料畫面
                  Case "R"
                     frm100101_14.Show
                     frm100101_14.Tag = strTmp
                     frm100101_14.StrMenu
                  Case Else
                     strExc(2) = "F"
                     If Left(strTmp, 1) = "X" Then
                        strExc(0) = "select st03 from customer,staff where cu01(+)='" & Left(strTmp, 8) & "' and cu02(+)='" & Mid(strTmp, 9, 1) & "' and st01(+)=cu13"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           strExc(2) = "" & RsTemp.Fields(0)
                        End If
                     End If
                     If Left(strExc(2), 1) = "F" Then
                        frm100101_17.Show
                        frm100101_17.Tag = strTmp
                        frm100101_17.StrMenu
                     Else
                        frm100101_18.Show
                        frm100101_18.SetParent Me
                        'Mark by Lydia 2024/03/13
                        'frm100101_18.Label2(1).Visible = False
                        'frm100101_18.Combo1.Visible = False
                        'end 2024/03/13
                        frm100101_18.CmdOk1(1).Visible = False
                        frm100101_18.CmdOk1(2).Caption = Me.cmdOK(4).Caption
                        frm100101_18.Tag = strTmp
                        frm100101_18.CmdOk1(2).Enabled = m_bolPrintRight
                        frm100101_18.StrMenu
                     End If
               End Select
               Screen.MousePointer = vbDefault
               grdDataList.col = 0
               grdDataList.Text = ""
               grdDataList.col = 1
               For j = 0 To grdDataList.Cols - 1
                  If j <> 1 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               Me.Enabled = True
               Exit Sub
            End If
            Next i
            Me.Enabled = True
      Case Else
   End Select
End Sub

Private Sub List1_DblClick()
Dim hLocalFile As Long
   
   'If Right(Trim(txtPath1), 1) <> "\" Then txtPath1 = Trim(txtPath1) & "\"
   If InStr(List1.List(List1.ListIndex), ":") > 0 Then
      Screen.MousePointer = vbHourglass
      ShellExecute hLocalFile, "open", txtPath1 & Trim(Mid(List1.List(List1.ListIndex), 1, InStr(List1.List(List1.ListIndex), ":") - 1)), vbNullString, vbNullString, 1
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      Txt1.SetFocus
   Else
      If txtPath1.Visible = True Then txtPath1.SetFocus
   End If
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
   strMsg = "請輸入寄發信函存放路徑！"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   txtPath1.SetFocus
   Exit Function
End If

TxtValidate = True
End Function

'Sub GetPleft()
'PLeft(1) = 500
'PLeft(2) = 2500
'PLeft(3) = 4000
'PLeft(4) = 5500
'End Sub
'
'Sub PrintTitle()
'GetPleft
'iLine1 = 1
'
'Printer.Font.Size = 16
'Printer.Font.Underline = False
'Printer.FontBold = False
'
'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("匯入錯誤訊息") / 2)
'Printer.CurrentY = iLine1 * 300
'Printer.Print "匯入錯誤訊息"
'
'Printer.Font.Size = 12
'Printer.Font.Underline = False
'Printer.FontBold = False
'
'iLine1 = iLine1 + 1
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = 900
'Printer.Print "列印人員：" & strUserName
'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'Printer.CurrentY = 900
'Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'iLine1 = iLine1 + 1
'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'Printer.CurrentY = 1200
'Printer.Print "頁　　次：" & Printer.Page
'
'iLine1 = 5
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iLine1 * 300
'Printer.Print "錯誤訊息"
'
'iLine1 = iLine1 + 1
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iLine1 * 300
'Printer.Print String(148, "-")
'iLine1 = iLine1 + 1
'End Sub
'
'Sub PrintDetail()
'Dim m_j As Integer
'   For m_j = 1 To 1
'      Printer.CurrentX = PLeft(m_j)
'      Printer.CurrentY = iLine1 * 300
'      Printer.Print strTemp(m_j)
'   Next m_j
'   iLine1 = iLine1 + 1
'End Sub

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
