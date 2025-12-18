VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140421 
   BorderStyle     =   1  '單線固定
   Caption         =   "整批匯入為潛在客戶"
   ClientHeight    =   7710
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   10410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10410
   Begin VB.OptionButton Option1 
      Caption         =   "國內"
      Height          =   195
      Index           =   0
      Left            =   1425
      TabIndex        =   4
      Top             =   1845
      Width           =   705
   End
   Begin VB.OptionButton Option1 
      Caption         =   "國外"
      Height          =   195
      Index           =   1
      Left            =   2145
      TabIndex        =   5
      Top             =   1845
      Width           =   705
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00800080&
      Height          =   1605
      Left            =   3660
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frm140421.frx":0000
      Top             =   5550
      Width           =   3675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   345
      Left            =   3900
      TabIndex        =   7
      Top             =   1920
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   315
      Left            =   9510
      TabIndex        =   1
      Top             =   510
      Width           =   375
   End
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1410
      TabIndex        =   0
      Top             =   480
      Width           =   8055
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯入(&T)"
      Height          =   345
      Left            =   7050
      TabIndex        =   6
      Top             =   90
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   8130
      TabIndex        =   8
      Top             =   90
      Width           =   885
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
      Height          =   3435
      Left            =   30
      TabIndex        =   16
      Top             =   2070
      Width           =   10335
      Begin VB.ListBox List1 
         Height          =   3100
         ItemData        =   "frm140421.frx":0071
         Left            =   60
         List            =   "frm140421.frx":0073
         TabIndex        =   17
         Top             =   210
         Width           =   10215
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00800080&
      Height          =   1635
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frm140421.frx":0075
      Top             =   5550
      Width           =   3195
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   30
      TabIndex        =   12
      Top             =   7170
      Width           =   10335
      Begin VB.TextBox TextCnt 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   30
         TabIndex        =   13
         Top             =   150
         Width           =   10260
      End
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   60
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   705
      Left            =   2040
      TabIndex        =   10
      Top             =   30
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
   Begin VB.FileListBox File1 
      Height          =   420
      Left            =   1410
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   810
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox txtPCC 
      Height          =   945
      Index           =   13
      Left            =   6390
      TabIndex        =   3
      Top             =   840
      Width           =   3735
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "6588;1667"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPCU 
      Height          =   945
      Index           =   40
      Left            =   1410
      TabIndex        =   2
      Top             =   840
      Width           =   3735
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6588;1667"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人備註："
      Height          =   210
      Index           =   1
      Left            =   5280
      TabIndex        =   23
      Top             =   900
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   210
      Index           =   19
      Left            =   840
      TabIndex        =   22
      Top             =   900
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢權限：                                     部門"
      Height          =   180
      Index           =   39
      Left            =   330
      TabIndex        =   21
      Top             =   1845
      Width           =   2985
   End
   Begin MSForms.ComboBox cboPCU11 
      Height          =   285
      Left            =   3420
      TabIndex        =   19
      Top             =   30
      Visible         =   0   'False
      Width           =   1710
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3016;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "類別:"
      Height          =   285
      Index           =   0
      Left            =   2970
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Excel存放路徑："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   540
      Width           =   1290
   End
End
Attribute VB_Name = "frm140421"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/17 Form2.0已修改
'Create By Sindy 2021/6/4
Option Explicit

Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
Dim i As Integer, j As Integer
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iNowLine As Integer 'Add Sindy 2021/1/28
Dim iRowHeight As Integer 'Add Sindy 2021/1/28


'回傳Excel欄位
Private Function GetFieldStr2(ByVal intAdd As Integer, ByVal intField As Integer) As String
    Dim intDiv As Integer, intMod As Integer
    
    GetFieldStr2 = ""
    If intAdd + intField > 90 Then
        intDiv = ((intAdd + intField) - 65) \ 26
        intMod = (intAdd + intField) - (64 + (intDiv * 26))
        GetFieldStr2 = Chr(intDiv + 64) & Chr(intMod + 64)
    Else
        GetFieldStr2 = Chr(intAdd + intField)
    End If
End Function

Private Sub cmdExcel_Click()
Dim xlsSalesPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim stFileName As String
Dim intMaxRow As Integer
Dim dblMaxWidth As Double
Dim intRow As Integer
Dim bolConn As Boolean
Dim intErrRow As Integer
Dim ii As Integer, strExCol As String
Dim strPCU(1 To 51) As String
Dim strNA04 As String, strCT02 As String, strNo As String
Dim stCols As String, stValues As String
Dim intPCC02 As Integer, strPCC03 As String, strPCC08 As String
Dim strPCC13 As String
Dim strTmp As String, strTmp1 As String, intPos As Integer
Dim strNA01 As String 'Add By Sindy 2021/8/3
Dim rsTmp As New ADODB.Recordset
Dim strName As String
Dim intColSetCnt As Integer '聯絡人是幾個欄位一組
Dim intXLSCol As Integer '第一組聯絡人開始的前一個Excel欄位;70=F,71=G
'Dim bolSave As Boolean
   
On Error GoTo flgErr
   
   intErrRow = 0
   stFileName = txtPath1
   If Dir(stFileName) = "" Then
      MsgBox "請選擇一個Excel檔案！"
      Exit Sub
   End If
   
'   If txtPCU(40) = "" Then
'      MsgBox "備註欄位不可空白！"
'      Exit Sub
'   End If
   
   If Option1(0).Value = False And Option1(1).Value = False Then
      MsgBox "請勾選一種查詢權限！"
      Exit Sub
   End If
   
   '開檔
   Screen.MousePointer = vbHourglass
   xlsSalesPoint.Workbooks.Open stFileName
   'xlsSalesPoint.Visible = True
   Set wksrpt = xlsSalesPoint.Worksheets(1)
   '把Excel的警告訊息關掉
   xlsSalesPoint.DisplayAlerts = False
   
   '檢查標題欄
   intRow = 1
   If Trim(wksrpt.Range("A" & intRow).Value) <> "" Then
      '[固定欄位]
      If Trim(wksrpt.Range("A" & intRow).Value) <> "開發人員" Then
         MsgBox "A 欄位必須是「開發人員」！"
         GoTo RunExit
      End If
      If Trim(wksrpt.Range("B" & intRow).Value) <> "性質" Then
         MsgBox "B 欄位必須是「性質」！"
         GoTo RunExit
      End If
      If Trim(wksrpt.Range("C" & intRow).Value) <> "國籍" Then
         MsgBox "C 欄位必須是「國籍」！"
         GoTo RunExit
      End If
      If Trim(wksrpt.Range("D" & intRow).Value) <> "指定匯入編號" Then
         MsgBox "D 欄位必須是「指定匯入編號」！"
         GoTo RunExit
      End If
      If Trim(wksrpt.Range("E" & intRow).Value) <> "名稱" Then
         MsgBox "E 欄位必須是「名稱」！"
         GoTo RunExit
      End If
      If UCase(Trim(wksrpt.Range("F" & intRow).Value)) <> UCase("Email代表號") Then
         MsgBox "F 欄位必須是「Email代表號」！"
         GoTo RunExit
      End If
      'Add By Sindy 2021/12/15
      If UCase(Trim(wksrpt.Range("G" & intRow).Value)) <> UCase("備註") Then
         MsgBox "G 欄位必須是「備註」！"
         GoTo RunExit
      End If
      '2021/12/15 END
      
      intColSetCnt = 3 '聯絡人是幾個欄位一組
      intXLSCol = 71 '第一組聯絡人開始的前一個Excel欄位;70=F,71=G
      '[變動欄位]
      For ii = 1 To 100 Step intColSetCnt
         strExCol = GetFieldStr2(intXLSCol, ii)
         If Trim(wksrpt.Range(strExCol & intRow).Value) <> "" Then
            If Not (InStr(Trim(wksrpt.Range(strExCol & intRow).Value), "聯絡人") > 0 And _
                  InStr(UCase(Trim(wksrpt.Range(strExCol & intRow).Value)), UCase("Email")) = 0) Then
               MsgBox strExCol & " 欄位必須是「聯絡人名稱」！"
               GoTo RunExit
            End If
            strExCol = GetFieldStr2(intXLSCol, ii + 1)
            If Not (InStr(Trim(wksrpt.Range(strExCol & intRow).Value), "聯絡人") > 0 And _
                  InStr(UCase(Trim(wksrpt.Range(strExCol & intRow).Value)), UCase("Email")) > 0) Then
               MsgBox strExCol & "欄位必須是「聯絡人Email」！"
               GoTo RunExit
            End If
            'Add By Sindy 2021/12/15
            strExCol = GetFieldStr2(intXLSCol, ii + 2)
            If Not (InStr(Trim(wksrpt.Range(strExCol & intRow).Value), "聯絡人") > 0 And _
                  InStr(UCase(Trim(wksrpt.Range(strExCol & intRow).Value)), UCase("備註")) > 0) Then
               MsgBox strExCol & "欄位必須是「聯絡人備註」！"
               GoTo RunExit
            End If
            '2021/12/15 END
         Else
            Exit For
         End If
      Next ii
      If ii >= 99 Then
         strExCol = GetFieldStr2(intXLSCol, ii)
         If Trim(wksrpt.Range(strExCol & intRow).Value) <> "" Then
            MsgBox "變動欄位已超過程式設計的數量，請洽電腦中心！"
            GoTo RunExit
         End If
      End If
   Else
      MsgBox "此Excel檔案，無內容可讀取！"
      GoTo RunExit
   End If
   
   Load frmpic002
   frmpic002.Label1.Caption = "檢查資料中...請稍候..."
   frmpic002.Show
   frmpic002.ZOrder 0
   List1.Clear
   
   '檢查明細資料
   intRow = 1
   Do While Trim(wksrpt.Range("A" & (intRow + 1)).Value) <> ""
      intRow = intRow + 1
      '[固定欄位]
      If Trim(wksrpt.Range("A" & intRow).Value) = "" Then
         List1.AddItem "A 欄第 " & intRow & " 列, 開發人員不可空白"
         wksrpt.Range("A" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
      End If
      
      If Trim(wksrpt.Range("D" & intRow).Value) <> "" Then '有指定匯入編號
         If Len(Trim(wksrpt.Range("D" & intRow).Value)) < 9 Then
            strNo = wksrpt.Range("D" & intRow).Value
            wksrpt.Range("D" & intRow).Value = ChangeCustomerL(strNo) '長度9碼
         End If
      End If
      
      'Modify By Sindy 2022/2/9 無輸入指定匯入編號,才檢查性質,國籍
      If Trim(wksrpt.Range("D" & intRow).Value) = "" Or _
         Len(Trim(wksrpt.Range("D" & intRow).Value)) <> 9 Then
      '2022/2/9 END
         
         If Trim(wksrpt.Range("B" & intRow).Value) = "" Then
            List1.AddItem "B 欄第 " & intRow & " 列, 性質不可空白"
            wksrpt.Range("B" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
         End If
         If InStr(國外潛在客戶類別, Trim(wksrpt.Range("B" & intRow).Value)) = 0 Then
            List1.AddItem "B 欄第 " & intRow & " 列, 無此性質"
            wksrpt.Range("B" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
         End If
         
         strNA01 = ""
         If Trim(wksrpt.Range("C" & intRow).Value) = "" Then
            List1.AddItem "C 欄第 " & intRow & " 列, 國籍不可空白"
            wksrpt.Range("C" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
         Else
            strSql = "select * from nation where na03='" & Trim(wksrpt.Range("C" & intRow).Value) & "' order by na01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
               List1.AddItem "C 欄第 " & intRow & " 列, 無此國籍"
               wksrpt.Range("C" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
            Else
               strNA01 = RsTemp.Fields("na01")
               '城市
               strSql = "SELECT * FROM City WHERE ct01='" & Left(strNA01, 3) & "' AND upper(ct03)='NONE'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
                  strExc(0) = "select lpad(nvl(max(ct02),0)+1,3,'0') from city where ct01='" & Left(strNA01, 3) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strSql = "insert into City(CT01,CT02,CT03) values('" & Left(strNA01, 3) & "','" & RsTemp(0) & "','NONE')"
                     cnnConnection.Execute strSql
                  End If
               End If
            End If
         End If
      End If
      
      If Trim(wksrpt.Range("D" & intRow).Value) <> "" Then '有指定匯入編號
'         If Len(Trim(wksrpt.Range("D" & intRow).Value)) < 9 Then
'            strNo = wksrpt.Range("D" & intRow).Value
'            wksrpt.Range("D" & intRow).Value = ChangeCustomerL(strNo) '長度9碼
'         End If
         strPCU(1) = Mid(wksrpt.Range("D" & intRow).Value, 1, 8)
         strPCU(2) = Mid(wksrpt.Range("D" & intRow).Value, 9, 1)
         '[變動欄位]
         intPCC02 = 0
         For ii = 1 To 100 Step intColSetCnt
            strExCol = GetFieldStr2(intXLSCol, ii)
            If Trim(wksrpt.Range(strExCol & intRow).Value) <> "" Then
               strPCC03 = ChgSQL(Trim(wksrpt.Range(strExCol & intRow).Value)) '聯絡人名稱
               '檢查聯絡人是否存在,存在不更新資料,欄位變色
               strSql = "SELECT pcc01,pcc02" & _
                         " FROM potcustcont" & _
                        " WHERE pcc01='" & strPCU(1) & "'" & _
                        " and (rtrim(pcc05)=rtrim('" & strPCC03 & "')" & _
                             " or rtrim(upper(pcc03))=rtrim('" & UCase(strPCC03) & "')" & _
                             ")"
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount = 1 Then
                  List1.AddItem strExCol & " 欄第 " & intRow & " 列, 聯絡人已存在"
                  wksrpt.Range(strExCol & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
               End If
               rsTmp.Close
            Else
               Exit For
            End If
         Next ii
      Else
         If Trim(wksrpt.Range("E" & intRow).Value) = "" Then
            List1.AddItem "E 欄第 " & intRow & " 列, 名稱不可空白"
            wksrpt.Range("E" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
         Else
            '檢查名稱是否已存在
            strName = Trim(wksrpt.Range("E" & intRow).Value)
            strExc(10) = ChkCustNameAndPotCust_21(strName, strNA01)
            
            If strExc(10) <> "" Then
               'Add By Sindy 2021/8/11
               If InStr(strExc(10), ":對造") > 0 Then
                  List1.AddItem "E 欄第 " & intRow & " 列, 名稱:" & strName & " " & strExc(10)
                  wksrpt.Range("E" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
               'Add By Sindy 2021/8/27
               ElseIf InStr(strExc(10), ":不得代理") > 0 Then
                  List1.AddItem "E 欄第 " & intRow & " 列, 名稱:" & strName & " " & strExc(10)
                  wksrpt.Range("E" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
               'Add By Sindy 2021/8/31
               ElseIf InStr(strExc(10), ":") > 0 Then
                  List1.AddItem "E 欄第 " & intRow & " 列, 名稱:" & strName & " " & strExc(10)
                  wksrpt.Range("E" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
               ElseIf InStr(strExc(10), ",") > 0 Then
                  List1.AddItem "E 欄第 " & intRow & " 列, 名稱已存在。" & strExc(10)
                  wksrpt.Range("E" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
               Else
                  wksrpt.Range("D" & intRow).Value = strExc(10) '指定匯入編號
                  wksrpt.Range("D" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
               End If
               '2021/8/11 END
            Else
               '檢查後面截取名稱是否正確
               If Option1(1).Value = True Then '國外權限,英文名稱
                  strPCU(3) = "": strPCU(4) = "": strPCU(5) = "": strPCU(6) = ""
                  strTmp = Trim(wksrpt.Range("E" & intRow).Value)
                  For ii = 3 To 6
                     If strTmp <> "" Then
                        strTmp1 = Mid(strTmp, 1, 30)
                        If Right(strTmp1, 1) = " " Or Len(strTmp) <= 30 Or (Len(strTmp) > 30 And Mid(strTmp, 31, 1) = " ") Then
                           strPCU(ii) = Trim(strTmp1)
                           If Len(strTmp) > 30 Then
                              strTmp = Trim(Mid(strTmp, 31))
                           Else
                              strTmp = ""
                           End If
                        Else
                           intPos = InStrRev(strTmp1, " ")
                           strPCU(ii) = Trim(Mid(strTmp1, 1, intPos))
                           strTmp = Trim(Mid(strTmp, intPos))
                        End If
                     Else
                        Exit For
                     End If
                  Next ii
                  If Trim(wksrpt.Range("E" & intRow).Value) <> Trim(strPCU(3) & " " & strPCU(4) & " " & strPCU(5) & " " & strPCU(6)) Then
                     List1.AddItem "E 欄第 " & intRow & " 列, 名稱截取有問題:" & Trim(strPCU(3) & " " & strPCU(4) & " " & strPCU(5) & " " & strPCU(6))
                     wksrpt.Range("E" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
                  End If
               End If
            End If
         End If
      End If
      
      If UCase(Trim(wksrpt.Range("F" & intRow).Value)) = "" Then
'         List1.AddItem "F 欄第 " & intRow & " 列, Email代表號不可空白"
'         wksrpt.Range("F" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
      Else
         '檢查E-Mail正確性
         If PUB_CheckMail(UCase(Trim(wksrpt.Range("F" & intRow).Value)), False, strExc(10)) = False Then
            List1.AddItem "F 欄第 " & intRow & " 列, " & strExc(10)
            wksrpt.Range("F" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
         End If
      End If
      
      '[變動欄位]
      For ii = 1 To 100 Step intColSetCnt
         strExCol = GetFieldStr2(intXLSCol, ii)
         If Trim(wksrpt.Range(strExCol & intRow).Value) <> "" Then
'            If Trim(wksrpt.Range(strExCol & intRow).Value) = "" Then
'               List1.AddItem strExCol & " 欄第 " & intRow & " 列, 聯絡人名稱不可空白"
'               wksrpt.Range(strExCol & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
'            End If
            strExCol = GetFieldStr2(intXLSCol, ii + 1)
            If Trim(wksrpt.Range(strExCol & intRow).Value) = "" Then
'               List1.AddItem strExCol & " 欄第 " & intRow & " 列, 聯絡人Email不可空白"
'               wksrpt.Range(strExCol & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
            End If
         Else
            Exit For
         End If
      Next ii
   Loop
   intMaxRow = intRow - 1 '總筆數
   If intMaxRow <= 0 Then
      MsgBox "此Excel檔案，無內容可讀取！"
      GoTo RunExit
   End If
   Unload frmpic002
   
   '無錯誤,即可匯入資料
   If List1.ListCount = 0 Then
      dblMaxWidth = 8820
      TextCnt.Width = 0
      cnnConnection.BeginTrans: bolConn = True
      For intRow = 2 To intMaxRow + 1 'Excel列欄數
         TextCnt.Width = dblMaxWidth / intMaxRow * (intRow - 1) '筆數
         
         For ii = 1 To 51
            strPCU(ii) = ""
         Next ii
         stCols = "": stValues = ""
         
         If Option1(0).Value = True Then '國內權限,中文名稱
            strPCU(8) = Trim(wksrpt.Range("E" & intRow).Value): stCols = stCols & ",PCU08": stValues = stValues & "," & CNULL(ChgSQL(strPCU(8)))
         Else '國外權限,英文名稱
            strTmp = Trim(wksrpt.Range("E" & intRow).Value)
            For ii = 3 To 6
               If strTmp <> "" Then
                  strTmp1 = Mid(strTmp, 1, 30)
                  If Right(strTmp1, 1) = " " Or Len(strTmp) <= 30 Or (Len(strTmp) > 30 And Mid(strTmp, 31, 1) = " ") Then
                     strPCU(ii) = Trim(strTmp1)
                     If Len(strTmp) > 30 Then
                        strTmp = Trim(Mid(strTmp, 31))
                     Else
                        strTmp = ""
                     End If
                  Else
                     intPos = InStrRev(strTmp1, " ")
                     strPCU(ii) = Trim(Mid(strTmp1, 1, intPos))
                     strTmp = Trim(Mid(strTmp, intPos))
                  End If
                  stCols = stCols & ",PCU" & Format(ii, "00"): stValues = stValues & "," & CNULL(ChgSQL(strPCU(ii)))
               Else
                  Exit For
               End If
            Next ii
         End If
         
         '國籍
         strPCU(9) = Trim(wksrpt.Range("C" & intRow).Value)
         '轉換代碼
         strSql = "select * from nation where na03='" & strPCU(9) & "' order by na01"
         intI = 1
         strNA04 = ""
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strPCU(9) = pub_NationByName(strPCU(3) & " " & strPCU(4) & " " & strPCU(5) & " " & strPCU(6), RsTemp.Fields("na01"), False, "客戶")
            strNA04 = "" & RsTemp.Fields("na04")
         End If
         strSql = "SELECT * FROM City WHERE ct01='" & Left(strPCU(9), 3) & "' AND upper(ct03)='NONE'"
         intI = 1
         strCT02 = ""
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strCT02 = "" & RsTemp.Fields("CT02")
         End If
         stCols = stCols & ",PCU09": stValues = stValues & "," & CNULL(strPCU(9))
         strPCU(10) = strCT02 '城市:NONE
         stCols = stCols & ",PCU10": stValues = stValues & "," & CNULL(strPCU(10))
         strPCU(20) = "None" '英文地址1
         stCols = stCols & ",PCU20": stValues = stValues & "," & CNULL(strPCU(20))
         strPCU(21) = strNA04 '英文地址2
         stCols = stCols & ",PCU21": stValues = stValues & "," & CNULL(strPCU(21))
         strPCU(28) = Left(strPCU(9), 3) '地址國籍
         stCols = stCols & ",PCU28": stValues = stValues & "," & CNULL(strPCU(28))
         '非台灣,非大陸者
         If Not ((Left(strPCU(9), 3) >= "000" And Left(strPCU(9), 3) <= "008") Or _
               Left(strPCU(9), 3) = "020") Then
            strPCU(35) = "N" '是否寄電子報
            strPCU(48) = "N" '是否寄專利雙週報
         Else
            strPCU(35) = ""
            strPCU(48) = ""
         End If
         stCols = stCols & ",PCU35": stValues = stValues & "," & CNULL(strPCU(35))
         stCols = stCols & ",PCU48": stValues = stValues & "," & CNULL(strPCU(48))
         
         '類別=性質
         strPCU(11) = Trim(wksrpt.Range("B" & intRow).Value)
         '轉換代碼
         For ii = 0 To cboPCU11.ListCount - 1
            If InStr(cboPCU11.List(ii), strPCU(11)) > 0 Then
               strPCU(11) = Left(cboPCU11.List(ii), 1)
               Exit For
            End If
         Next ii
         stCols = stCols & ",PCU11": stValues = stValues & "," & CNULL(strPCU(11))
         
         strPCU(18) = Trim(wksrpt.Range("F" & intRow).Value) 'E-Mail
         stCols = stCols & ",PCU18": stValues = stValues & "," & CNULL(ChgSQL(strPCU(18)))
         strPCU(37) = strSrvDate(1) '開發日期
         stCols = stCols & ",PCU37": stValues = stValues & "," & CNULL(strPCU(37), True)
         
         '開發人員
         strPCU(38) = Trim(wksrpt.Range("A" & intRow).Value)
         '轉換代碼
         strPCU(38) = GetPrjSalesNM_2(CStr(strPCU(38)))
         stCols = stCols & ",PCU38": stValues = stValues & "," & CNULL(strPCU(38))
         
         '備註
         'Add By Sindy 2021/12/15 可輸入各自的備註
         If Trim(wksrpt.Range("G" & intRow).Value) <> "" Then
            strPCU(40) = Trim(wksrpt.Range("G" & intRow).Value)
            stCols = stCols & ",PCU40": stValues = stValues & "," & CNULL(strPCU(40))
         Else
         '2021/12/15 END
            If Trim(txtPCU(40)) <> "" Then
               strPCU(40) = Trim(txtPCU(40))
               stCols = stCols & ",PCU40": stValues = stValues & "," & CNULL(strPCU(40))
            End If
         End If
         
         '國外權限
         If Option1(0).Value = True Then
            strPCU(51) = "C" '國內
         Else
            strPCU(51) = "F" '國外
         End If
         stCols = stCols & ",PCU51": stValues = stValues & "," & CNULL(strPCU(51))
         
   '更新資料:
         '*****************************************************************
         'Modify By Sindy 2021/8/11
          '有編號,僅新增聯絡人資料
         If Mid(wksrpt.Range("D" & intRow).Value, 1, 1) = "R" Or _
            Mid(wksrpt.Range("D" & intRow).Value, 1, 1) = "X" Or _
            Mid(wksrpt.Range("D" & intRow).Value, 1, 1) = "Y" Then
            strPCU(1) = Mid(wksrpt.Range("D" & intRow).Value, 1, 8)
            strPCU(2) = Mid(wksrpt.Range("D" & intRow).Value, 9, 1)
            
            '[變動欄位]
            intPCC02 = 0
            '目前最大的聯絡人編號
            strSql = "SELECT nvl(max(pcc02),0)" & _
                      " FROM potcustcont" & _
                     " WHERE pcc01='" & strPCU(1) & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               intPCC02 = rsTmp.Fields(0)
            End If
            rsTmp.Close
            For ii = 1 To 100 Step intColSetCnt
               strExCol = GetFieldStr2(intXLSCol, ii)
               If Trim(wksrpt.Range(strExCol & intRow).Value) <> "" Then
                  strPCC03 = ChgSQL(Trim(wksrpt.Range(strExCol & intRow).Value))
'                  '檢查聯絡人是否存在,存在不更新資料,欄位變色
'                  strSql = "SELECT pcc01,pcc02" & _
'                            " FROM potcustcont" & _
'                           " WHERE pcc01='" & strPCU(1) & "'" & _
'                           " and (rtrim(pcc05)=rtrim('" & strPCC03 & "')" & _
'                                " or rtrim(upper(pcc03))=rtrim('" & UCase(strPCC03) & "')" & _
'                                ")"
'                  rsTmp.CursorLocation = adUseClient
'                  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsTmp.RecordCount = 0 Then
                     strExCol = GetFieldStr2(intXLSCol, ii + 1)
                     strPCC08 = ChgSQL(Trim(wksrpt.Range(strExCol & intRow).Value))
                     
                     'Add By Sindy 2021/12/15 可輸入各自的聯絡人備註
                     strExCol = GetFieldStr2(intXLSCol, ii + 2)
                     strPCC13 = Trim(wksrpt.Range(strExCol & intRow).Value)
                     If Trim(strPCC13) <> "" Then
                        strPCC13 = ChgSQL(strPCC13)
                     Else
                        If Trim(txtPCC(13)) <> "" Then
                           strPCC13 = ChgSQL(Trim(txtPCC(13)))
                        End If
                     End If
                     '2021/12/15 END
                     
                     intPCC02 = intPCC02 + 1
                     'Modify By Sindy 2021/10/6 + ,PCC13
                     strSql = "INSERT INTO PotCustCont (PCC01,PCC02,PCC03,PCC08,PCC10,PCC23,PCC24,PCC11,PCC12,PCC13)" & _
                        " Values (" & CNULL(strPCU(1)) & "," & CNULL(Format(intPCC02, "00")) & "," & CNULL(strPCC03) & _
                        "," & CNULL(strPCC08) & ",'Y'," & CNULL(strPCU(48)) & _
                        "," & CNULL(strPCU(48)) & "," & CNULL(strPCU(37), True) & "," & CNULL(strPCU(38)) & "," & CNULL(strPCC13) & ")"
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql
'                  Else
'                     wksrpt.Range("G" & intRow).Interior.ColorIndex = 43 '設置儲存格填充色(綠)
'                     bolSave = True
'                  End If
'                  rsTmp.Close
                     
               Else
                  Exit For
               End If
            Next ii
          
         '無編號,新增一筆國外潛在客戶
         Else
            If ClsPDGetAutoNumber("R", strNo, True, False) Then
               strPCU(1) = "R" + Right(strNo, 5) & "00"
               stCols = stCols & ",PCU01": stValues = stValues & "," & CNULL(strPCU(1))
               strPCU(2) = "0"
               stCols = stCols & ",PCU02": stValues = stValues & "," & CNULL(strPCU(2))
            End If
            'Add By Sindy 2021/8/25
            wksrpt.Range("D" & intRow).Value = strPCU(1) & strPCU(2)
            wksrpt.Range("D" & intRow).Interior.ColorIndex = 44 '設置儲存格填充色(淺橘黃)
            '2021/8/25 END
            stCols = Mid(stCols, 2)
            stValues = Mid(stValues, 2)
            strSql = "INSERT INTO PotCustomer (" & stCols & ") Values (" & stValues & ")"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
            
            'Add By Sindy 2021/10/4 往下筆檢查是否有相同的客戶名稱+國籍(無編號的)
            Dim jj As Integer, strNation As String
'楊雯芳   事務所   英國  R18455000   MARKS & CLERK LLP    Daniel Thorpe  dthorpe@marks-clerk.com
'楊雯芳   事務所   英國  R18456000   MARKS & CLERK LLP    Robert Lind rlind@marks-clerk.com
'楊雯芳   事務所   英國  R18457000   MARKS & CLERK LLP    Matthew Pinney mpinney@marks-clerk.com
'楊雯芳   事務所   英國  R18458000   MARKS & CLERK LLP    Clare Smith clsmith@marks-clerk.com
            strNation = wksrpt.Range("C" & intRow).Value
            strName = wksrpt.Range("E" & intRow).Value
            For jj = intRow + 1 To intMaxRow + 1
               If wksrpt.Range("D" & jj).Value = "" And _
                  strNation = wksrpt.Range("C" & jj).Value And _
                  strName = wksrpt.Range("E" & jj).Value Then
                  wksrpt.Range("D" & jj).Value = strPCU(1) & strPCU(2)
                  wksrpt.Range("D" & jj).Interior.ColorIndex = 44 '設置儲存格填充色(淺橘黃)"
               End If
            Next jj
            '2021/10/4 END
            
            '[變動欄位]
            intPCC02 = 0
            For ii = 1 To 100 Step intColSetCnt
               strExCol = GetFieldStr2(intXLSCol, ii)
               If Trim(wksrpt.Range(strExCol & intRow).Value) <> "" Then
                  strPCC03 = ChgSQL(Trim(wksrpt.Range(strExCol & intRow).Value))
                  strExCol = GetFieldStr2(intXLSCol, ii + 1)
                  strPCC08 = ChgSQL(Trim(wksrpt.Range(strExCol & intRow).Value))
                  
                  'Add By Sindy 2021/12/15 可輸入各自的聯絡人備註
                  strExCol = GetFieldStr2(intXLSCol, ii + 2)
                  strPCC13 = Trim(wksrpt.Range(strExCol & intRow).Value)
                  If Trim(strPCC13) <> "" Then
                     strPCC13 = ChgSQL(strPCC13)
                  Else
                     If Trim(txtPCC(13)) <> "" Then
                        strPCC13 = ChgSQL(Trim(txtPCC(13)))
                     End If
                  End If
                  '2021/12/15 END
                  
                  intPCC02 = intPCC02 + 1
                  'Modify By Sindy 2021/10/6 + ,PCC13
                  strSql = "INSERT INTO PotCustCont (PCC01,PCC02,PCC03,PCC08,PCC10,PCC23,PCC24,PCC11,PCC12,PCC13)" & _
                     " Values (" & CNULL(strPCU(1)) & "," & CNULL(Format(intPCC02, "00")) & "," & CNULL(strPCC03) & _
                     "," & CNULL(strPCC08) & ",'Y'," & CNULL(strPCU(48)) & _
                     "," & CNULL(strPCU(48)) & "," & CNULL(strPCU(37), True) & "," & CNULL(strPCU(38)) & "," & CNULL(strPCC13) & ")"
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql
               Else
                  Exit For
               End If
            Next ii
         End If
      Next intRow
      
      cnnConnection.CommitTrans: bolConn = False
      
      TextCnt.Width = dblMaxWidth: DoEvents
      
      'Add By Sindy 2021/8/25
      '判斷若版本2007以上改變存格式
      If Val(xlsSalesPoint.Version) < 12 Then
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & "_OK.xls", FileFormat:=-4143
      Else
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & "_OK.xls", FileFormat:=56
      End If
      '2021/8/25 END
      
      MsgBox "資料匯入完畢！ " & vbCrLf & vbCrLf & _
             "(共計 " & intMaxRow & " 筆)" ', 錯誤有 " & intErrRow & " 筆
   End If
   
   If List1.ListCount > 0 Then 'Or bolSave = True
      '另存
      '.SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & ".xls"
      'xlsSalesPoint.Workbooks(1).Save 'FileName:=stFileName, FileFormat:=56
      '關閉相容性檢查的對話框
      xlsSalesPoint.Workbooks(1).CheckCompatibility = False
      'xlsSalesPoint.Workbooks(1).DoNotPromptForConvert = True
      '判斷若版本2007以上改變存格式
      If Val(xlsSalesPoint.Version) < 12 Then
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & ".xls", FileFormat:=-4143
      Else
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & ".xls", FileFormat:=56
      End If
'      If bolSave = True Then
'         MsgBox "有編號聯絡人資料已存在！"
'      Else
         MsgBox "資料有誤, 請修改後再重新匯入！"
'      End If
   End If
   
RunExit:
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit '離開
   Set wksrpt = Nothing
   Set xlsSalesPoint = Nothing
   Screen.MousePointer = vbDefault
   
   Set rsTmp = Nothing
   Exit Sub
   
flgErr:
   Set rsTmp = Nothing
   If TypeName(frmpic002) <> "Nothing" Then Unload frmpic002
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   
   '另存
   '判斷若版本2007以上改變存格式
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & ".xls", FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=Mid(stFileName, 1, Len(stFileName) - 4) & strSrvDate(2) & ServerTime & ".xls", FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit '離開
   Set wksrpt = Nothing
   Set xlsSalesPoint = Nothing
   If Err.Number <> 0 Then
       MsgBox intRow & " 筆 : " & Err.Description
   End If
End Sub

'比對國內外潛在客戶名稱相同者
'Modify By Sindy 2021/8/3
Private Function ChkCustNameAndPotCust_21(strPCUName As String, strNA01 As String) As String
Dim rsTmp As New ADODB.Recordset
Dim strCustID As String
   
   If Trim(strPCUName) = "" Then Exit Function
   ChkCustNameAndPotCust_21 = ""
   
   strNA01 = Left(strNA01, 3)
   strCustID = ""
   
   '檢查對造
   Dim strCheckWay As String
   Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
   Dim strTp(3) As String
   
   strTp(3) = ChgSQL(UCase(Trim(strPCUName)))
   strSQL1 = " AND CP01 IN(" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
   strSQL2 = " AND CP01 IN(" & SQLGrpStr("", 1) & ") "
   StrSQL3 = " AND CP01 IN(" & SQLGrpStr("", 3) & ") "
   StrSQL4 = " AND CP01 IN(" & SQLGrpStr("", 4) & ") "
   strSQL5 = " AND CP01 IN(" & SQLGrpStr("", 5) & ") "
   'strCheckWay = ">0" '模糊比對
   'strCheckWay = "=1" '字首比對
   strCheckWay = "=" '完全比對
   '*****
'   Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, strTp(3), strCheckWay, True)
   'Modify By Sindy 2023/5/8
   '刪除對造暫存檔資料
   cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' "
   '改使用另外一個共用函數
   strSql = GetSearchNameSql(Me.Name, strTp(3), strCheckWay, True, True, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, strNA01)
   '*****
'   strSql = "select * from R100102_1 where id='" & strUserNum & "@" & Me.Name & "' And R021004='1'"
   '2023/5/8 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ChkCustNameAndPotCust_21 = ":對造"
      Set rsTmp = Nothing
      Exit Function
   End If
   rsTmp.Close
   
   'Add By Sindy 2021/8/31 檢查國內開拓函特定公司不列印者(TMBulletinNp)之特定公司(商標權人)
   strSql = "SELECT TBNP01" & _
             " FROM TMBulletinNp" & _
            " WHERE rtrim(TBNP01)=rtrim('" & ChgSQL(strPCUName) & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         If IsNull(rsTmp.Fields(0)) = False Then
            strCustID = strCustID & "," & rsTmp.Fields(0)
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   If strCustID <> "" Then
      ChkCustNameAndPotCust_21 = ":國內開拓函特定公司不列印者"
      Set rsTmp = Nothing
      Exit Function
   End If
   
'   '名稱是否存在, 改以名稱+國籍 做檢查; 不用檢查日文
'   '選擇國內的時候, 才需要檢查國內潛在客戶
'   '順序一樣 R, Y, X
'   strNA01 = Left(strNA01, 3)
'   strCustID = ""
'
''********
'   'Add By Sindy 2021/8/27 檢查不得代理案件(NotAgent)之客戶或代理人
'   '比對中文名稱
'   strSql = "SELECT nt01" & _
'             " FROM NotAgent" & _
'            " WHERE rtrim(nt02)=rtrim('" & ChgSQL(strPCUName) & "') and substr(nt08,1,3)='" & strNA01 & "'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'   '比對英文名稱
'   strSql = "SELECT nt01" & _
'             " FROM NotAgent" & _
'            " WHERE rtrim(upper(nt03||' '||nt04||' '||nt05||' '||nt06))=rtrim('" & ChgSQL(UCase(Trim(strPCUName))) & "')" & _
'            " and substr(nt08,1,3)='" & strNA01 & "'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'   If strCustID <> "" Then
'      ChkCustNameAndPotCust_21 = Mid(strCustID, 2) & ":不得代理"
'      Exit Function
'   End If
'   '2021/8/27 END
''********
'
'   'Add By Sindy 2021/8/31 檢查投資法務開拓客戶資料檔(ExpandCusDetail)之收件人1,收件人2,公司名稱1,公司名稱2
'   strSql = "SELECT ECD03,ECD04,ECD11,ECD12" & _
'             " FROM ExpandCusDetail" & _
'            " WHERE rtrim(ECD03)=rtrim('" & ChgSQL(strPCUName) & "') or rtrim(ECD04)=rtrim('" & ChgSQL(strPCUName) & "') " & _
'                "or rtrim(ECD11)=rtrim('" & ChgSQL(strPCUName) & "') or rtrim(ECD12)=rtrim('" & ChgSQL(strPCUName) & "')"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         ElseIf IsNull(rsTmp.Fields(1)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(1)
'         ElseIf IsNull(rsTmp.Fields(2)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(2)
'         ElseIf IsNull(rsTmp.Fields(3)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(3)
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'   If strCustID <> "" Then
'      ChkCustNameAndPotCust_21 = ":投資法務開拓客戶資料"
'      Exit Function
'   End If
'   '2021/8/31 END
'
''依序比對R,Y,X:
'   '比對中文名稱
'   strSql = "SELECT pcu01||pcu02" & _
'             " FROM PotCustomer" & _
'            " WHERE rtrim(pcu08)=rtrim('" & ChgSQL(strPCUName) & "') and substr(pcu09,1,3)='" & strNA01 & "' and pcu02='0'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'       GoTo gotoExit
'   End If
'   rsTmp.Close
'
'   If Option1(0).Value = True Then '國內
'      strSql = "SELECT poc01||poc02" & _
'                " FROM PotCustomer1" & _
'               " WHERE rtrim(poc03)=rtrim('" & ChgSQL(strPCUName) & "') and substr(poc04,1,3)='" & strNA01 & "' and poc02='0'"
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount > 0 Then
'         rsTmp.MoveFirst
'         Do While Not rsTmp.EOF
'            If IsNull(rsTmp.Fields(0)) = False Then
'               strCustID = strCustID & "," & rsTmp.Fields(0)
'            End If
'            rsTmp.MoveNext
'         Loop
'         GoTo gotoExit
'      End If
'      rsTmp.Close
'   End If
'
'   strSql = "SELECT fa01||fa02" & _
'             " FROM Fagent" & _
'            " WHERE rtrim(fa04)=rtrim('" & ChgSQL(strPCUName) & "') and substr(fa10,1,3)='" & strNA01 & "' and fa02='0'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'      GoTo gotoExit
'   End If
'   rsTmp.Close
'
'   strSql = "SELECT cu01||cu02" & _
'             " FROM Customer" & _
'            " WHERE rtrim(cu04)=rtrim('" & ChgSQL(strPCUName) & "') and substr(cu10,1,3)='" & strNA01 & "' and cu02='0'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'      GoTo gotoExit
'   End If
'   rsTmp.Close
'
'   '比對英文名稱
'   strSql = "SELECT pcu01||pcu02" & _
'             " FROM PotCustomer" & _
'            " WHERE rtrim(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06))=rtrim('" & ChgSQL(UCase(Trim(strPCUName))) & "')" & _
'            " and substr(pcu09,1,3)='" & strNA01 & "' and pcu02='0'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'      GoTo gotoExit
'   End If
'   rsTmp.Close
'
'   If Option1(0).Value = True Then '國內
'      strSql = "SELECT poc01||poc02" & _
'                " FROM PotCustomer1" & _
'               " WHERE rtrim(upper(poc23||' '||poc24||' '||poc25||' '||poc26))=rtrim('" & ChgSQL(UCase(Trim(strPCUName))) & "')" & _
'               " and substr(poc04,1,3)='" & strNA01 & "' and poc02='0'"
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount > 0 Then
'         rsTmp.MoveFirst
'         Do While Not rsTmp.EOF
'            If IsNull(rsTmp.Fields(0)) = False Then
'               strCustID = strCustID & "," & rsTmp.Fields(0)
'            End If
'            rsTmp.MoveNext
'         Loop
'         GoTo gotoExit
'      End If
'      rsTmp.Close
'   End If
'
'   strSql = "SELECT fa01||fa02" & _
'             " FROM Fagent" & _
'            " WHERE rtrim(upper(fa05||' '||fa63||' '||fa64||' '||fa65))=rtrim('" & ChgSQL(UCase(Trim(strPCUName))) & "')" & _
'            " and substr(fa10,1,3)='" & strNA01 & "' and fa02='0'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'      GoTo gotoExit
'   End If
'   rsTmp.Close
'
'   strSql = "SELECT cu01||cu02" & _
'             " FROM Customer" & _
'            " WHERE rtrim(upper(cu05||' '||cu88||' '||cu89||' '||cu90))=rtrim('" & ChgSQL(UCase(Trim(strPCUName))) & "')" & _
'            " and substr(cu10,1,3)='" & strNA01 & "' and cu02='0'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If IsNull(rsTmp.Fields(0)) = False Then
'            strCustID = strCustID & "," & rsTmp.Fields(0)
'         End If
'         rsTmp.MoveNext
'      Loop
'      GoTo gotoExit
'   End If
'   rsTmp.Close
'
'gotoExit:
   If strCustID <> "" Then
      ChkCustNameAndPotCust_21 = Mid(strCustID, 2, Len(strCustID))
   Else
      ChkCustNameAndPotCust_21 = ""
   End If
   
   Set rsTmp = Nothing
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

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

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "" '"*.xlsx"
      .Filter = "files (*.xls)|*.xls|files (*.xlsx)|*.xlsx|"
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
   
   Call PUB_SetComboPCU11(cboPCU11, "") '設定國外潛在客戶類別選項
   
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
   
   Set frm140421 = Nothing
End Sub

Private Sub txtPath1_GotFocus()
   InverseTextBox txtPath1
End Sub
