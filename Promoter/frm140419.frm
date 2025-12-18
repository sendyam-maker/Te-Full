VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140419 
   BorderStyle     =   1  '單線固定
   Caption         =   "潛在案量客戶名稱比對"
   ClientHeight    =   4692
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4692
   ScaleWidth      =   6180
   Begin VB.CheckBox Chk2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "客戶名稱資訊　不得宣傳比對"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   4440
      TabIndex        =   18
      Top             =   510
      Width           =   1700
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Excel 英文檔"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   4920
      TabIndex        =   16
      Top             =   1140
      Width           =   1400
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Excel中文檔(中文規則)"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   2820
      TabIndex        =   15
      Top             =   1140
      Width           =   2300
   End
   Begin VB.TextBox txtPutField 
      Height          =   264
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1380
      Width           =   500
   End
   Begin VB.TextBox txtCountry 
      Height          =   264
      Left            =   630
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1380
      Width           =   615
   End
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
      Height          =   2160
      ItemData        =   "frm140419.frx":0000
      Left            =   30
      List            =   "frm140419.frx":0002
      TabIndex        =   9
      Top             =   2400
      Width           =   6100
   End
   Begin VB.TextBox txtFileName 
      Height          =   264
      Left            =   30
      TabIndex        =   3
      Top             =   2112
      Width           =   5650
   End
   Begin VB.CommandButton CmdOpenFile 
      Caption         =   "<="
      Height          =   250
      Left            =   5805
      TabIndex        =   2
      Top             =   2112
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
      Caption         =   "比對"
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
   Begin MSForms.TextBox txtOrgName 
      Height          =   264
      Left            =   0
      TabIndex        =   12
      Top             =   4536
      Visible         =   0   'False
      Width           =   612
      VariousPropertyBits=   671107099
      Size            =   "1080;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "不得宣傳比對：因置入資料有7欄,故欄位最大只能輸至S欄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   6
      Left            =   1280
      TabIndex        =   19
      Top             =   1896
      Width           =   5508
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "以Excel 中、英文檔決定比對方式："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   30
      TabIndex        =   17
      Top             =   1170
      Width           =   3000
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "因置入資料有6欄,故欄位最大只能輸至T欄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   3
      Left            =   2610
      TabIndex        =   14
      Top             =   1656
      Width           =   4596
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "比對結果欄位置入開始位置："
      Height          =   210
      Index           =   2
      Left            =   3240
      TabIndex        =   13
      Top             =   1410
      Width           =   2505
   End
   Begin VB.Label lblCountry 
      Caption         =   "lblCountry"
      Height          =   210
      Left            =   1320
      TabIndex        =   11
      Top             =   1410
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "國籍："
      Height          =   210
      Index           =   1
      Left            =   30
      TabIndex        =   10
      Top             =   1410
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "1."
      ForeColor       =   &H00FF0000&
      Height          =   885
      Left            =   30
      TabIndex        =   8
      Top             =   210
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注意事項："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   960
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "資料檔案："
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   1680
      Width           =   1320
   End
End
Attribute VB_Name = "frm140419"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2023/03/27 修正txtOrgName未改成Form2.0 元件,導致查上海眾啟(口在下)生物科技有限公司 查不出
'Memo By Lydia 2022/02/16 Form2.0已檢查 (無需修改的物件)
'Create By Amy 2018/10/03
Option Explicit
Dim strTmp As String
Dim i As Integer
Dim strF(), strF2(), intWidth2()
Dim intRow_O As Integer, intRow As Integer '原始筆數/Excel列數
Dim intRun As Integer 'Add by Amy 2021/03/23
Dim strExtension As String 'Add by Amy 2022/07/06 副檔名
'Add by Amy 2022/08/17
Dim stSQL1 As String, stSQL2 As String, stSQL3 As String, stSQL4 As String, stSQL5 As String  'Add by Amy 2022/08/17 從GetQuerySql搬出來
Dim strTxtSpec As String, strTxtSpec_E As String, arrTmp '特取字
Dim bolNoAdvertise As Boolean 'Add by Amy 2023/07/04 不得宣傳
Dim bolShowPTArea As Boolean 'Add by Amy 2024/05/31 代理人/更代 顯示P/T 台灣/非台灣統計

Private Sub Chk1_Click(Index As Integer)
    If Chk1(Index).Value = vbChecked Then
        '擇一勾選
        If Index = 0 Then
            Chk1(1).Value = 0
        Else
            Chk1(0).Value = 0
        End If
    End If
End Sub

Private Sub CmdChk_Click()
    'Modify by Amy 2023/07/04 原程式搬至FormCheck,+「客戶名稱資訊不得宣傳比對」勾選
    If FormCheck = False Then
      Exit Sub
    End If
    bolNoAdvertise = False
    '只有業務及電腦中心 可使用 「客戶名稱資訊不得宣傳比對」
    If Chk2.Visible = True Then
      If Chk2.Value = vbChecked Then
        bolNoAdvertise = True
        Chk2.Enabled = False
      End If
    End If
    'end 2023/07/04
    
    Screen.MousePointer = vbHourglass
    CmdChk.Enabled = False
    CmdExit.Enabled = False
    Call DelR140419
    If RunExcelChk = True Then MsgBox "檢查已完成！"
    'Add by Amy 2023/07/04
    If Chk2.Visible = True Then
      Chk2.Enabled = True
    End If
    bolNoAdvertise = False
    'end 2023/07/04
    Chk1(0).Value = 0: Chk1(1).Value = 0 'Add by Amy 2022/08/23
    Screen.MousePointer = vbDefault
    CmdChk.Enabled = True
    CmdExit.Enabled = True
End Sub

'Add by Amy 2023/07/04 從CmdChk_Click搬過來共用
Private Function FormCheck() As Boolean
   Dim bCancel As Boolean
   
   FormCheck = False
   'Add by Amy 2021/04/15
    If txtPutField = MsgText(601) Then
        MsgBox Mid(Label1(2), 1, Len(Label1(2)) - 1) & "不可空白！"
        txtPutField.SetFocus
        Exit Function
    End If
    Call txtPutField_Validate(bCancel)
    If bCancel = True Then
      Exit Function
    End If
    'end 2021/04/15
    If txtFileName = MsgText(601) Then
        MsgBox "檔案不可空白！"
        txtFileName.SetFocus
        Exit Function
    End If
    'Add by Amy 2022/08/23
    If Chk1(0).Value = 0 And Chk1(1).Value = 0 Then
        MsgBox "請勾選Excel 中文檔或Excel 英文檔！"
        Exit Function
    End If
    'Mark by Amy 2021/03/23 國籍可空白
'    If txtCountry = MsgText(601) Then
'        MsgBox "國籍不可空白！"
'        txtCountry.SetFocus
'        Exit Function
'    End If
   FormCheck = True
End Function

Private Function RunExcelChk_Old() As Boolean
'    Dim RsQ As New ADODB.Recordset, RsQ2 As New ADODB.Recordset
'    Dim xlsAp As New Excel.Application
'    Dim wksrpt As New Worksheet
'    Dim strQ As String, strName As String
'    Dim intQ As Integer, intQ2 As Integer
'    Dim bolFirst As Boolean
'    'Add by Amy 2018/10/17
'    Dim bolFName As Boolean
'
'On Error GoTo ErrHnd
'
'    ReDim strF(2), stF2(3)
'    ReDim intWidth2(3)
'    strF = Array("名稱", "案量", " ", " ", " ", " ", " ") 'Modify by Amy 2018/10/17 留到G欄(因下載的原始資料欄位不一致)-陳增廣
'    strF2 = Array("編號", "名稱", "代表號")
'    intWidth2 = Array(11, 33, 11)
'
'    intRow_O = 0: intRow = 2
'    RunExcelChk = False
'    List1.Clear
'    xlsAp.Visible = False
'    List1.AddItem "檢查開始：", 0
'    xlsAp.Workbooks.Open txtFileName
'    Set wksrpt = xlsAp.Worksheets(1)
'    strName = ChgSQL(RTrim(UCase(xlsAp.Range("A" & intRow))))
'    Do While strName <> MsgText(601)
'        bolFirst = True
'        bolFName = False 'Add by Amy 2018/10/17
'        intRow_O = intRow_O + 1
'        List1.AddItem "　" & strName, 0
'        '查詢整個名稱
'        'Modify by Amy 2020/10/29 改function
'        strQ = "Select fa01||fa02 as FNo,fa05||Decode(fa63,null,null,' '||fa63)||Decode(fa64,null,null,' '||fa64)||Decode(fa65,null,null,' '||fa65) as FName,fa16 as MailAddr From FAgent Where SubStr(fa10,1,3)='" & txtCountry & "' And InStr(Upper(fa05||' '||fa63||' '||fa64||' '||fa65),'" & strName & "')=1 And Length(RTrim(Upper(fa05||' '||fa63||' '||fa64||' '||fa65)))=" & Len(strName) & _
'        " Union Select pcu01||pcu02,pcu03||Decode(pcu04,null,null,' '||pcu04)||Decode(pcu05,null,null,' '||pcu05)||Decode(pcu06,null,null,' '||pcu06) as FName,pcu18 as MailAddr From PotCustomer Where SubStr(pcu09,1,3)='" & txtCountry & "' And InStr(Upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & strName & "')=1 And Length(RTrim(Upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06)))=" & Len(strName) & _
'        " Union Select poc01||poc02,poc23||Decode(poc24,null,null,' '||poc24)||Decode(poc25,null,null,' '||poc25)||Decode(poc26,null,null,' '||poc26) as FName,poc09 as MailAddr From PotCustomer1 Where SubStr(poc04,1,3)='" & txtCountry & "' And InStr(Upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & strName & "')=1 And Length(Rtrim(Upper(poc23||' '||poc24||' '||poc25||' '||poc26)))=" & Len(strName)
'        intQ = 1
'        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'        If intQ = 0 Then
'            '查詢字首
'            'Modify by Amy 2018/10/17 欄位加空白再查 ex:Madderns Patent & Trade Mark Attorneys
'            bolFName = True
'            strName = Mid(strName, 1, IIf(InStr(strName, " ") = 0, Len(strName) & " ", InStr(strName, " ")))
'            strQ = "Select fa01||fa02 as FNo,fa05||Decode(fa63,null,null,' '||fa63)||Decode(fa64,null,null,' '||fa64)||Decode(fa65,null,null,' '||fa65) as FName,fa16 as MailAddr From FAgent Where SubStr(fa10,1,3)='" & txtCountry & "' And Upper(fa05)||' ' Like '" & strName & "%' " & _
'            "Union Select pcu01||pcu02,pcu03||Decode(pcu04,null,null,' '||pcu04)||Decode(pcu05,null,null,' '||pcu05)||Decode(pcu06,null,null,' '||pcu06) as FName,pcu18 as MailAddr From PotCustomer Where SubStr(pcu09,1,3)='" & txtCountry & "' And Upper(pcu03)||' ' Like '" & strName & "%' " & _
'            "Union Select poc01||poc02,poc23||Decode(poc24,null,null,' '||poc24)||Decode(poc25,null,null,' '||poc25)||Decode(poc26,null,null,' '||poc26) as FName,poc09 as MailAddr From PotCustomer1 Where SubStr(poc04,1,3)='" & txtCountry & "' And Upper(poc23)||' ' Like '" & strName & "%' "
'            'end 2018/10/17
'            intQ = 1
'            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'        End If
'        If intQ = 1 Then
'            Do While RsQ.EOF = False
'                For i = 0 To UBound(strF2)
'                    If bolFirst = False And i = 0 Then
'                        wksrpt.Range("A" & intRow).Select
'                        xlsAp.Selection.EntireRow.Insert
'                    End If
'                    'Modify by Amy 2019/03/15 原:Right("" & RsQ.Fields("FNo"), 1) = "1"
'                    If Right("" & RsQ.Fields("FNo"), 1) <> "0" And i = 0 Then
'                        wksrpt.Range(Chr(UBound(strF) + 66 + i) & intRow).Value = "" & RsQ.Fields(i) & "＊"
'                    Else
'                        wksrpt.Range(Chr(UBound(strF) + 66 + i) & intRow).Value = "" & RsQ.Fields(i)
'                    End If
'                Next i
'                intRow = intRow + 1
'                If bolFirst = True Then bolFirst = False
'                'Add by Amy 2018/10/17 若整個名稱查詢出來資料為更名,將目前最新名稱顯示
'                'Modify by Amy 2019/03/15 原:Right("" & RsQ.Fields("FNo"), 1) = "1"
'                If bolFName = False And Right("" & RsQ.Fields("FNo"), 1) <> "0" Then
'                    strQ = "" & RsQ.Fields("FNo")
'                    strQ = "Select fa01||fa02 as FNo,fa05||Decode(fa63,null,null,' '||fa63)||Decode(fa64,null,null,' '||fa64)||Decode(fa65,null,null,' '||fa65) as FName,fa16 as MailAddr From FAgent Where fa01='" & Mid(strQ, 1, 8) & "' And fa02='0' " & _
'                            " Union Select pcu01||pcu02,pcu03||Decode(pcu04,null,null,' '||pcu04)||Decode(pcu05,null,null,' '||pcu05)||Decode(pcu06,null,null,' '||pcu06) as FName,pcu18 as MailAddr From PotCustomer Where pcu01='" & Mid(strQ, 1, 8) & "' And pcu02='0' " & _
'                            " Union Select poc01||poc02,poc23||Decode(poc24,null,null,' '||poc24)||Decode(poc25,null,null,' '||poc25)||Decode(poc26,null,null,' '||poc26) as FName,poc09 as MailAddr From PotCustomer1 Where poc01='" & Mid(strQ, 1, 8) & "' And poc02='0' "
'                    intQ2 = 1
'                    Set RsQ2 = ClsLawReadRstMsg(intQ2, strQ)
'                    Do While RsQ2.EOF = False
'                        For i = 0 To UBound(strF2)
'                            If i = 0 Then
'                                wksrpt.Range("A" & intRow).Select
'                                xlsAp.Selection.EntireRow.Insert
'                            End If
'                            wksrpt.Range(Chr(UBound(strF) + 66 + i) & intRow).Value = "" & RsQ2.Fields(i)
'                        Next i
'                        intRow = intRow + 1
'                        RsQ2.MoveNext
'                    Loop
'                End If
'                'end 2018/10/17
'                RsQ.MoveNext
'            Loop
'        Else
'            intRow = intRow + 1
'        End If
'        strName = ChgSQL(RTrim(UCase(xlsAp.Range("A" & intRow))))
'    Loop
'    '設定欄位名稱/欄寬
'    For i = 0 To UBound(strF2)
'        wksrpt.Range(Chr(UBound(strF) + 66 + i) & "1").Value = strF2(i)
'        wksrpt.Range(Chr(UBound(strF) + 66 + i) & "1").ColumnWidth = intWidth2(i)
'    Next i
'    wksrpt.Range(Chr(UBound(strF) + 66) & "1:" & Chr(UBound(strF) + 66 + UBound(strF2)) & "1").Interior.ColorIndex = 44
'    List1.AddItem "檢查完成 " & intRow_O & " 筆", 0
'
'    xlsAp.Workbooks(1).Save
'    xlsAp.Workbooks.Close
'    xlsAp.Quit
'    Set xlsAp = Nothing
'
'    RunExcelChk = True
'    Exit Function
'
'ErrHnd:
'    List1.AddItem "匯入失敗！請通知電腦中心(" & Err.Description & ")", 0
'    xlsAp.Workbooks(1).Save
'    xlsAp.Workbooks.Close
'    xlsAp.Quit
'    Set xlsAp = Nothing
End Function

Private Sub cmdExit_Click()
     Unload Me
End Sub

Private Sub CmdOpenFile_Click()
    Dim stFileName As String
    Dim sFile
On Error GoTo ErrHnd
  
    stFileName = ""
    strExtension = "" 'Add by Amy 2022/07/06 副檔名
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
               For i = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), i, 1) = "\" Then
                     '記錄選過的路徑
                     PUB_SaveLastDate Me.Name, "Dir", Mid(Trim(.FileName), 1, i - 1)
                     Exit For
                  End If
               Next i
            End If
            'Add by Amy 2022/07/06 記錄副檔名,避免匯入之 xlsx 檔案另存成 xls(格式可能與xls無法相容,出現相容性檢查訊息)會彈錯誤
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
                               "2.名稱欄中間不可有空白列。" & vbCrLf & _
                               "3.不輸入國籍條件時，所需比對時間比較久！" & vbCrLf & _
                               "4.Excel 資料中、英文請分成兩個檔比對" & vbCrLf & _
                               "   避免英文資料比對「中文」規則時，非常耗時且不準確"
    lblCountry.Caption = ""
    
    'Modify by Amy 2024/04/25 從RunExcelChk 搬過來,原只有杜協理及業拓用,之後增加管理部其他人員使用,Label1(3)及(6)說明改動態
    strF = Array("名稱", "案量", " ", " ", " ", " ", " ") '留到G欄(因下載的原始資料欄位不一致)-陳增廣
'*** Memo by Amy 此處有加欄位,也要確認 RunExcelChk 是否要加 ***
    'Modify by Amy 2024/05/31 業拓產生的檔案 +P/T 台灣/非台灣統計
    If Pub_StrUserSt03 = "F41" Or Pub_StrUserSt03 = "M51" Then
      bolShowPTArea = True
      strF2 = Array("編號", "名稱", "狀態", "代表號", "智權人員", "業務區", "案件統計", "P台灣", "P非台灣", "T台灣", "T非台灣", "更代案件統計", "P台灣-更", "P非台灣-更", "T台灣-更", "T非台灣-更")
      intWidth2 = Array(14, 33, 10, 20, 11, 11, 30, 10, 10, 10, 10, 30, 10, 10, 10, 10)
    Else
      'Modify by Amy 2024/05/10 +更代案件統計
      strF2 = Array("編號", "名稱", "狀態", "代表號", "智權人員", "業務區", "案件統計", "更代案件統計")
      intWidth2 = Array(14, 33, 10, 20, 11, 11, 30, 30)
    End If
    'Add by Amy 2021/04/15
    'If strUserNum <> "74018" Then
        'txtPutField = "G"
    If Pub_StrUserSt03 <> "F41" Then
        txtPutField = Chr(65 + UBound(strF))
    End If
    Label1(3).Caption = "因置入資料有" & UBound(strF2) + 1 & "欄,故欄位最大只能輸至" & Chr(Asc("Z") - (UBound(strF2) + 1)) & "欄"
    Label1(6).Caption = "不得宣傳比對：因置入資料有" & UBound(strF2) + 2 & "欄,故欄位最大只能輸至" & Chr(Asc("Z") - (UBound(strF2) + 2)) & "欄"
    'end 2024/04/25
    
    'Add by Amy 2023/07/04 +客戶名稱資訊不得宣傳比對 勾選,業拓及電腦中心 用
    Chk2.Visible = False: Label1(6).Visible = False
    If Pub_StrUserSt03 = "F41" Or Pub_StrUserSt03 = "M51" Then
      Chk2.Visible = True: Label1(6).Visible = True
    End If
    
    'Add by Amy 2022/08/17 搬出來
    'Modify by Amy 2023/06/08 傳入要查之字串先將數字、英文變全型,故將原(股)->半型 括號 改全型
    '                                                     +（股份有限）公司 ex:X5467500 /（股）公司 ex:X6319600
    'Modify by Amy 2023/07/05 +股份公司/公司
    strTxtSpec = "台灣區;臺灣區;台灣;臺灣;中華民國;社團法人;財團法人;股份有限公司;（股）有限公司;有限公司;（股份有限）公司;（股）公司;股份公司;公司;" & _
                            "個人工作室;工作室;社會福利基金會;研究發展基金會;基金會"
    'Memo by Amy 2022/08/17 不加縮寫,怕取代到不該取代的字
    strTxtSpec_E = "INCORPORATED;CORPORATION;LIMITED;COMPANY"
    arrTmp = Split(strTxtSpec_E, ";")
    stSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
    stSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
    stSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
    stSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
    stSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm140419 = Nothing
End Sub

Private Sub txtCountry_GotFocus()
    CloseIme
    TextInverse txtCountry
End Sub

Private Sub txtCountry_Validate(Cancel As Boolean)
    Dim strTmp As String
    
    lblCountry.Caption = ""
    If txtCountry = MsgText(601) Then Exit Sub
    
    If ClsPDGetNation(txtCountry, strTmp) = True Then
        lblCountry.Caption = strTmp
    End If
End Sub

'Add by Amy 2020/10/29
'原只比對代理人、國內外潛在客戶英文欄位,增加客戶檔並檢查中文日欄
'Modify by Amy 2021/03/23 抓到更名前料需顯示最新名稱改暫存檔加,避免有沒抓到的資料,+bolReplace
'intChoose:1-整個名稱相等比對/2-字首/3-名稱模糊比對
'bolReplace:是否特取
Private Function GetQuerySql(ByVal intChoose As Integer, stSearchTxt As String, Optional ByVal bolReplace As Boolean = False) As String
    Dim i As Integer, j As Integer
    Dim stSearchC(4) As String, stSearchE(4) As String, stSearchJ(4) As String  'Modify by Amy 2021/08/31 原:3
    Dim stFieldC(4) As String, stFieldE(4) As String, stFieldJ(4) As String 'Modify by Amy 2021/08/31 原:3
    Dim strQ As String, stField As String, stSearch As String, stTB As String, stWhere As String
    Dim stTP As String, stTp2 As String
    Dim strCheckWay As String, stSQL1 As String, stSQL2 As String, stSQL3 As String, stSQL4 As String, stSQL5 As String 'Add by Amy 2021/08/16
    Dim stVTB As String 'Add by Amy 2021/08/31
    
    'Modify by Amy 2021/03/23 因國籍空白,會Run 台灣及非台灣,資料可能重覆抓,故寫暫存檔
    '中文
    For i = LBound(stSearchC) To UBound(stSearchC)
        Select Case i
            Case 0
                stFieldC(i) = "cu04"
            Case 1
                stFieldC(i) = "fa04"
            Case 2
                stFieldC(i) = "pcu08"
            Case 3
                stFieldC(i) = "poc03"
            Case 4 'Add by Amy 2021/08/31 +不得代理
                stFieldC(i) = "nt02"
        End Select
        If bolReplace = True Then
            stTP = "Decode(SubStr(" & stFieldC(i) & ",1,2),'台灣',SubStr(" & stFieldC(i) & ",3,length(" & stFieldC(i) & ")),Decode(SubStr(" & stFieldC(i) & ",1,2),'臺灣',SubStr(" & stFieldC(i) & ",3,length(" & stFieldC(i) & "))," & stFieldC(i) & "))"
            stTP = "Replace(Replace(Replace(Replace(Replace(Upper(" & stTP & "),'股份有限公司',''),'(股)有限公司',''),'有限公司',''),'個人工作室',''),'工作室','')"
            stTP = "Upper(" & stTP & ")"
        Else
            stTP = "Upper(" & stFieldC(i) & ")"
        End If
         
        '整個名稱/模糊比對
        If intChoose = 1 Or intChoose = 3 Then
            stSearchC(i) = "=1 "
            If intChoose = 3 Then stSearchC(i) = ">0 "
            stSearchC(i) = "And InStr(" & stTP & ",'" & ChgSQL(stSearchTxt) & "')" & stSearchC(i)
        '字首
        ElseIf intChoose = 2 Then
            stSearchC(i) = "And " & stTP & " Like '" & ChgSQL(stSearchTxt) & "%' "
        End If
    Next i
    '英文
    For i = LBound(stSearchE) To UBound(stSearchE)
        Select Case i
            Case 0
                    stFieldE(i) = "cu05||Decode(cu88,null,null,' '||cu88)||Decode(cu89,null,null,' '||cu89)||Decode(cu90,null,null,' '||cu90) "
                    stTP = "cu05||' '||cu88||' '||cu89||' '||cu90"
                    If intChoose = 2 Then stTP = "cu05"
                Case 1
                    stFieldE(i) = "fa05||Decode(fa63,null,null,' '||fa63)||Decode(fa64,null,null,' '||fa64)||Decode(fa65,null,null,' '||fa65)"
                    stTP = "fa05||' '||fa63||' '||fa64||' '||fa65"
                    If intChoose = 2 Then stTP = "fa05"
                Case 2
                    stFieldE(i) = "pcu03||Decode(pcu04,null,null,' '||pcu04)||Decode(pcu05,null,null,' '||pcu05)||Decode(pcu06,null,null,' '||pcu06)"
                    stTP = "pcu03||' '||pcu04||' '||pcu05||' '||pcu06"
                    If intChoose = 2 Then stTP = "pcu03"
                Case 3
                    stFieldE(i) = "poc23||Decode(poc24,null,null,' '||poc24)||Decode(poc25,null,null,' '||poc25)||Decode(poc26,null,null,' '||poc26)"
                    stTP = "poc23||' '||poc24||' '||poc25||' '||poc26"
                    If intChoose = 2 Then stTP = "poc23"
                Case 4 'Add by Amy 2021/08/31 +不得代理
                    stFieldE(i) = "nt03||' '||nt04||' '||nt05||' '||nt06"
                    stTP = "nt03||' '||nt04||' '||nt05||' '||nt06"
                    If intChoose = 2 Then stTP = "nt03"
        End Select
        '整個名稱
        If intChoose = 1 Or intChoose = 3 Then
            stSearchE(i) = "=1 "
            If intChoose = 3 Then stSearchE(i) = ">0 "
            stSearchE(i) = "And InStr(Upper(" & stTP & "),'" & ChgSQL(stSearchTxt) & "')" & stSearchE(i)
        '英文字首需加空白查 ex:Madderns Patent & Trade Mark Attorneys,Madderns(Y20697)不出現
        ElseIf intChoose = 2 Then
            stSearchE(i) = " And Upper(" & stTP & ")||' ' Like '" & ChgSQL(stSearchTxt) & "%' "
        End If
    Next i
    '日文
    For i = LBound(stSearchJ) To UBound(stSearchJ)
        Select Case i
            Case 0
                stFieldJ(i) = "cu06"
            Case 1
                stFieldJ(i) = "fa06"
            Case 2
                stFieldJ(i) = "pcu07"
            Case 3
                stFieldJ(i) = "poc27"
            Case 4 'Add by Amy 2021/08/31 +不得代理
                stFieldJ(i) = "nt07"
        End Select
        stTP = "Upper(" & stFieldJ(i) & ")"
        '整個名稱
        If intChoose = 1 Or intChoose = 3 Then
            stSearchJ(i) = "=1 "
            If intChoose = 3 Then stSearchJ(i) = ">0 "
            stSearchJ(i) = "And InStr(" & stTP & ",'" & ChgSQL(stSearchTxt) & "')" & stSearchJ(i)
        '字首
        ElseIf intChoose = 2 Then
            stSearchJ(i) = "And " & stTP & " Like '" & ChgSQL(stSearchTxt) & "%' "
        End If
    Next i
    'end 2021/03/23
    
    For i = LBound(stSearchC) To UBound(stSearchC)
        Select Case i
            'Modify by Amy 2021/03/23 增加顯示智權人員及業務區
            Case 0
                stField = "cu01||cu02 as FNo,cu20 as MailAddr,cu13 as SalesNo,cu12 as SalesAreaNo"
                stTB = "Customer"
                'Modify by Amy 2021/03/21 +intRun
                If txtCountry = 台灣國家代號 Or intRun = 2 Then
                    stWhere = " Where SubStr(cu10,1,3)>='000' And SubStr(cu10,1,3)<'010' "
                'Modify by Amy 2021/03/21 +If txtCountry <> MsgText(601)
                ElseIf txtCountry <> MsgText(601) Then
                    stWhere = " Where SubStr(cu10,1,3)='" & txtCountry & "' "
                End If
            Case 1
                stField = "fa01||fa02 as FNo,fa16 as MailAddr,'' as SalesNo,'' as SalesAreaNo"
                stTB = "FAgent"
                'Modify by Amy 2021/03/21 +intRun
                If txtCountry = 台灣國家代號 Or intRun = 2 Then
                    stWhere = " Where SubStr(fa10,1,3)>='000' And SubStr(fa10,1,3)<'010' "
                'Modify by Amy 2021/03/21 +If txtCountry <> MsgText(601)
                ElseIf txtCountry <> MsgText(601) Then
                    stWhere = " Where SubStr(fa10,1,3)='" & txtCountry & "' "
                End If
            Case 2
                stField = "pcu01||pcu02 as FNo,pcu18 as MailAddr,pcu38 as SalesNo,'' as SalesAreaNo"
                stTB = "PotCustomer"
                'Modify by Amy 2021/03/21 +intRun
                If txtCountry = 台灣國家代號 Or intRun = 2 Then
                    stWhere = " Where SubStr(pcu09,1,3)>='000' And SubStr(pcu09,1,3)<'010' "
                'Modify by Amy 2021/03/21 +If txtCountry <> MsgText(601)
                ElseIf txtCountry <> MsgText(601) Then
                    stWhere = " Where SubStr(pcu09,1,3)='" & txtCountry & "' "
                End If
            Case 3
                stField = "poc01||poc02 as FNo,poc09 as MailAddr,poc13 as SalesNo,'' as SalesAreaNo"
                stTB = "PotCustomer1"
                'Modify by Amy 2021/03/21 +intRun
                If txtCountry = 台灣國家代號 Or intRun = 2 Then
                    stWhere = " Where SubStr(poc04,1,3)>='000' And SubStr(poc04,1,3)<'010' "
                'Modify by Amy 2021/03/21 +If txtCountry <> MsgText(601)
                ElseIf txtCountry <> MsgText(601) Then
                    stWhere = " Where SubStr(poc04,1,3)='" & txtCountry & "' "
                End If
            'end 2021/03/23
            Case 4 'Add by Amy 2021/08/31 +不得代理
                stField = "nt01 as FNo,'' as MailAddr,'' as SalesNo,'' as SalesAreaNo"
                stTB = "NotAgent"
                If txtCountry = 台灣國家代號 Or intRun = 2 Then
                    stWhere = " Where SubStr(nt08,1,3)>='000' And SubStr(nt08,1,3)<'010' "
                ElseIf txtCountry <> MsgText(601) Then
                    stWhere = " Where SubStr(nt08,1,3)='" & txtCountry & "' "
                End If
        End Select
        For j = 1 To 3
            Select Case j
                Case 1
                    stTP = "," & stFieldC(i)
                    stTp2 = stSearchC(i)
                Case 2
                    stTP = "," & stFieldE(i)
                    stTp2 = stSearchE(i)
                Case 3
                    stTP = "," & stFieldJ(i)
                    stTp2 = stSearchJ(i)
            End Select
            strQ = strQ & " Union Select " & stField & stTP & " as FName," & j & " as sField From " & stTB & IIf(stWhere = "", " Where " & Mid(stTp2, 5), stWhere & stTp2)
        Next j
    Next i
    'Add by Amy 2021/08/31
    If txtCountry = 台灣國家代號 Or intRun = 2 Then
        stWhere = " And SubStr(ecd10,1,3)>='000' And SubStr(ecd10,1,3)<'010' "
    ElseIf txtCountry <> MsgText(601) Then
        stWhere = " And SubStr(ecd10,1,3)='" & txtCountry & "' "
    End If
    '*** 法務開拓 ***
    '整個名稱/模糊比對
    If intChoose = 1 Or intChoose = 3 Then
        stSearchC(0) = "=1 "
        If intChoose = 3 Then stSearchC(0) = ">0 "
        stSearchC(0) = " And (InStr(Upper(ecd03) ,'" & ChgSQL(stSearchTxt) & "')" & stSearchC(0) & _
                                    " Or InStr(Upper(ecd04) ,'" & ChgSQL(stSearchTxt) & "')" & stSearchC(0) & ")"
    '字首
    ElseIf intChoose = 2 Then
        stSearchC(0) = " And (Upper(ecd03) Like '" & ChgSQL(stSearchTxt) & "%' Or Upper(ecd04) Like '" & ChgSQL(stSearchTxt) & "%' )"
    End If
    strQ = strQ & " Union Select ecd02||'-'||LPAD(ecd01,6,'0') as FNo,ecd13 as MailAddr,'' as SalesNo,'' as SalesAreaNo,NVL(ecd03,'')||NVL(ecd04,'') as FName,9 as sField From ExPandCusDetail Where 1=1 " & stWhere & stSearchC(0)
    strQ = strQ & " Union Select ecd02||'-'||LPAD(ecd01,6,'0') as FNo,ecd13 as MailAddr,'' as SalesNo,'' as SalesAreaNo,NVL(ecd11,'')||NVL(ecd12,'') as FName,9 as sField From ExPandCusDetail Where 1=1 " & stWhere & Replace(Replace(stSearchC(0), "ecd03", "ecd11"), "ecd04", "ecd12")
    '*** End 法務開拓 ***
    
    '整個名稱/模糊比對
    If intChoose = 1 Or intChoose = 3 Then
        stSearchC(0) = "=1 "
        If intChoose = 3 Then stSearchC(0) = ">0 "
        stSearchC(0) = " And InStr(Upper(tbnp01) ,'" & ChgSQL(stSearchTxt) & "')" & stSearchC(0)
    '字首
    ElseIf intChoose = 2 Then
        stSearchC(0) = " And (Upper(tbnp01) Like '" & ChgSQL(stSearchTxt) & "%' )"
    End If
    '*** 國內開拓函特定公司不列印者 ***
    strQ = strQ & " Union Select '' as FNo,'' as MailAddr,'' as SalesNo,'' as SalesAreaNo,tbnp01 as FName,9 as sField From TMBulletinnp Where 1=1 " & stSearchC(0)
    
    '*** 聯絡人 ***
    stVTB = "Select * From PotCustCont Where 1=1 "
    For j = 1 To 3
        Select Case j
            Case 1 '中
                stField = "pcc05"
                stVTB = stVTB & Replace(stSearchC(0), "tbnp01", "pcc05")
            Case 2 '英
                stField = "pcc03"
                stVTB = stVTB & Replace(stSearchC(0), "tbnp01", "pcc03")
            Case 3 '日
                stField = "pcc04"
                stVTB = stVTB & Replace(stSearchC(0), "tbnp01", "pcc04")
        End Select
        For i = 1 To 4
            Select Case i
                Case 1
                    stTB = ",Customer,Staff"
                    stWhere = " And cu01(+)=pcc01 And cu02='0' And cu13=st01(+) "
                Case 2
                    stTB = ",PotCustomer,Staff"
                    stWhere = " And pcu01(+)=pcc01 And pcu02='0' And SubStr(LTrim(pcu38),1,5)=st01(+) "
                Case 3
                    stTB = ",PotCustomer1,Staff"
                    stWhere = " And poc01(+)=pcc01 And poc='0' And poc13=st01(+) "
                Case 4
                    stTB = ",Fagent"
                    stWhere = " And fa01(+)=pcc01 And fa02='0'"
            End Select
        Next i
        strQ = strQ & " Union Select pcc01||'0-'||pcc02 as FNo,pcc08 as MailAddr,'' as SalesNo,'' as SalesAreaNo," & stField & " as FName," & j & " as sField From (" & stVTB & ")" & stTB & " Where 1=1 " & stWhere
    Next j
    'end 2021/08/31
    'Add by Amy 2021/08/16 +對造(怕資料量太多以=比對)
    If intChoose = 1 Then
        stSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
        stSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
        stSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
        stSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
        stSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
        strCheckWay = "="
        Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5, ChgSQL(stSearchTxt), strCheckWay, True)
        strQ = strQ & " Union Select Distinct '對造' as FNo,'' as MailAddr,cp13,cp12,R021002||' '||R021001 as FName,9 as sField From R100102_1,CaseProgress " & _
                                            "Where ID='" & strUserNum & "@" & Me.Name & "' And R021004<3 And R021006=cp09(+) "
    End If
    'end 2021/08/16
    
    'Modify by Amy 2021/03/23 +智權人員及業務區
    GetQuerySql = Mid(strQ, 7)
End Function

Private Function ReplaceTWName(ByVal stName) As String
    'Mark by Amy 2022/08/17 改共用function,故不使用
'    ReplaceTWName = ""
'    If Left(stName, 2) = "台灣" Or Left(stName, 2) = "臺灣" Then stName = Mid(stName, 3)
'    stName = Replace(Replace(Replace(stName, "股份有限公司", ""), "(股)有限公司", ""), "有限公司", "")
'    stName = Replace(Replace(stName, "個人工作室", ""), "工作室", "")
'    ReplaceTWName = stName
End Function

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strF2)
       If UCase(strF2(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
'end 2020/10/29

'Mark by Amy 2021/03/25 改成存暫存檔
Private Function RunExcelChk_Old2() As Boolean
'    Dim RsQ As New ADODB.Recordset, RsQ2 As New ADODB.Recordset
'    Dim xlsAp As New Excel.Application
'    Dim wksrpt As New Worksheet
'    Dim strQ As String, strName As String
'    Dim intQ As Integer, intQ2 As Integer
'    Dim bolFirst As Boolean, bolFName As Boolean '第一次/字首查
'    Dim strTmp As String
'
'On Error GoTo ErrHnd
'
'    ReDim strF(2), stF2(3)
'    ReDim intWidth2(3)
'    strF = Array("名稱", "案量", " ", " ", " ", " ", " ") '留到G欄(因下載的原始資料欄位不一致)-陳增廣
'    strF2 = Array("編號", "名稱", "代表號")
'    intWidth2 = Array(11, 33, 11)
'
'    intRow_O = 0: intRow = 2
'    RunExcelChk = False
'    List1.Clear
'    xlsAp.Visible = False
'    List1.AddItem "檢查開始：", 0
'    xlsAp.Workbooks.Open txtFileName
'    Set wksrpt = xlsAp.Worksheets(1)
'    txtOrgName = ChgSQL(RTrim(UCase(xlsAp.Range("A" & intRow))))
'    strName = txtOrgName
'    'Excel A欄 有空白就離開
'    Do While strName <> MsgText(601)
'        bolFirst = True
'        bolFName = False
'        intRow_O = intRow_O + 1
'        List1.AddItem "　" & strName, 0
'        '*** 畫面國籍 台灣 (全名字首查->去除特取若名稱後有空白抓空白前的字模糊查->名稱+空白+名稱,抓空白後名稱查 有更名前資料,目前最新名稱資料也抓) ***
'        If txtCountry = 台灣國家代號 Then
'            '全名字首查(InStr=1)
'            strQ = GetQuerySql(1, strName)
'            intQ = 1
'            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'            If intQ = 0 Then
'                '全名模糊比對(InStr>0)
'                strName = ReplaceTWName(strName)
'                If InStr(strName, " ") > 0 Then
'                    strName = Mid(strName, 1, Val(InStr(strName, " ")) - 1)
'                ElseIf InStr(strName, "　") > 0 Then
'                    strName = Mid(strName, 1, Val(InStr(strName, "　")) - 1)
'                End If
'                strQ = GetQuerySql(3, strName)
'                intQ = 1
'                Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'                If intQ = 0 Then
'                     strQ = GetQuerySql(1, strName)
'                    intQ = 1
'                    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'                    If intQ = 0 And (InStr(txtOrgName, " ") > 0 Or InStr(txtOrgName, "　") > 0) Then
'                        '名稱+空白+名稱,抓空白後名稱查
'                        bolFName = True
'                        If InStr(txtOrgName, " ") > 0 Then
'                            strName = Mid(txtOrgName, Val(InStr(txtOrgName, " ")) + 1)
'                        ElseIf InStr(txtOrgName, "　") > 0 Then
'                            strName = Mid(txtOrgName, Val(InStr(txtOrgName, "　")) + 1)
'                        End If
'
'                        strQ = GetQuerySql(1, strName)
'                        intQ = 1
'                        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'                    End If
'                End If 'end 全名模糊比對(InStr>0)
'            End If
'
'        '*** End 畫面國籍 台灣 ***
'        Else
'        '*** 畫面國籍 非台灣 (全名查->字首+空白查 有更名前資料,目前最新名稱資料也抓) ***
'            '全名查(InStr=1)
'            strQ = GetQuerySql(1, strName)
'            intQ = 1
'            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'            If intQ = 0 Then
'                '查詢字首(like 字首%)
'                bolFName = True
'                strName = Mid(strName, 1, IIf(InStr(strName, " ") = 0, Len(strName) & " ", InStr(strName, " ")))
'                strQ = GetQuerySql(2, strName)
'                intQ = 1
'                Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'            End If
'        '*** End 畫面國籍 非台灣 ***
'        End If
'        '造字設顏色
'        If InStr(txtOrgName, "?") > 0 Then
'            xlsAp.Range("A" & intRow).Interior.ColorIndex = 41 '設置儲存格填充色(藍)
'            xlsAp.Range("A" & intRow).Interior.tintandshade = 0.5 '設深淺
'        End If
'        If intQ = 1 Then
'            Do While RsQ.EOF = False
'                For i = 0 To UBound(strF2)
'                    '不是第一筆需新增列,將資料插入
'                    If bolFirst = False And i = 0 Then
'                        wksrpt.Range("A" & intRow).Select
'                        xlsAp.Selection.EntireRow.Insert
'                    End If
'                    '欄名對應抓資料
'                    strTmp = ""
'                    Select Case i
'                        Case GetValue("編號")
'                            strTmp = "" & RsQ.Fields("FNo")
'                            '非最新名稱抓顯示＊
'                            If Right("" & RsQ.Fields("FNo"), 1) <> "0" Then
'                                strTmp = strTmp & "＊"
'                            End If
'                        Case GetValue("名稱")
'                            strTmp = "" & RsQ.Fields("FName")
'                        Case GetValue("代表號")
'                            strTmp = "" & RsQ.Fields("MailAddr")
'                    End Select
'                    wksrpt.Range(Chr(UBound(strF) + 66 + i) & intRow).Value = strTmp
'                Next i
'                intRow = intRow + 1
'                If bolFirst = True Then bolFirst = False
'                '非全名查時,若整個名稱查詢出來資料為更名,將目前最新名稱顯示
'                If bolFName = False And Right("" & RsQ.Fields("FNo"), 1) <> "0" Then
'                    strQ = GetQuerySql(3, Mid("" & RsQ.Fields("FNo"), 1, 8))
'                    intQ2 = 1
'                    Set RsQ2 = ClsLawReadRstMsg(intQ2, strQ)
'                    Do While RsQ2.EOF = False
'                        For i = 0 To UBound(strF2)
'                            If i = 0 Then
'                                wksrpt.Range("A" & intRow).Select
'                                xlsAp.Selection.EntireRow.Insert
'                            End If
'                            '以欄名對應抓資料
'                            strTmp = ""
'                            Select Case i
'                                Case GetValue("編號")
'                                    strTmp = "" & RsQ2.Fields("FNo")
'                                Case GetValue("名稱")
'                                    strTmp = "" & RsQ2.Fields("FName")
'                                Case GetValue("代表號")
'                                    strTmp = "" & RsQ2.Fields("MailAddr")
'                            End Select
'                            wksrpt.Range(Chr(UBound(strF) + 66 + i) & intRow).Value = strTmp
'                        Next i
'                        intRow = intRow + 1
'                        RsQ2.MoveNext
'                    Loop
'                End If
'                RsQ.MoveNext
'            Loop
'        Else
'            intRow = intRow + 1
'        End If
'        txtOrgName = ChgSQL(RTrim(UCase(xlsAp.Range("A" & intRow))))
'        strName = txtOrgName
'    Loop
'
'    '設定欄位名稱/欄寬
'    For i = 0 To UBound(strF2)
'        wksrpt.Range(Chr(UBound(strF) + 66 + i) & "1").Value = strF2(i)
'        wksrpt.Range(Chr(UBound(strF) + 66 + i) & "1").ColumnWidth = intWidth2(i)
'    Next i
'    wksrpt.Range(Chr(UBound(strF) + 66) & "1:" & Chr(UBound(strF) + 66 + UBound(strF2)) & "1").Interior.ColorIndex = 44
'    List1.AddItem "檢查完成 " & intRow_O & " 筆", 0
'
'    '另存
'    If Val(xlsAp.Version) < 12 Then
'        xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'    Else
'        xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'    End If
'    xlsAp.Workbooks.Close
'    xlsAp.Quit
'    Set xlsAp = Nothing
'
'    RunExcelChk = True
'    Exit Function
'
'ErrHnd:
'    List1.AddItem "匯入失敗！請通知電腦中心(" & Err.Description & ")", 0
'    '另存
'    If Val(xlsAp.Version) < 12 Then
'        xlsAp.Workbooks(1).SaveAs FileName:=txtFileName & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'    Else
'        xlsAp.Workbooks(1).SaveAs FileName:=txtFileName & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'    End If
'    xlsAp.Workbooks.Close
'    xlsAp.Quit
'    Set xlsAp = Nothing
End Function

'Add by Amy 2021/03/23 國籍空白需Run台灣及非台灣,資料可能重覆,故寫至暫存檔,+顯示業務/業務區
Private Function RunExcelChk() As Boolean
    Dim RsQ As New ADODB.Recordset, RsQ2 As New ADODB.Recordset
    Dim xlsAp As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim strQ As String, strName As String, intQ As Integer, intQ2 As Integer, strTmp As String
    Dim bolFirst As Boolean, bolFName As Boolean '第一次/字首查
    Dim strIns As String 'Add by Amy 2021/03/23
    Dim intField As Integer 'Add by Amy 2021/04/15 比對結果置入開始位置
    Dim strCaseQty As String 'Add by Amy 2024/05/10 更代案件統計
    Dim intArea As Integer, strArea1(3) As String, strArea2(3) As String 'Add by Amy 2024/05/31 P/T 台灣/非台灣統計
    
On Error GoTo ErrHnd
     
'*** Memo by Amy 此處有加欄位,也要確認Form_Load是否要加 ***
    'Modify by Amy 2024/02/16 +狀態
    'Modify by Amy 2024/04/25 原strF改至Form_Load,加 案件統計
    If bolNoAdvertise = True Then
      ReDim strF(7): ReDim intWidth2(7)
      'Modify by Amy 2024/05/10 +更代案件統計
      'Modify by Amy 2024/05/31 +P/T 台灣/非台灣統計
      strF2 = Array("建立日期", "國籍", "編號", "名稱", "狀態", "智權人員", "業務區", "案件統計", "更代案件統計")
      intWidth2 = Array(11, 13, 13, 33, 10, 11, 13, 30, 30)
      'end 2024/05/10
    End If
    'end 2024/02/16
    
    intField = Asc(txtPutField) 'Add by Amy 2021/04/15
    intRow_O = 0: intRow = 2
    RunExcelChk = False
    List1.Clear
    xlsAp.Visible = False
    
    'Modify by Amy 2022/08/17 GetQuerySql 改抓共用function GetSearchNameSql
    '            原畫面國籍空白先Run 中文國家(台灣/中國/香港/澳門),改先Run 非中文,避免匯入資料為英文字取第一個單字回傳太多 ex:AB Food & Beverages Taiwan,  Incorporated
    intRun = 2 '非中文
    'Modify by Amy 2022/08/23 改以勾選判斷Run的規則
'    If txtCountry = 台灣國家代號 Or txtCountry = "013" Or txtCountry = "020" Or txtCountry = "044" Then
'        intRun = 1 '中文
'    End If
    If Chk1(0).Value = 1 Then intRun = 1
    'end 2022/08/23

    List1.AddItem "檢查開始：", 0
    xlsAp.Workbooks.Open txtFileName
    Set wksrpt = xlsAp.Worksheets(xlsAp.ActiveSheet.Name) 'Modify by Amy 2024/02/16 原:xlsAp.Worksheets(1),避免存錯工作表造成錯誤(多工作表且有資料)
    'txtOrgName = ChgSQL(RTrim(UCase(xlsAp.Range("A" & intRow))))
    txtOrgName = RTrim(Replace(UCase(xlsAp.Range("A" & intRow)), "　", " ")) '取代全型空白為空白,後面才好取字(此處改還有1處要改)
    strName = txtOrgName
    'Excel A欄 有空白就離開
    Do While strName <> MsgText(601)
        bolFirst = True
        bolFName = False
        intRow_O = intRow_O + 1
        List1.AddItem "　" & strName, 0
        
Again:
        '*** 畫面國籍 台灣 (全名字首查->去除特取若名稱後有空白抓空白前的字模糊查->名稱+空白+名稱,抓空白後名稱查 有更名前資料,目前最新名稱資料也抓) ***
        'ex:台灣華歌爾股份有限公司 陳文華 ->字首 查(台灣華歌爾股份有限公司 陳文華)->模糊 查(台灣華歌爾股份有限公司 陳文華)
        '                                                               ->去除特取 取空白「前」文字 模糊 查(華歌爾)->去除特取 取空白「後」文字 字首 查(陳文華)
        If intRun = 1 Then
            '** 全名字首查(InStr=1,未拿掉特取字,含對造) **
            'strQ = GetQuerySql(1, strName)
            'Modify by Amy 2023/07/04 +bolNoAdvertise
             strQ = GetSearchNameSql(Me.Name, strName, "=1", True, True, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5, txtCountry, , bolNoAdvertise)
             intQ = 1
             Set RsQ = ClsLawReadRstMsg(intQ, strQ)
             If intQ = 0 Then
                '** 全名模糊比對(InStr>0,拿掉特取字,未含對造) **
                'strName = ReplaceTWName(strName)
                'strQ = GetQuerySql(3, strName, True)
                'Modify by Amy 2023/07/04 +bolNoAdvertise
                strQ = GetSearchNameSql(Me.Name, strName, ">0", True, False, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5, txtCountry, strTxtSpec, bolNoAdvertise)
                intQ = 1
                Set RsQ = ClsLawReadRstMsg(intQ, strQ)
                If intQ = 0 And InStr(strName, " ") > 0 Then
                    'Memo by Amy 2022/08 與Widen、秀玲討論之結果,Excel 有 公司+空白+名稱 者,User 要先將資料修改為「公司」在前
                    '秀玲:原字首比對改為「模糊比對」
                    '** 取空白「前」名稱 模糊比對(InStr>0,拿掉特取字,含對造) **
                    bolFName = True
                    If InStr(strName, " ") > 0 Then
                        strName = Mid(strName, 1, Val(InStr(strName, " ")) - 1)
                    Else
                        strName = ""
                    End If
                    
                    If strName <> MsgText(601) Then
                        'strQ = GetQuerySql(1, strName, True)
                        'Modify by Amy 2023/07/04 +bolNoAdvertise
                        strQ = GetSearchNameSql(Me.Name, strName, ">0", True, True, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5, txtCountry, strTxtSpec, bolNoAdvertise)
                        strName = Replace(txtOrgName, strName & " ", "") '原字串去掉 空白「前」名稱
                        intQ = 1
                        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
                        If intQ = 0 And strName <> MsgText(601) Then
                            '** 取空白「後」第一個「字首比對」(InStr=1,拿掉特取字,不含對造) **
                            bolFName = True
                            If InStr(strName, " ") > 0 Then
                                strName = Mid(strName, 1, Val(InStr(strName, " ")) - 1)
                            End If
                            
                            'strName = ReplaceTWName(strName)
                            'strQ = GetQuerySql(1, strName, True)
                            'Modify by Amy 2023/07/04 +bolNoAdvertise
                            strQ = GetSearchNameSql(Me.Name, strName, "=1", True, False, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5, txtCountry, strTxtSpec, bolNoAdvertise)
                            intQ = 1
                            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
                        End If
                    End If 'strName <> MsgText(601)
                End If 'end 全名模糊比對(InStr>0)
             End If
        '*** End 畫面國籍 台灣 ***
        ElseIf intRun = 2 Then
        '*** 畫面國籍 非台灣 (有更名前資料,目前最新名稱資料也抓) ***
            'ex:Knobbe Martens Olson & Bear ->字首比對 查(Knobbe Martens Olson & Bear) ->取前3個單字 字首比對 查(Knobbe Martens Olson)
            '** 全名 字首查(InStr=1,未拿掉特取字,含對造) **
            'strQ = GetQuerySql(1, strName)
            'Modify by Amy 2023/07/04 +bolNoAdvertise
            strQ = GetSearchNameSql(Me.Name, strName, "=1", True, True, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5, txtCountry, , bolNoAdvertise)
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 0 Then
                'Mark by Amy 2022/08/17 不使用
'                '查詢(like 字首%)
'                bolFName = True
'                strName = Mid(strName, 1, IIf(InStr(strName, " ") = 0, Len(strName) & " ", InStr(strName, " ")))
'                strQ = GetQuerySql(2, strName)
                '** 去特取 取前3個單字 字首比對(InStr=1,拿掉特取字,含對造) **
                strName = GetSearchTxt(strName)
                'Modify by Amy 2023/07/04 +bolNoAdvertise
                strQ = GetSearchNameSql(Me.Name, strName, "=1", True, True, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5, txtCountry, , bolNoAdvertise)
                intQ = 1
                Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            End If
        '*** End 畫面國籍 非台灣 ***
        End If
        
        '查出來有資料寫入暫存檔
        If intQ = 1 Then
            'Modify by Amy 2023/07/04 +bolNoAdvertise
            'Modify by Amy 2024/02/16 +狀態(R009)
            If bolNoAdvertise = False Then
               strIns = "Insert Into R140419 (ID,R001,R002,R003,R004,R005,R006,R009) " & _
                            "Select '" & strUserNum & "',a.* From (" & strQ & ") a " & _
                            "Where FNo not in(Select R001 From R140419 Where ID='" & strUserNum & "' ) "
               cnnConnection.Execute strIns
            '不得宣傳
            Else
               strIns = "Insert Into R140419 (ID,R001,R002,R003,R004,R005,R006,R008,R009) " & _
                            "Select '" & strUserNum & "',FNo,CDate,SalesNo,SalesAreaNo,FName,sField,Na,Status " & _
                            "From (" & strQ & ") a " & _
                            "Where FNo not in(Select R001 From R140419 Where ID='" & strUserNum & "' ) "
               cnnConnection.Execute strIns
            End If
        End If
        
        'Mark by Amy 2022/08/23 以畫面勾選Run 規則
'        '國籍為空,未查到資料,還需再Run 非台灣的資料
'        If txtCountry = MsgText(601) Then
'            If intQ = 0 Then
'                bolFName = False
'                If intRun = 2 Then
'                    intRun = 1
'                    txtOrgName = RTrim(Replace(UCase(xlsAp.Range("A" & intRow)), "　", " ")) '取代全型空白為空白,後面才好取字
'                    strName = txtOrgName
'                    GoTo Again
'                Else
'                    intRun = 2
'                End If
'            '已查到資料,下一筆仍先Run 非台灣的資料
'            Else
'                intRun = 2
'            End If
'        End If
'        'end 2022/08/17
        
        '抓更名資料,產生更名後的資料
        Call InsertNewName
        
        '*** 印資料 ***
        '造字設顏色
        If InStr(txtOrgName, "?") > 0 Then
            strExc(9) = Chr(Asc(txtPutField) + GetValue("名稱"))
            xlsAp.Range("A" & intRow).Interior.ColorIndex = 41 '設置儲存格填充色(藍)
            xlsAp.Range("A" & intRow).Interior.tintandshade = 0.5 '設深淺
            'Add by Amy 2025/03/26 查:上海眾啟生物科技有限公司(啟 口在下會顯示?無法查到),因為只查2筆資料以為沒查
            xlsAp.Range(strExc(9) & intRow).Value = txtOrgName
            xlsAp.Range(strExc(9) & intRow).Interior.ColorIndex = 41 '設置儲存格填充色(藍)
            xlsAp.Range(strExc(9) & intRow).Interior.tintandshade = 0.5 '設深淺
            'end 2025/03/26
        End If
        
        '填入查詢的資料(智權人員可能多人,因此檔為參考用,故帶第一個即可-秀玲)
        'Modify by Amy 2021/08/16 +對造改排序 原:Decode(R007,'Y',R001||'Y',R001)
        'Modify by Amy 2023/07/04 +bolNoAdvertise
        'Modify by Amy 2024/02/16 +狀態(R009)
        'Modify by Amy 2024/04/25 +Distinct ex:[弁理士法人 有古特許事務所] 會查到中/日欄位都有會顯示2筆
        If bolNoAdvertise = False Then
            strQ = "Select Distinct R001 as FNo,R005 as FName,R009 as Status,R002 as MailAddr,st02,a0902,Decode(R007,'Y',R001||'Y',R001) " & _
                         "From R140419,Staff,Acc090 "
        Else
            strQ = "Select Distinct Sqldatet(R002) as CDate,Na03,R001 as FNo,R005 as FName,R009 as Status,st02,a0902,Decode(R007,'Y',R001||'Y',R001) " & _
                        "From R140419,Staff,Acc090,Nation "
        End If
        strQ = strQ & "Where ID='" & strUserNum & "' And SubStr(R003,1,5)=st01(+) And Nvl(R004,st15)=a0901(+) "
        '不得宣傳 顯示國籍
        If bolNoAdvertise = True Then strQ = strQ & "And R008=NA01(+) "
        strQ = strQ & "Order by FNo,Decode(R007,'Y',R001||'Y',R001),R005"
        'end 2023/07/04
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            Do While RsQ.EOF = False
                'Add by Amy 2024/05/31 清除strArea1/2陣列資料
                For i = LBound(strArea1) To UBound(strArea1)
                  strArea1(i) = "": strArea2(i) = ""
                Next i
                For i = 0 To UBound(strF2)
                    '不是第一筆需新增列,將資料插入
                    If bolFirst = False And i = 0 Then
                        wksrpt.Range("A" & intRow).Select
                        xlsAp.Selection.EntireRow.Insert
                    End If
                    '欄名對應抓資料
                    strTmp = ""
                    Select Case i
                        Case GetValue("編號")
                            strTmp = "" & RsQ.Fields("FNo")
                            '非最新名稱抓顯示＊
                            'Modify by Amy 2022/08/17 bug-排除聯絡人編號
                            If Right("" & RsQ.Fields("FNo"), 1) <> "0" And InStr("" & RsQ.Fields("FNo"), "-") = 0 Then
                                strTmp = strTmp & "＊"
                            End If
                        'Add by Amy 2023/07/04 不得宣傳 顯示欄位
                        Case GetValue("建立日期")
                           strTmp = "" & RsQ.Fields("CDate")
                        Case GetValue("國籍")
                           strTmp = "" & RsQ.Fields("Na03")
                        'end 2023/07/04
                        Case GetValue("名稱")
                            strTmp = "" & RsQ.Fields("FName")
                        'Add by Amy 2024/02/16 +狀態
                        Case GetValue("狀態")
                            strTmp = "" & RsQ.Fields("Status")
                        Case GetValue("代表號")
                            strTmp = "" & RsQ.Fields("MailAddr")
                        Case GetValue("智權人員")
                            strTmp = "" & RsQ.Fields("st02")
                        Case GetValue("業務區")
                            strTmp = "" & RsQ.Fields("a0902")
                        'Add by Amy 2024/04/25 +案件統計
                        Case GetValue("案件統計")
                           'Modify by Amy 2024/05/10 更代資料分開顯示
                           strCaseQty = ""
                           If Left("" & RsQ.Fields("FNo"), 1) = "Y" And InStr("" & RsQ.Fields("FNo"), "-") = 0 Then
                              'Modify by Amy 2024/05/31 +strArea1/strArea2 P/T 台灣/非台灣統計
                              strTmp = SetCaseStatistic("" & RsQ.Fields("FNo"), strCaseQty, strArea1(), strArea2(), bolShowPTArea)
                           Else
                              strTmp = "NA"
                           End If
                        'Add by Amy 2024/05/10 +更代案件統計
                        Case GetValue("更代案件統計")
                           If Left("" & RsQ.Fields("FNo"), 1) = "Y" And InStr("" & RsQ.Fields("FNo"), "-") = 0 Then
                              strTmp = strCaseQty
                           Else
                              strTmp = "NA"
                           End If
                        'Add by Amy 2024/05/31 +P/T 台灣/非台灣統計
                        Case Else
                           If Left("" & RsQ.Fields("FNo"), 1) = "Y" And InStr("" & RsQ.Fields("FNo"), "-") = 0 And InStr(strF2(i), "台灣") > 0 Then
                              intArea = -1 'ex:Y31435000 有TS資料
                              Select Case Replace(strF2(i), "-更", "")
                                 Case "P台灣"
                                    intArea = 0
                                 Case "P非台灣"
                                    intArea = 1
                                 Case "T台灣"
                                    intArea = 2
                                 Case "T非台灣"
                                    intArea = 3
                              End Select
                              If intArea >= 0 Then
                                 If InStr(strF2(i), "-更") > 0 Then
                                    strTmp = strArea2(intArea)
                                 Else
                                    strTmp = strArea1(intArea)
                                 End If
                              End If
                           End If
                    End Select
                    'Modify by Amy 2021/04/15 原:Chr(UBound(strF) + 66) ,改為使用者輸入
                    wksrpt.Range(Chr(intField + i) & intRow).Value = strTmp
                Next i
                intRow = intRow + 1
                If bolFirst = True Then bolFirst = False
               
                RsQ.MoveNext
            Loop
        Else
            intRow = intRow + 1
        End If
        '*** End 印資料 ***
        '刪暫存檔資料
        Call DelR140419
        
        'Modify by Amy 2022/08/17
        txtOrgName = RTrim(Replace(UCase(xlsAp.Range("A" & intRow)), "　", " ")) '取代全型空白為空白,後面才好取字(此處改還有1處要改)
        strName = txtOrgName
    Loop
    
    '設定欄位名稱/欄寬
    'Modify by Amy 2021/04/15 原:UBound(strF) + 66 改使用者輸入結果置入開始位置
    For i = 0 To UBound(strF2)
        'Modify by Amy 2024/05/31 拿掉-更 字樣
        wksrpt.Range(Chr(intField + i) & "1").Value = Replace(strF2(i), "-更", "")
        wksrpt.Range(Chr(intField + i) & "1").ColumnWidth = intWidth2(i)
    Next i
    wksrpt.Range(Chr(intField) & "1:" & Chr(intField + UBound(strF2)) & "1").Interior.ColorIndex = 44
    List1.AddItem "檢查完成 " & intRow_O & " 筆", 0
    
    '另存
    'Modify by Amy 2022/07/06 避免匯入之 xlsx 檔案另存成 xls(格式可能與xls無法相容,出現相容性檢查訊息)會彈錯誤,故存成原始副檔名 原:MsgText(43)
    'Modify by Amy 2022/10/04 bug 原xlsx
    If Val(xlsAp.Version) < 12 Then
        '一般活頁簿 (xlWorkbookNormal)
        xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xlsx", ""), ".xls", "") & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=-4143
    Else
        If strExtension = ".xlsx" Then
            '預設活頁簿 (xlWorkbookDefault)
            xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xlsx", ""), ".xls", "") & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=51
        Else
            'Excel 97-2003 活頁簿 (xlExcel8)
            xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xlsx", ""), ".xls", "") & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=56
        End If
    End If
    'end 2022/10/04
    xlsAp.Workbooks.Close
    xlsAp.Quit
    Set xlsAp = Nothing
    
    RunExcelChk = True
    Exit Function
    
ErrHnd:
    List1.AddItem "匯入失敗！請通知電腦中心(" & Err.Description & ")", 0
    MsgBox "資料有誤！請洽電腦中心"
    '另存
    'Modify by Amy 2022/07/06
    If Val(xlsAp.Version) < 12 Then
        xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=-4143
    Else
        If strExtension = ".xlsx" Then
            '預設活頁簿 (xlWorkbookDefault)
            xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=51
        Else
            'Excel 97-2003 活頁簿 (xlExcel8)
            xlsAp.Workbooks(1).SaveAs FileName:=Replace(Replace(txtFileName, ".xls", ""), "xlsx", "") & ACDate(ServerDate) & ServerTime & strExtension, FileFormat:=56
        End If
    End If
    xlsAp.Workbooks.Close
    xlsAp.Quit
    Set xlsAp = Nothing
End Function

Private Sub DelR140419()
    Dim stDel As String
    
    stDel = "Delete From R140419 Where ID='" & strUserNum & "' "
    cnnConnection.Execute stDel
End Sub

'查到更名資料, 新增一筆更名後資料
Private Sub InsertNewName()
    Dim i As Integer
    Dim stIns As String, stQ As String, stField(1) As String
    
    '子查詢
    stQ = "Select R001 From R140419 Where ID='" & strUserNum & "' "
    
    '客戶
    '固定欄位
    stField(0) = "'" & strUserNum & "',cu01||cu02,cu20,cu13,cu12"
    For i = 1 To 3
        Select Case i
            Case 1
                stField(1) = "cu04"
            Case 2
                stField(1) = "cu05||' '||cu88||' '||cu89||' '||cu90"
            Case 3
                stField(1) = "cu06"
        End Select
        'Modify by Amy 2024/02/16 +cu80
        stIns = stIns & IIf(i <> 1, " Union ", "") & _
                    "Select " & stField(0) & "," & stField(1) & ",'" & i & "','Y',cu80 From Customer " & _
                    "Where cu01 in(Select SubStr(R001,1,8) From R140419 " & _
                                            "Where ID='" & strUserNum & "' And SubStr(R001,1,1)='X' And SubStr(R001,9,1)<>'0' And R006='" & i & "' ) " & _
                    "And cu02='0' And cu01||cu02 not in(" & stQ & " And R001=cu01||cu02) "
    Next i
    'Modify by Amy 2024/02/16 +狀態
    stIns = "Insert into R140419 (ID,R001,R002,R003,R004,R005,R006,R007,R009) " & stIns
    cnnConnection.Execute stIns
    
    stIns = ""
    '代理人
    '固定欄位
    stField(0) = "'" & strUserNum & "',fa01||fa02,fa16,'' as SalesNo,'' as SalesArea"
    For i = 1 To 3
        Select Case i
            Case 1
                stField(1) = "fa04"
            Case 2
                stField(1) = "fa05||' '||fa63||' '||fa64||' '||fa65"
            Case 3
                stField(1) = "fa06"
        End Select
        'Modify by Amy 2024/02/16 +fa69
        stIns = stIns & IIf(i <> 1, " Union ", "") & _
                    "Select " & stField(0) & "," & stField(1) & ",'" & i & "','Y',fa69 From Fagent " & _
                    "Where fa01 in(Select SubStr(R001,1,8) From R140419 " & _
                                            "Where ID='" & strUserNum & "' And SubStr(R001,1,1)='Y' And SubStr(R001,9,1)<>'0' And R006='" & i & "' ) " & _
                    "And fa02='0' And fa01||fa02 not in(" & stQ & " And R001=fa01||fa02) "
    Next i
    stIns = "Insert into R140419 (ID,R001,R002,R003,R004,R005,R006,R007,R009) " & stIns
    cnnConnection.Execute stIns
End Sub

'Add by Amy 2021/04/15 資料檢查後回傳資料起始位置
Private Sub txtPutField_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPutField_LostFocus()
    TextInverse txtPutField
End Sub

Private Sub txtPutField_Validate(Cancel As Boolean)
    Dim stInput As String 'Add by Amy 2023/07/04
    
    If txtPutField = MsgText(601) Then Exit Sub
    
    'Add by Amy 2023/07/04
    'Modify by Amy 2024/02/16 +狀態
    'Modify by Amy 2024/04/25 原寫死數字
    stInput = Asc("Z") - (UBound(strF2) + 1)
    If Chk2.Value = vbChecked Then stInput = Asc("Z") - (UBound(strF2) + 2)
    'end 2024/02/16
    'end 2024/04/25
    If Not (Asc(txtPutField) >= Asc("B") And Asc(txtPutField) <= Val(stInput)) Then
    'end 2023/07/04
        Cancel = True
        MsgBox Mid(Label1(2), 1, Len(Label1(2)) - 1) & "輸入錯誤！" & vbCrLf & _
                        "只能是B ~" & Chr(stInput), vbExclamation + vbOKOnly
        txtPutField.SetFocus
    End If
End Sub
'end 2021/04/15

'Add by Amy 2022/08/17
Private Function GetSearchTxt(ByVal stTxt As String) As String
    Dim stTmp As String, stOrgTxt As String
    
    stOrgTxt = Replace(Trim(stTxt), "　", " ") '去掉前後空白,全型空白->半型空白
    '去特取字後,取前3個單字
    For i = LBound(arrTmp) To UBound(arrTmp)
        stTmp = Replace(stOrgTxt, arrTmp(i), "")
    Next i
    If InStr(stTmp, " ") = 0 Then
        GetSearchTxt = stTmp
    Else
        stOrgTxt = stTmp: stTmp = ""
        For i = 1 To 3
            If InStr(stOrgTxt, " ") > 0 Then
                stTmp = Mid(stOrgTxt, 1, InStr(stOrgTxt, " ") - 1)
                GetSearchTxt = GetSearchTxt & " " & stTmp
                stOrgTxt = Replace(stOrgTxt, stTmp & " ", "")
            Else
                Exit For
            End If
        Next i
        GetSearchTxt = Mid(GetSearchTxt, 2)
    End If
End Function

'Add by Amy 2024/04/25 案件統計(參考frm100114_6)
'Modify by Amy 2024/05/10 +stCaseQ,更代資料分開顯示
'Modify by Amy 2024/05/31 +stArea1/stArea1:P/T 台灣/非台灣統計 及 bolShowArea
Private Function SetCaseStatistic(ByVal stNo As String, ByRef stCaseQ As String, ByRef stArea1() As String, ByRef stArea2() As String, ByVal bolShowArea As Boolean) As String
   Dim rsA As New ADODB.Recordset, intA() As Integer, ii As Integer, sta() As String, stData(1) As String 'Modify by 2024/05/31 原:intA(1)/stA(1)
   Dim intArea As Integer 'Add by Amy 2024/05/31
   
   SetCaseStatistic = ""
   stCaseQ = "" 'Add by Amy 2024/05/10
   'Modify by Amy 2024/05/31 +P/T 台灣/非台灣統計
   If bolShowArea = True Then
      ReDim intA(3): ReDim sta(3)
      Call Pub_frm100114_6_StrMenu(strUserNum, Me.Name, stNo, strExc(1), sta(0), sta(1), sta(2), sta(3))
   Else
      ReDim intA(1): ReDim sta(1)
      Call Pub_frm100114_6_StrMenu(strUserNum, Me.Name, stNo, strExc(1), sta(0), sta(1))
   End If
   
   For ii = LBound(sta) To UBound(sta)
      intA(ii) = 1
      Set rsA = ClsLawReadRstMsg(intA(ii), sta(ii))
      If intA(ii) = 1 Then
         rsA.MoveFirst
         Do While Not rsA.EOF
            'P/T 台灣/非台灣統計
            If ii >= 2 Then
               intArea = -1
               Select Case "" & rsA.Fields("CP01")
                  Case "P台灣"
                     intArea = 0
                  Case "P非台灣"
                     intArea = 1
                  Case "T台灣"
                     intArea = 2
                  Case "T非台灣"
                     intArea = 3
               End Select
               If intArea >= 0 Then
                  strExc(2) = rsA.Fields(1)
                  '只取()內的值
                  If ii = 2 Then
                     stArea1(intArea) = Mid(strExc(2), Val(InStr(strExc(2), "(")) + 1)
                     stArea1(intArea) = Replace(stArea1(intArea), ")", "")
                  Else
                     stArea2(intArea) = Mid(strExc(2), Val(InStr(strExc(2), "(")) + 1)
                     stArea2(intArea) = Replace(stArea2(intArea), ")", "")
                  End If
               End If
            '案件統計/更代案件統計
            Else
               stData(ii) = stData(ii) & ";" & rsA.Fields(0) & "-" & rsA.Fields(1)
            End If
            rsA.MoveNext
         Loop
      End If
   Next ii
   If stData(0) <> MsgText(601) Then SetCaseStatistic = Mid(stData(0), 2)
   'Modify by Amy 2024/05/10 更代資料分開顯示
'   If stData(1) <> MsgText(601) Then
'      If SetCaseStatistic <> MsgText(601) Then SetCaseStatistic = SetCaseStatistic & vbCrLf
'      SetCaseStatistic = SetCaseStatistic & Mid(stData(1), 2)
'   End If
   If stData(1) <> MsgText(601) Then stCaseQ = Mid(stData(1), 2)
   'end 2024/05/10
   Set rsA = Nothing
End Function
