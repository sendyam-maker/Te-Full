VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm210140 
   BorderStyle     =   1  '單線固定
   Caption         =   "開拓名單轉入國內潛在客戶作業"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7140
   Begin VB.TextBox TestSpecTxt 
      Height          =   264
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox txtSubject 
      Height          =   264
      Left            =   1440
      TabIndex        =   12
      Top             =   2200
      Width           =   5085
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "寄信"
      Enabled         =   0   'False
      Height          =   405
      Left            =   5400
      TabIndex        =   10
      Top             =   60
      Width           =   750
   End
   Begin VB.TextBox txtSource 
      Height          =   264
      Left            =   1410
      TabIndex        =   6
      Top             =   1560
      Width           =   5085
   End
   Begin VB.TextBox txtFileName 
      Height          =   264
      Left            =   1410
      TabIndex        =   3
      Top             =   1050
      Width           =   5085
   End
   Begin VB.CommandButton CmdOpenFile 
      Caption         =   "<="
      Height          =   345
      Left            =   6540
      TabIndex        =   2
      Top             =   1020
      Width           =   345
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Default         =   -1  'True
      Height          =   405
      Left            =   6200
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdTrans 
      Caption         =   "轉檔"
      Height          =   405
      Left            =   4600
      TabIndex        =   0
      Top             =   60
      Width           =   750
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   130
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "要改為輸入收受者，目前寫死69009"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1170
      TabIndex        =   18
      Top             =   120
      Width           =   3300
   End
   Begin VB.Label Label2 
      Caption         =   "轉檔中請勿使用Word、Excel"
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   5640
      TabIndex        =   16
      Top             =   550
      Width           =   1380
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0 / 0 )"
      Height          =   165
      Left            =   2715
      TabIndex        =   15
      Top             =   3210
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "例：2013/04/01車展"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   1440
      TabIndex        =   13
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "  寄  信  主  旨："
      Height          =   210
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   2200
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注意事項："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   960
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frm210140.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   540
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   4440
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "例：20130401車展-杜副總"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1410
      TabIndex        =   7
      Top             =   1920
      Width           =   2145
   End
   Begin VB.Label Label1 
      Caption         =   "開拓名單來源："
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "潛在客戶檔案："
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1050
      Width           =   1320
   End
End
Attribute VB_Name = "frm210140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/27 Form2.0不用改
'Create By Amy 2013/04/02
Option Explicit

Dim strDocTransFileName As String, strTmp As String, sPoc15Source As String
Dim i As Integer, intWordRow As Integer, intTransRow As Integer, intSalesRow As Integer, intErrRow As Integer
Dim ff1 As Integer, FF2 As Integer 'Add by Amy 2013/11/13 +ff2
Dim strErrMsg As String, m_strFileName1 As String
Dim Txt_Head1 As Boolean, Txt_Head2(1 To 4) As Boolean 'Add by Amy 2013/11/13 +Txt_Head2
Dim strTemp(4) As String
'Modify by Amy 2014/03/27
Dim ZipCode As String, GetAddr As String, Email As String, TelNo As String, intU As Integer
Dim sPoc01 As String, sPoc02 As String, sPoc03 As String, sPoc04 As String, sPoc05 As String, sPoc09 As String, sPoc10 As String, sPoc13 As String, sPoc15 As String
'Add by Amy 2014/05/09
Dim intPoc03 As Integer, intPoc09 As Integer, intPoc10 As Integer, intPoc15 As Integer '潛在客戶基本資料欄位大小(for 避免資料過長無法寫入)
Dim bolOverSize As Boolean '欄位是否過長
Dim bolIsSpecTxt As Boolean '是否為造字

Private Sub cmdExit_Click()
     Unload Me
End Sub

Private Sub CmdOpenFile_Click()
    Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.doc"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "Word檔案 (*.doc 或 *.docx)|*.doc;*.docx" 'Modify by Amy 2014/03/27 +docx
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtFileName.Text = .FileName
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'Add by Amy 2013/11/13 依國籍抓所別產生Txt (國別無資料都寄)
Private Sub cmdSend_Click()
Dim strTo As String
Dim strContent As String
Dim strReportMsg As String

On Error GoTo ErrHnd

 If txtSubject = "" Then
    MsgBox "寄信主旨不可空白！"
    txtSubject.SetFocus
    Exit Sub
 End If
 
strReportMsg = ""
strContent = "Dear Sirs," & Chr(10) & Chr(10) & _
                  "        附件名單為研發處此次開拓資料, 請參閱." & Chr(10) & Chr(10) & _
                  "        若為您開拓之客戶, 或是您既有客戶的關係企業," & Chr(10) & _
                  "請於三日內向電腦中心提出. 經修改後轉為您的潛在客戶," & Chr(10) & _
                  "否則將來接洽案件時就將發放介紹案源獎金給管理人員." & Chr(10) & Chr(10) & _
                  "       敬請配合作業, 謝謝 !" & Chr(10) & Chr(10) & _
                  "                                                          電腦中心"

 Screen.MousePointer = vbHourglass
 For i = 1 To UBound(Txt_Head2)
   Txt_Head2(i) = False
   
   'Modify By Sindy 2015/2/6 mailzip檔案不要了,改統一抓取postzipdata檔案
'   strExc(0) = "Select Poc01||Poc02 as CusNo,Poc03 as Name,Nvl(Poc10,'') as Addr,Nvl(Poc05,'') as TEL1 From PotCustomer1,MailZip " & _
'                    "Where Poc18=TO_CHAR(sysdate,'YYYYMMDD') And Poc13='001-1' And MZ06='" & i & "' And Substr(Poc10,1,3)=MZ01(+) And Poc04=MZ07(+)
   strExc(0) = "Select Poc01||Poc02 as CusNo,Poc03 as Name,Nvl(Poc10,'') as Addr,Nvl(Poc05,'') as TEL1 From PotCustomer1,(select to_multi_byte(substr(PZD01,1,3)) PZD01,PZD10,PZD11 from postzipdata group by substr(PZD01,1,3),PZD10,PZD11) Z " & _
                    "Where Poc18=TO_CHAR(sysdate,'YYYYMMDD') And Poc13='001-1' And PZD10='" & i & "' And Substr(Poc10,1,3)=PZD01(+) And Poc04=PZD11(+) " & _
          "Union Select Poc01||Poc02 as CusNo,Poc03 as Name,Nvl(Poc10,'') as Addr,Nvl(Poc05,'') as TEL1 From PotCustomer1 " & _
                    "Where Poc18=TO_CHAR(sysdate,'YYYYMMDD') And Poc13='001-1' And Poc04 is null "
   '2015/2/6 END
   strExc(0) = "Select * From (" & strExc(0) & ") Order by CusNo"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
          Call ReadTxt2("" & RsTemp("CusNo"), "" & RsTemp("Name"), "" & RsTemp("Addr"), "" & RsTemp("TEL1"), i)
          RsTemp.MoveNext
      Loop
   End If
   
   If Txt_Head2(i) = True Then
      Close FF2
      
      '抓取各所別智權人員
      'Modified by Lydia 2019/08/08 +創新業務部W1001或W2001
      'strExc(0) = "Select st01,st02,st06 所別,st15 From Staff Where st04='1' And st06='" & i & "' And substr(st15,1,2)>='S1' And substr(st15,1,2)<='S4' " & _
                       "And st01>'6' And st01<'F' And substr(st01,4,1)<>'9' Order by st06,st15,st01 "
      strExc(0) = "Select st01,st02,st06 所別,st15 From Staff Where st04='1' And st06='" & i & "' And substr(st15,1,2)>='S1' And substr(st15,1,2)<='S4' " & _
                       "And ((st01>'6' And st01<'F' And substr(st01,4,1)<>'9') or substr(st01,1,1) = 'W') Order by st06,st15,st01 "
       intI = 1: strTo = ""
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          RsTemp.MoveFirst
          Do While Not RsTemp.EOF
              strTo = strTo & RsTemp("st01") & ";"
              RsTemp.MoveNext
          Loop
      End If
      '寄信給智權人員
      If Right(strTo, 1) = ";" Then
          strTo = Left(strTo, Len(strTo) - 1)
          PUB_SendMail strUserNum, strTo, "", "開拓名單轉入國內潛在客戶通知--" & txtSubject, strContent, "", PUB_Getdesktop & "\開拓轉潛在客戶\" & m_strFileName1
          strReportMsg = strReportMsg & Mid(m_strFileName1, InStr(m_strFileName1, "-") + 1, 2) & "、"
      End If
      
   End If
 Next i
 If Right(strReportMsg, 1) = "、" Then MsgBox (" 資料已mail給" & Left(strReportMsg, Len(strReportMsg) - 1) & " 智權人員！")
 Screen.MousePointer = vbDefault
 Exit Sub
 
ErrHnd:
    Close FF2
    MsgBox "產生txt失敗！" & vbCrLf & Err.Description
   Screen.MousePointer = vbDefault
End Sub

Private Sub CmdTrans_Click()
 
On Error GoTo ErrHnd

 If txtFileName = "" Then
    MsgBox "檔案不可空白！"
    txtFileName.SetFocus
    Exit Sub
 End If
 If txtSource = "" Then
    MsgBox "開拓名單來源不可空白！"
    txtSource.SetFocus
    Exit Sub
 End If
 
 strDocTransFileName = txtFileName.Text
 sPoc15Source = txtSource.Text
 
 Dim strIns As String, strTp As String
 
 'Add by Amy  2013/09/09增加產生重覆資料Excel
 Dim xlsAgentPoint As New Excel.Application
 Dim wksrpt As New Worksheet
 Dim intExcel As Integer, jj As Integer
 Dim ExcelReport As String
 Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
 'end 2013/09/09
 Dim intUpdateRow As Integer  'Add by Amy 2014/03/27
 'Add by Amy 2015/03/23
 Dim bolUpd As Boolean  '資料庫是否已有資料/是否更新(業務助理建的資料)
 Dim strCusList As String, strCusStatus As String, strSales As String '已存在資料的 客戶編號或案號/何處資料/智權人員
 Dim strArrList() As String, strArrStatus() As String, strArrSales() As String, ss As Integer
 Dim intArr As Integer '目前搜尋的陣列
 Dim strCheckWay As String 'Add by Amy 2021/08/17
 
 Txt_Head1 = False
 Screen.MousePointer = vbHourglass
 CmdTrans.Enabled = False
 CmdExit.Enabled = False
 Set g_WordAp = New Word.Application
 g_WordAp.Visible = False
 g_WordAp.Documents.Open FileName:=strDocTransFileName
 
 '建立資料夾(For txt)
 If Dir(PUB_Getdesktop & "\開拓轉潛在客戶", vbDirectory) = MsgText(601) Then
      MkDir PUB_Getdesktop & "\開拓轉潛在客戶"
 End If
 
 'Add by Amy  2013/09/09增加產生重覆資料Excel
 If Dir(strExcelPath & "潛在客戶重覆" & ServerDate & MsgText(43)) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
      End If
 Else
            Kill strExcelPath & "潛在客戶重覆" & ServerDate & MsgText(43)
 End If
 
 xlsAgentPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
 xlsAgentPoint.Workbooks.add
 Set wksrpt = xlsAgentPoint.Worksheets(1)
 wksrpt.PageSetup.Orientation = xlLandscape '橫印
 'end 2013/09/09
   
 intTransRow = 0: intSalesRow = 0: intErrRow = 0: intExcel = 1 'Add by Amy 2013/09/09
 intWordRow = g_WordAp.Selection.Tables(1).Rows.Count - 1
 
 '*對造
 strSQL1 = " And CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
 strSQL2 = " And CP01 IN (" & SQLGrpStr("", 1) & ") "
 StrSQL3 = " And CP01 IN (" & SQLGrpStr("", 3) & ") "
 StrSQL4 = " And CP01 IN (" & SQLGrpStr("", 4) & ") "
 strSQL5 = " And CP01 IN (" & SQLGrpStr("", 5) & ") "
           
 'Add by Amy 2014/03/03 +bar
  ProgressBar1.Min = 0
  ProgressBar1.max = intWordRow
  ProgressBar1.Value = 0
  ProgressBar1.Visible = True
  lblCount.Visible = True
  DoEvents
 'end 2014/03/03
 'cnnConnection.BeginTrans 'Mark by Amy 2017/09/11 避免資料被lock
 For i = 1 To intWordRow
    strSql = "": GetAddr = "": Email = "": TelNo = "": strErrMsg = ""
    sPoc01 = "": sPoc02 = "": sPoc03 = "": sPoc04 = "": sPoc05 = "": sPoc09 = "": sPoc10 = "": sPoc13 = "": sPoc15 = ""
    ZipCode = "" 'Add by Amy 2013/06/10
    bolUpd = False 'Add by Amy 2015/03/23
    
    '#抓取word資料
    Call RunWordFind(i) '執行尋找
    sPoc03 = Trim(RunWordReadData("公司名稱", i))
    'Add by Amy 2017/09/11 取代(股)
    sPoc03 = Replace(Replace(sPoc03, "(股)公司", "股份有限公司"), "（股）公司", "股份有限公司")
    sPoc03 = Replace(Replace(sPoc03, "(股)有限公司", "股份有限公司"), "（股）有限公司", "股份有限公司")
    'end 2017/09/11
    sPoc13 = Trim(RunWordReadData("本所智權", i)) '本所智權有值表已是客戶
    TelNo = Trim(RunWordReadData("電話", i))
    Email = Trim(RunWordReadData("Email", i))
    GetAddr = Trim(Replace(RunWordReadData("地址", i), Chr(32), ""))
           
    'Modify by Amy 2014/03/27 +本所智權若為 業務助理or日期(第一次開拓日),依名稱更新Email,地址,電話
    If sPoc13 = "" Or sPoc13 = "業務助理" Or IsDate(sPoc13) Then
        '#切割地址
        If GetAddr <> MsgText(601) Then
            sPoc10 = CutAddr(GetAddr, ZipCode)
        End If
        
        'Modify by Amy 2015/03/23 因為檢查與實際轉檔有時間差,故轉入時需再檢查一次是否已有資料
        If sPoc13 = "" Then
            strCusList = GetHasDataList(strCusStatus, strSales)
            If InStr(strSales, "業務助理") > 0 Then
                bolUpd = True
                sPoc01 = Mid(strSales, InStr(strSales, "業務助理") + 4, 8)
                Call ChkErrData(sPoc01, sPoc10, 1)
            ElseIf strCusList <> "" Then
                '檢查時不為客戶,轉檔時已是客戶,只產生重覆資料ReportExcel
            Else
                sPoc13 = "001-1"
                '#自動編潛在客戶編號
                If ClsPDGetAutoNumber("R", strTmp, True, False) Then
                    strTmp = "R" + Right(strTmp, 5) & "00"
                    sPoc01 = strTmp
                    sPoc02 = "0"
                End If
                'Moidfy by Amy 2014/05/09 判斷名稱是否過長是否有造字
                bolIsSpecTxt = False
                TestSpecTxt = sPoc03
                If InStr(TestSpecTxt, "?") > 0 Then bolIsSpecTxt = True
                If CheckLengthIsOK(sPoc03, intPoc03, False) = False Then
                    intErrRow = intErrRow + 1
                    Call ReadTxt1(sPoc01, sPoc03, sPoc03, "名稱欄位有誤！" & IIf(bolIsSpecTxt = True, "「資料過長，無法寫入」並請確認是否有造字", "「資料過長，無法寫入」"))
                ElseIf bolIsSpecTxt = True Then
                    intErrRow = intErrRow + 1
                    Call ReadTxt1(sPoc01, sPoc03, sPoc03, "名稱欄位可能有造字，請確認！")
                End If
                'end 2014/05/09
                Call ChkErrData(sPoc01)
            End If
            'end 2015/03/23
            
            'Modify by Amy  2013/03/03 產生重覆資料Report-Excel
            strExc(0) = "Select CU04 as Name,ST02 as Sales,Nvl(CU20,'') as Email,Nvl(CU30||Cu31,Nvl(CU112||CU23,'')) as Addr,Nvl(CU16,'') as TEL1,Nvl(CU17,'') as TEL2,CU01||CU02||Decode(CU02,'0','','＊')||decode(CU111,'Y','$','')||decode(CU121,'Y','●','') as CusNo,'' as Poc04,'' as Poc15 From Customer,Staff ,(Select Distinct CU01 As A1 From Customer Where InStr(cu04,'" & ChgSQL(sPoc03) & "')>0) A Where CU01=A.A1 And CU13=ST01(+) " & _
                   "Union Select PCU08 as Name,ST02 as Sales,Nvl(PCU18,'') as EMail,Nvl(PCU27,'') as Addr,Nvl(PCU13,'') as TEL1,Nvl(PCU14,'') as TEL2,PCU01||PCU02||Decode(PCU02,'0','','＊') as CusNo,'' as Poc04,'' as Poc15 From PotCustomer,Staff, (Select Distinct pcu01 As A1 From PotCustomer Where InStr(pcu08,'" & ChgSQL(sPoc03) & "')>0) A Where PCU01=A.A1 And substr(LTrim(PCU38),1,5)=ST01(+) " & _
                   "Union Select POC03 as Name,ST02 as Sales,Nvl(POC09,'') as EMail,Nvl(POC10,'') as Addr,Nvl(POC05,'') as TEL1,Nvl(POC06,'') as TEL2,POC01||POC02||Decode(POC02,'0','','＊') as CusNo,Poc04,Poc15 From PotCustomer1,Staff, (Select Distinct poc01 As A1 From PotCustomer1 Where instr(poc03,'" & ChgSQL(sPoc03) & "')>0) A Where POC01=A.A1 And POC13=ST01(+) " & _
                   "Union Select FA04 as Name,'' as Sales,Nvl(FA16,'') as EMail,Nvl(FA17,'') as Addr,Nvl(FA12,'') as TEL1,Nvl(FA13,'') as TEL2,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') as CusNo,'' as Poc04,'' as Poc15 From Fagent, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(sPoc03) & "')>0) A Where FA01=A.A1 " & _
                   "Union Select NT02 as Name,ST02 as Sales,'' as EMail,Nvl(NT09,'') as Addr,'' as TEL1,'' as TEL2,NT01||Decode(NT21,null,'⊕','') as CusNo,'' as Poc04,'' as Poc15 From NotAgent,Staff, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & ChgSQL(sPoc03) & "')>0) A Where NT01=A.A1 And NT18=ST01(+) "
            'Add by Amy  +查英、日
            strExc(0) = strExc(0) & " Union " & _
                           "Select CU05||' '||CU88||' '||CU89||' '||CU90 as Name,ST02 as Sales,Nvl(CU20,'') as Email,Nvl(CU30||Cu31,Nvl(CU112||CU23,'')) as Addr,Nvl(CU16,'') as TEL1,Nvl(CU17,'') as TEL2,CU01||CU02||Decode(CU02,'0','','＊')||decode(CU111,'Y','$','')||decode(CU121,'Y','●','') as CusNo,'' as Poc04,'' as Poc15 From Customer,Staff ,(Select Distinct CU01 As A1 From Customer Where instr(Upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(sPoc03)) & "')>0) A Where CU01=A.A1 And CU13=ST01(+) " & _
                 "Union Select CU06 as Name,ST02 as Sales,Nvl(CU20,'') as Email,Nvl(CU30||Cu31,Nvl(CU112||CU23,'')) as Addr,Nvl(CU16,'') as TEL1,Nvl(CU17,'') as TEL2,CU01||CU02||Decode(CU02,'0','','＊')||decode(CU111,'Y','$','')||decode(CU121,'Y','●','') as CusNo,'' as Poc04,'' as Poc15 From Customer,Staff ,(Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(sPoc03) & "')>0) A Where CU01=A.A1 And CU13=ST01(+) " & _
                 "Union Select PCU03||' '||PCU04||' '||PCU05||' '||PCU06 as Name,ST02 as Sales,Nvl(PCU18,'') as EMail,Nvl(PCU27,'') as Addr,Nvl(PCU13,'') as TEL1,Nvl(PCU14,'') as TEL2,PCU01||PCU02||Decode(PCU02,'0','','＊') as CusNo,'' as Poc04,'' as Poc15 From PotCustomer,Staff, (Select Distinct PCU01 As A1 From PotCustomer Where instr(Upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(sPoc03)) & "')>0) A Where PCU01=A.A1 And substr(LTrim(PCU38),1,5)=ST01(+) " & _
                 "Union Select PCU07 as Name,ST02 as Sales,Nvl(PCU18,'') as EMail,Nvl(PCU27,'') as Addr,Nvl(PCU13,'') as TEL1,Nvl(PCU14,'') as TEL2,PCU01||PCU02||Decode(PCU02,'0','','＊') as CusNo,'' as Poc04,'' as Poc15 From PotCustomer,Staff, (Select Distinct PCU01 As A1 From PotCustomer Where instr(pcu07,'" & ChgSQL(sPoc03) & "')>0) A Where PCU01=A.A1 And substr(LTrim(PCU38),1,5)=ST01(+) " & _
                 "Union Select POC23||' '||POC24||' '||POC25||' '||POC26 as Name,ST02 as Sales,Nvl(POC09,'') as EMail,Nvl(POC10,'') as Addr,Nvl(POC05,'') as TEL1,Nvl(POC06,'') as TEL2,POC01||POC02||Decode(POC02,'0','','＊') as CusNo,'' as Poc04,'' as Poc15 From PotCustomer1,Staff, (Select Distinct poc01 As A1 From PotCustomer1 Where instr(Upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(sPoc03)) & "')>0) A Where POC01=A.A1 And POC13=ST01(+) " & _
                 "Union Select POC27 as Name,ST02 as Sales,Nvl(POC09,'') as EMail,Nvl(POC10,'') as Addr,Nvl(POC05,'') as TEL1,Nvl(POC06,'') as TEL2,POC01||POC02||Decode(POC02,'0','','＊') as CusNo,'' as Poc04,'' as Poc15 From PotCustomer1,Staff, (Select Distinct poc01 As A1 From PotCustomer1 Where instr(poc27,'" & ChgSQL(sPoc03) & "')>0) A Where POC01=A.A1 And POC13=ST01(+) "
            strExc(0) = strExc(0) & " Union " & _
                           "Select FA05||' '||FA63||' '||FA64||' '||FA65 as Name,'' as Sales,Nvl(FA16,'') as EMail,Nvl(FA17,'') as Addr,Nvl(FA12,'') as TEL1,Nvl(FA13,'') as TEL2,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') as CusNo,'' as Poc04,'' as Poc15 From Fagent, (Select Distinct FA01 As A1 From Fagent Where instr(Upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(sPoc03)) & "')>0) A Where FA01=A.A1 " & _
                 "Union Select FA06 as Name,'' as Sales,Nvl(FA16,'') as EMail,Nvl(FA17,'') as Addr,Nvl(FA12,'') as TEL1,Nvl(FA13,'') as TEL2,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') as CusNo,'' as Poc04,'' as Poc15 From Fagent, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(sPoc03) & "')>0) A Where FA01=A.A1 " & _
                 "Union Select NT03||' '||NT04||' '||NT05||' '||NT06 as Name,ST02 as Sales,'' as EMail,Nvl(NT09,'') as Addr,'' as TEL1,'' as TEL2,NT01||Decode(NT21,null,'⊕','') as CusNo,'' as Poc04,'' as Poc15 From NotAgent,Staff, (Select Distinct NT01 As A1 From NotAgent Where instr(upper(NT03||' '||NT04||' '||NT05||' '||NT06),'" & UCase(ChgSQL(sPoc03)) & "')>0) A Where NT01=A.A1 And NT18=ST01(+) " & _
                 "Union Select NT07 as Name,ST02 as Sales,'' as EMail,Nvl(NT09,'') as Addr,'' as TEL1,'' as TEL2,NT01||Decode(NT21,null,'⊕','') as CusNo,'' as Poc04,'' as Poc15 From NotAgent,Staff, (Select Distinct NT01 As A1 From Notagent Where instr(NT07,'" & ChgSQL(sPoc03) & "')>0) A Where NT01=A.A1 And NT18=ST01(+) "

            '*聯絡人
            strExc(0) = strExc(0) & " Union " & _
                           "Select PCC05 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(sPoc03) & "')>0) A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' " & _
                 "Union Select PCC05 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(sPoc03) & "')>0) A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) " & _
                 "Union Select PCC05 as Name,'' as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(sPoc03) & "')>0) A,Fagent Where FA01(+)=PCC01 And FA02='0' " & _
                 "Union Select PCC05 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(sPoc03) & "')>0) A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
            'Add by Amy  +查英、日
            strExc(0) = strExc(0) & " Union " & _
                           "Select PCC03 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & UCase(ChgSQL(sPoc03)) & "')>0) A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' " & _
                 "Union Select PCC03 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & UCase(ChgSQL(sPoc03)) & "')>0) A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) " & _
                 "Union Select PCC03 as Name,'' as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & UCase(ChgSQL(sPoc03)) & "')>0) A,Fagent Where FA01(+)=PCC01 And FA02='0' " & _
                 "Union Select PCC03 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & UCase(ChgSQL(sPoc03)) & "')>0) A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
            strExc(0) = strExc(0) & " Union " & _
                           "Select PCC04 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(sPoc03) & "')>0) A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' " & _
                 "Union Select PCC04 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(sPoc03) & "')>0) A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) " & _
                 "Union Select PCC04 as Name,'' as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(sPoc03) & "')>0) A,Fagent Where FA01(+)=PCC01 And FA02='0' " & _
                 "Union Select PCC04 as Name,ST02 as Sales,Nvl(PCC08,'') as Email,Nvl(PCC21||PCC22,'') as Addr,'' as TEL1,'' as TEL2,PCC01||'0-'||PCC02 as CusNo,'' as Poc04,'' as Poc15 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(sPoc03) & "')>0) A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
          
            '*開拓客戶
            strExc(0) = strExc(0) & " Union " & _
                           "Select ecd03||' '||ecd04 as Name,'' as Sales,Nvl(ECD13,'') as Email,ECD05||' '||ECD06||' '||ECD07||' '||ECD08||' '||ECD09 as Addr,'' as TEL1,'' as TEL2,ecd02||'-'||LPAD(ecd01,6,'0') as CusNo,'' as Poc04,'' as Poc15 From ExPandCusDetail ,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd03),'" & ChgSQL(UCase(sPoc03)) & "')>0 or instr(Upper(ecd04),'" & ChgSQL(UCase(sPoc03)) & "')>0) A Where nvl(ecd01,'')||nvl(ecd02,'')=A.A1 " & _
                 "Union Select ecd11||' '||ecd12 as Name,'' as Sales,Nvl(ECD13,'') as Email,ECD05||' '||ECD06||' '||ECD07||' '||ECD08||' '||ECD09 as Addr,'' as TEL1,'' as TEL2,ecd02||'-'||LPAD(ecd01,6,'0') as CusNo,'' as Poc04,'' as Poc15 From ExPandCusDetail ,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd11),'" & ChgSQL(UCase(sPoc03)) & "')>0 or instr(Upper(ecd12),'" & ChgSQL(UCase(sPoc03)) & "')>0) A Where nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
            'Add by Amy 2021/08/31
            '*國內開拓函特定公司不列印者
            strExc(0) = strExc(0) & " Union " & _
                            "Select tbnp01 as Name,'' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,'' as CusNo,'' as Poc04,'' as Poc15 From TMBulletinnp Where InStr(Upper(tbnp01) ,'" & ChgSQL(UCase(sPoc03)) & "')>0"
            'end 2021/08/31
 
            '*對造
            'Modify by Amy 2021/08/17 改共用function
'            '商標
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP40 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP40>' '),TradeMark " & _
'                           "Where InStr(Upper(CP40),'" & ChgSQL(sPoc03) & "')>0 And CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) " & strSQL1 & _
'                 "Union Select CP50 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP50>' '),TradeMark " & _
'                           "Where InStr(Upper(CP50),'" & ChgSQL(sPoc03) & "')>0 And CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) " & strSQL1
'            'Add by Amy  +查英、日
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP41 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP41>' '),TradeMark " & _
'                           "Where InStr(Upper(CP41),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) " & strSQL1 & _
'                 "Union Select CP51 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP51>' '),TradeMark " & _
'                           "Where InStr(Upper(CP51),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) " & strSQL1 & _
'                 "Union Select CP42 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP42>' '),TradeMark " & _
'                           "Where InStr(Upper(CP42),'" & ChgSQL(sPoc03) & "')>0 And CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) " & strSQL1 & _
'                 "Union Select CP52 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP52>' '),TradeMark " & _
'                           "Where InStr(Upper(CP52),'" & ChgSQL(sPoc03) & "')>0 And CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) " & strSQL1
'
'            '專利
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP40 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP40>' '),Patent " & _
'                           "Where InStr(Upper(CP40),'" & ChgSQL(sPoc03) & "')>0 And CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2 & _
'                 "Union Select CP50 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP50>' '),Patent " & _
'                           "Where InStr(Upper(CP50),'" & ChgSQL(sPoc03) & "')>0 And CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2
'            'Add by Amy  +查英、日
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP41 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP41>' '),Patent " & _
'                           "Where InStr(Upper(CP41),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2 & _
'                 "Union Select CP51 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP51>' '),Patent " & _
'                           "Where InStr(Upper(CP51),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2 & _
'                 "Union Select CP42 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP42>' '),Patent " & _
'                           "Where InStr(Upper(CP42),'" & ChgSQL(sPoc03) & "')>0 And CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2 & _
'                 "Union Select CP52 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP52>' '),Patent " & _
'                           "Where InStr(Upper(CP52),'" & ChgSQL(sPoc03) & "')>0 And CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2
'
'            '法務
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP40 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP40>' '),LawCase " & _
'                           "Where InStr(Upper(CP40),'" & ChgSQL(sPoc03) & "')>0 And CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) " & StrSQL3 & _
'                 "Union Select CP50 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP50>' '),LawCase " & _
'                           "Where InStr(Upper(CP50),'" & ChgSQL(sPoc03) & "')>0 And CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) " & StrSQL3
'            'Add by Amy  +查英、日
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP41 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP41>' '),LawCase " & _
'                           "Where InStr(Upper(CP41),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) " & StrSQL3 & _
'                 "Union Select CP51 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP51>' '),LawCase " & _
'                           "Where InStr(Upper(CP51),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) " & StrSQL3 & _
'                 "Union Select CP42 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP42>' '),LawCase " & _
'                           "Where InStr(Upper(CP42),'" & ChgSQL(sPoc03) & "')>0 And CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) " & StrSQL3 & _
'                 "Union Select CP52 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP52>' '),LawCase " & _
'                           "Where InStr(Upper(CP52),'" & ChgSQL(sPoc03) & "')>0 And CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) " & StrSQL3
'
'            '顧問
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP40 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP40>' '),HireCase " & _
'                           "Where InStr(Upper(CP40),'" & ChgSQL(sPoc03) & "')>0 And CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) " & StrSQL4 & _
'                 "Union Select CP50 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP50>' '),HireCase " & _
'                           "Where InStr(Upper(CP50),'" & ChgSQL(sPoc03) & "')>0 And CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) " & StrSQL4
'            'Add by Amy  +查英、日
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP41 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP41>' '),HireCase " & _
'                           "Where InStr(Upper(CP41),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) " & StrSQL4 & _
'                 "Union Select CP51 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP51>' '),HireCase " & _
'                           "Where InStr(Upper(CP51),'" & UCase(ChgSQL(sPoc03)) & "')>0 And CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) " & StrSQL4 & _
'                 "Union Select CP42 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP42>' '),HireCase " & _
'                           "Where InStr(Upper(CP42),'" & ChgSQL(sPoc03) & "')>0 And CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) " & StrSQL4 & _
'                 "Union Select CP52 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP52>' '),HireCase " & _
'                           "Where InStr(Upper(CP52),'" & ChgSQL(sPoc03) & "')>0 And CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) " & StrSQL4
'
'            '服務
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP40 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP40>' '),ServicePractice " & _
'                           "Where InStr(Upper(CP40),'" & ChgSQL(sPoc03) & "')>0 And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & strSQL5 & _
'                 "Union Select CP50 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP50>' '),ServicePractice " & _
'                           "Where InStr(Upper(CP50),'" & ChgSQL(sPoc03) & "')>0 And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & strSQL5
'            'Add by Amy  +查英、日
'            strExc(0) = strExc(0) & " Union " & _
'                           "Select CP41 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP41>' '),ServicePractice " & _
'                           "Where InStr(Upper(CP41),'" & ChgSQL(sPoc03) & "')>0 And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & strSQL5 & _
'                 "Union Select CP51 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP51>' '),ServicePractice " & _
'                           "Where InStr(Upper(CP51),'" & ChgSQL(sPoc03) & "')>0 And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & strSQL5 & _
'                 "Union Select CP42 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP42>' '),ServicePractice " & _
'                           "Where InStr(Upper(CP42),'" & ChgSQL(sPoc03) & "')>0 And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & strSQL5 & _
'                 "Union Select CP52 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as CusNo,'' as Poc04,'' as Poc15 From (Select * From CaseProgress Where CP52>' '),ServicePractice " & _
'                           "Where InStr(Upper(CP52),'" & ChgSQL(sPoc03) & "')>0 And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & strSQL5
'            'end 2014/03/03
            strCheckWay = ">0"
            Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, ChgSQL(sPoc03), strCheckWay, True)
             strExc(0) = strExc(0) & " Union " & _
                                "Select R021002 as Name,'對造' as Sales,'' as Email,'' as Addr,'' as TEL1,'' as TEL2,R021001 as CusNo,'' as Poc04,'' as Poc15 From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004<3 "

            strExc(0) = "Select * From (" & strExc(0) & ") X Order by CusNo"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    
            If intI > 0 Or bolUpd = True Then
                intExcel = intExcel + 1
                wksrpt.Range("a" & intExcel).Value = sPoc03
                wksrpt.Range("b" & intExcel).Value = sPoc13
                wksrpt.Range("c" & intExcel).Value = sPoc09
                wksrpt.Range("d" & intExcel).Value = sPoc10
                wksrpt.Range("e" & intExcel).Value = sPoc05
                If bolUpd = False And strCusList <> "" Then
                    '轉檔檢查時有資料但非國內潛在客戶業助理所建
                    wksrpt.Range("f" & intExcel).Value = "Word轉檔資料"
                Else
                    wksrpt.Range("f" & intExcel).Value = "***"
                End If
                If bolUpd = True Then wksrpt.Range("g" & intExcel).Value = "已是國內潛在客戶資料"
                
                If intI > 0 Then
                    If strCusList <> "" Then
                        strArrList = Split(strCusList, ",")
                        strArrStatus = Split(strCusStatus, ",")
                        strArrSales = Split(strSales, ",")
                    End If
                    intArr = 0
                    For jj = 0 To RsTemp.RecordCount - 1
                        intExcel = intExcel + 1
                        wksrpt.Range("a" & intExcel).Value = Trim(RsTemp.Fields("Name"))
                        wksrpt.Range("b" & intExcel).Value = Trim(RsTemp.Fields("Sales"))
                        wksrpt.Range("c" & intExcel).Value = Trim(RsTemp.Fields("Email"))
                        wksrpt.Range("d" & intExcel).Value = Trim(RsTemp.Fields("Addr"))
                        If Trim(RsTemp.Fields("TEL1")) <> "" And Trim(RsTemp.Fields("TEL2")) <> "" Then
                            wksrpt.Range("e" & intExcel).Value = Trim(RsTemp.Fields("TEL1")) & "," & Trim(RsTemp.Fields("TEL2"))
                        Else
                            wksrpt.Range("e" & intExcel).Value = Trim(RsTemp.Fields("TEL1")) & Trim(RsTemp.Fields("TEL2"))
                        End If
                        wksrpt.Range("f" & intExcel).Value = RsTemp.Fields("CusNo")
                        If strCusList <> "" Then
                            For ss = intArr To UBound(strArrList)
                                strTp = RsTemp.Fields("Cusno")
                                If RsTemp.Fields("Sales") = "對造" Then strTp = Replace(Replace(strTp, "＊", ""), "●", "")
                                If strTp = strArrList(ss) Then
                                    '國內潛在客戶為業務助理且系統有資料更新Word資料(時間差)
                                    If InStr(strArrSales(ss), "業務助理") > 0 Then
                                        Call UpdRecord(sPoc01, "" & RsTemp("Poc04"), "" & RsTemp("TEL1"), "" & RsTemp("Email"), "" & RsTemp("Addr"), "" & RsTemp("Poc15"))
                                        If intU = 9999 Then
                                            wksrpt.Range("g" & intExcel).Value = "已存在 " & strArrStatus(ss) & "-" & "資料相同不需更新"
                                        ElseIf intU > 0 Then
                                            intSalesRow = intSalesRow + 1
                                            intUpdateRow = intUpdateRow + 1
                                            wksrpt.Range("g" & intExcel).Value = strArrStatus(ss) & " 已存在-" & "更新前資料"
                                            wksrpt.Range("g" & intExcel).Font.Color = RGB(255, 0, 0)
                                        Else
                                            wksrpt.Range("g" & intExcel).Value = "已存在 " & strArrStatus(ss) & "-" & "未更新"
                                        End If
                                    Else
                                        wksrpt.Range("g" & intExcel).Value = "已存在 " & strArrStatus(ss)
                                    End If
                                    intArr = intArr + 1
                                    Exit For
                                End If
                            Next ss
                        End If
                        RsTemp.MoveNext
                    Next jj
                End If
            End If
            'end2013/09/09
      
            If bolUpd = True Or strCusList <> "" Then
                '檢查時不為客戶,轉檔時已是客戶,R且業務助理需更新資料,其他不需更新資料
            Else
                '#mail若有誤仍可新增,電話過長取20個字再將全部記錄於備註
                 strIns = "Insert Into PotCustomer1 (POC01,POC02,POC03,POC04,POC05,POC09,POC10,POC12,POC13,POC15) " & _
                                     "Values('" & sPoc01 & "' , '" & sPoc02 & "' ," & CNULL(ChgSQL(sPoc03)) & "," & _
                                     CNULL(ChgSQL(sPoc04)) & "," & CNULL(ChgSQL(sPoc05)) & "," & CNULL(ChgSQL(sPoc09)) & "," & CNULL(ChgSQL(sPoc10)) & "," & _
                                     strSrvDate(1) & "," & CNULL(ChgSQL(sPoc13)) & "," & "'" & ChgSQL(sPoc15) & "')"
                cnnConnection.Execute strIns
                intTransRow = intTransRow + 1
            End If
             
        Else
            '本所智權為 業務助理or日期 更新資料
            intSalesRow = intSalesRow + 1
            strExc(0) = "Select poc01,poc03,poc04,poc05,poc09,poc10,poc15 From PotCustomer1 Where poc03='" & sPoc03 & "' and poc13='001-1' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                Call ChkErrData(RsTemp(0), "" & RsTemp("poc10"), 1)
                Call UpdRecord(RsTemp("poc01"), "" & RsTemp("poc04"), "" & RsTemp("poc05"), "" & RsTemp("poc09"), "" & RsTemp("poc10"), "" & RsTemp("poc15"))
               
                If intU > 0 Then
                    intUpdateRow = intUpdateRow + 1
                End If
            End If
        End If
    'end 2014/03/27
    'end 2015/03/23
    Else
        intSalesRow = intSalesRow + 1
    End If
    'Add by Amy 2014/03/03
    ProgressBar1.Value = ProgressBar1.Value + 1
    lblCount.Caption = "( " & ProgressBar1.Value & " / " & ProgressBar1.max & " )"
    DoEvents
    'end 2014/03/03
 Next i
 'cnnConnection.CommitTrans'Mark by Amy 2017/09/11 避免資料被lock
  If Txt_Head1 = True Then
      Close ff1
   End If
   'Add by Amy  2013/09/09增加產生重覆資料Excel
  If intExcel > 1 Then
    wksrpt.Range("a1").Value = "公司名稱"
    wksrpt.Range("b1").Value = "業務員"
    wksrpt.Range("c1").Value = "Email"
    wksrpt.Range("d1").Value = "地址"
    wksrpt.Range("e1").Value = "電話"
    wksrpt.Range("f1").Value = "客戶編號"
    wksrpt.Range("g1").Value = "備　註" 'Add by Amy 2015/03/23
    wksrpt.Columns("a:a").ColumnWidth = 26
    wksrpt.Columns("b:b").ColumnWidth = 8.5
    wksrpt.Columns("c:c").ColumnWidth = 19
    wksrpt.Columns("d:d").ColumnWidth = 25
    wksrpt.Columns("e:e").ColumnWidth = 15
    wksrpt.Columns("f:f").ColumnWidth = 10
    wksrpt.Columns("g:g").ColumnWidth = 30 'Add by Amy 2015/03/23
    ExcelReport = "潛在客戶重覆資料已產生！"
  Else
    ExcelReport = ""
  End If
  'end 2013/09/09
 Screen.MousePointer = vbDefault
 'Add by Amy 2014/03/03
 CmdTrans.Enabled = True
 CmdExit.Enabled = True
 ProgressBar1.Visible = False
 lblCount.Visible = False
 'end 2014/03/03
 g_WordAp.ActiveDocument.Close
 g_WordAp.Quit
 Set g_WordAp = Nothing
 
'Add by Amy  2013/09/09增加產生重覆資料Excel
 RsTemp.Close
 'Modify by Amy 2017/09/11 判斷版本
 If Val(xlsAgentPoint.Version) < 12 Then
    xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "潛在客戶重覆" & ServerDate & MsgText(43), FileFormat:=-4143
 Else
    xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "潛在客戶重覆" & ServerDate & MsgText(43), FileFormat:=56
 End If
 xlsAgentPoint.Workbooks.Close
 xlsAgentPoint.Quit
 Set xlsAgentPoint = Nothing
 'end 2013/09/09
     
 MsgBox ("資料匯入完畢！原檔案共 " & intWordRow & " 筆, 有智權人員 " & intSalesRow & " 筆 (業務助理重轉" & intUpdateRow & " 筆), 匯入 " & intTransRow & " 筆, 有問題資料 " & intErrRow & " 筆！" & vbCrLf & ExcelReport & " 請記得按「寄信」發Mail給各所智權人員！")
 'Modify by Amy 2014/03/03 原發給秀玲,改發給楊特助
 'PUB_SendMail strUserNum, "83002", "", "開拓名單轉入國內潛在客戶作業(" & sPoc15Source & ")", "資料匯入完畢！原檔案共 " & intWordRow & " 筆, 有智權人員 " & intSalesRow & " 筆, 匯入 " & intTransRow & " 筆, 有問題資料 " & intErrRow & " 筆！" & ExcelReport
 PUB_SendMail strUserNum, "69009", "", "開拓名單轉入國內潛在客戶作業(" & sPoc15Source & ")", "資料匯入完畢！原檔案共 " & intWordRow & " 筆, 有智權人員 " & intSalesRow & " 筆 (業務助理重轉" & intUpdateRow & " 筆), 匯入 " & intTransRow & " 筆, 有問題資料 " & intErrRow & " 筆！" & IIf(intErrRow > 0, "有問題資料電腦中心會進行修正！", "")

  'Add by Amy 2013/11/13 +if
  If intTransRow > 0 Then
    cmdSend.Enabled = True
  Else
    cmdSend.Enabled = False
  End If
 Exit Sub
   
ErrHnd:
   'cnnConnection.RollbackTrans'Mark by Amy 2017/09/11 避免資料被lock
   g_WordAp.ActiveDocument.Close
   g_WordAp.Quit
   Set g_WordAp = Nothing
   'Add by Amy  2013/09/09增加產生重覆資料Excel
   'Modify by Amy 2017/09/11 判斷版本
    If Val(xlsAgentPoint.Version) < 12 Then
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "潛在客戶重覆" & ServerDate & MsgText(43), FileFormat:=-4143
    Else
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "潛在客戶重覆" & ServerDate & MsgText(43), FileFormat:=56
    End If
   xlsAgentPoint.Workbooks.Close
   xlsAgentPoint.Quit
   Set xlsAgentPoint = Nothing
   'end 2013/09/09
         
   MsgBox "匯入失敗！" & vbCrLf & Err.Description
   Screen.MousePointer = vbDefault
   CmdTrans.Enabled = True
   CmdExit.Enabled = True
End Sub

Private Sub Form_Load()
    'Add by Amy 2014/05/09 +潛在客戶基本資料欄位大小設定
    intPoc03 = 80: intPoc09 = 150: intPoc10 = 80: intPoc15 = 2000
    MoveFormToCenter Me
End Sub

'執行Find
Private Sub RunWordFind(intRow As Integer)
Dim strKey As String
   strKey = "ROW" & String(Len(CStr(intWordRow)) - Len(CStr(intRow)), "0") & intRow
   With g_WordAp
      .Selection.HomeKey Unit:=wdStory
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = strKey
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
      .Selection.Find.Execute
      .Selection.SelectRow
   End With
End Sub

'讀取欄位內容
Private Function RunWordReadData(strTitle As String, intRow As Integer) As String
Dim intCol As Integer
   
   RunWordReadData = ""
   Select Case strTitle
      Case "公司名稱"
         intCol = 1
      Case "本所智權"
        intCol = 2
      Case "Email"
         intCol = 3
      Case "地址"
         intCol = 4
      Case "電話"
         intCol = 5
   End Select
   
   With g_WordAp
      .Selection.Tables(1).Rows(intRow + 1).Select
      .Selection.Cells(intCol).Select
      RunWordReadData = Replace(PUB_StringFilter(Trim(.Selection.Text)), Chr(7), "")
   End With
End Function

'依郵遞區號取國別碼
'Modify by Amy 2013/11/13 改抓MZ07 (國籍)
'台灣只區分 北/中/南/高 MZ06(所別)
Private Function GetTWCode(strZipCode As String) As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    
    GetTWCode = ""
    
   'strSql = "Select MZ06 From MailZip Where MZ01='" & strZipCode & "'"
   'Modify By Sindy 2015/2/6 mailzip檔案不要了,改統一抓取postzipdata檔案
   'strSql = "Select MZ07 From MailZip Where MZ01='" & strZipCode & "'"
   strSql = "select PZD11 from postzipdata where to_multi_byte(substr(PZD01,1,3))='" & strZipCode & "' group by PZD11"
   '2015/2/6 END
   CheckOC
   rs.CursorLocation = adUseClient
   rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rs.RecordCount > 0 Then
        If Not IsNull(rs(0)) Then
'            Select Case rs(0)
'                Case "1"
'                    GetTWCode = "001"
'                Case "2"
'                    GetTWCode = "004"
'                Case "3"
'                     GetTWCode = "007"
'                Case "4"
'                     GetTWCode = "008"
'            End Select
             GetTWCode = "" & rs(0)
        End If
   End If
End Function

Private Function ChangeZIP_StrNum(strNum As String) As String
    '2014/03/27 避免Word 郵遞區號為全型而抓不到 ZipCode所以Num +全型字
    Dim k As Integer
    Dim Num, returnVal As String
    ChangeZIP_StrNum = ""
    For k = 1 To Len(strNum)
        Num = Mid(strNum, k, 1)
        Select Case Num
        Case "1", "１"
            returnVal = "１"
        Case "2", "２"
            returnVal = "２"
        Case "3", "３"
            returnVal = "３"
        Case "4", "４"
            returnVal = "４"
        Case "5", "５"
            returnVal = "５"
        Case "6", "６"
            returnVal = "６"
        Case "7", "７"
            returnVal = "７"
        Case "8", "８"
            returnVal = "８"
        Case "9", "９"
            returnVal = "９"
        Case "0", "０"
            returnVal = "０"
        End Select
        ChangeZIP_StrNum = ChangeZIP_StrNum & returnVal
    Next k
End Function

'資料 失敗記錄檢核表
Private Sub ReadTxt1(strCusNo As String, StrCusName As String, strErrData As String, ByVal strMsg As String)
Dim k As Integer
   
   If Txt_Head1 = False Then '第一次進入印抬頭
      Txt_Head1 = True
      If ff1 > 0 Then Close #ff1
      ff1 = FreeFile
      m_strFileName1 = "開拓名單轉入國內潛在客戶錯誤資料檢核表.txt"
      Open PUB_Getdesktop & "\開拓轉潛在客戶\" & m_strFileName1 For Output As ff1
      Print #ff1, "客戶編號   客戶名稱                         錯誤資料                                                     錯誤訊息"
      Print #ff1, "========== ================================ ================================================== ============================================================= "
   End If
   For k = 1 To 4
      strTemp(k) = ""
   Next k
   
   strTemp(1) = Trim(strCusNo)
   strTemp(2) = Trim(StrCusName)
   strTemp(3) = Trim(strErrData)
   strTemp(4) = Trim(strMsg)
      
   strTemp(1) = convForm(CheckStr(strTemp(1)), 10)
   strTemp(2) = convForm(CheckStr(strTemp(2)), 32)
   strTemp(3) = convForm(CheckStr(strTemp(3)), 50)
   strTemp(4) = convForm(CheckStr(strTemp(4)), 60)
   
   Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4)
End Sub

'Add by Amy 2013/11/13 產生Txt for 各所智權人員
Private Sub ReadTxt2(strCusNo As String, StrCusName As String, strAddr As String, strTEL As String, jj As Integer)
    Dim k As Integer
    Dim strTp As String
   
    strTp = ""
    
   If Txt_Head2(jj) = False Then '第一次進入印抬頭
      Txt_Head2(jj) = True
      If FF2 > 0 Then Close #FF2
      FF2 = FreeFile
      
      Select Case jj
        Case 1
            strTp = "北所"
        Case 2
            strTp = "中所"
        Case 3
            strTp = "南所"
        Case 4
            strTp = "高所"
        Case Else
      End Select
      m_strFileName1 = "開拓名單轉入國內潛在客戶-" & strTp & ".txt"
      Open PUB_Getdesktop & "\開拓轉潛在客戶\" & m_strFileName1 For Output As FF2
      
      Print #FF2, "客戶編號   客戶名稱                         地址                                                         電話"
      Print #FF2, "========== ================================ ============================================================ ============================================== "
   End If
   For k = 1 To 4
      strTemp(k) = ""
   Next k
   
   strTemp(1) = Trim(strCusNo)
   strTemp(2) = Trim(StrCusName)
   strTemp(3) = Trim(strAddr)
   strTemp(4) = Trim(strTEL)
      
   strTemp(1) = convForm(CheckStr(strTemp(1)), 10)
   strTemp(2) = convForm(CheckStr(strTemp(2)), 32)
   strTemp(3) = convForm(CheckStr(strTemp(3)), 60)
   strTemp(4) = convForm(CheckStr(strTemp(4)), 50)
   
   Print #FF2, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4)
End Sub
'end 2013/11/13

Private Function PUB_CheckMail1(strMail As String, strComb As String, strErr As String) As Boolean
   '修改自PUB_CheckMail
   '多個mail addr 分割及判斷,判斷後再合併
   strErr = ""
   strComb = ""
   If UCase(strMail) = "NO" Or Trim(strMail) = "" Or Trim(strMail) = "無" Then
      PUB_CheckMail1 = True
      Exit Function
   End If
   
    Dim D_mail() As String
    Dim k As Integer

    D_mail() = Split(Trim(Replace(strMail, Chr(13) & Chr(10), "")), ";")
    
    For k = 0 To UBound(D_mail)
        If InStr(1, D_mail(k), "@") = 0 Then
            strErr = strErr & "Mail 必需要有 @ 符號！"
            PUB_CheckMail1 = False
        ElseIf InStr(1, D_mail(k), "@") = 1 Or InStr(1, D_mail(k), "@") = Len(D_mail(k)) Then
            strErr = strErr & "@ 符號不可為字首或字尾！"
            PUB_CheckMail1 = False
        ElseIf InStr(1, D_mail(k), ",") > 0 Or InStr(1, D_mail(k), "[") > 0 Or InStr(1, D_mail(k), "]") > 0 Or InStr(1, D_mail(k), "!") > 0 Or InStr(1, D_mail(k), "(") > 1 Or InStr(1, D_mail(k), ")") > 0 Or InStr(1, D_mail(k), "=") > 0 Or InStr(1, D_mail(k), "\") > 0 Or InStr(1, D_mail(k), "/") > 0 Or InStr(1, D_mail(k), "<") > 0 Or InStr(1, D_mail(k), ">") > 0 Or InStr(1, D_mail(k), "$") > 0 Or InStr(1, D_mail(k), "%") > 0 Or InStr(1, D_mail(k), "^") > 0 Or InStr(1, D_mail(k), "*") > 0 Or InStr(1, D_mail(k), """") > 0 Then
            strErr = strErr & "Mail 不允許有 ,、[、]、!、(、)、=、\、/、<、>、$、%、^、*、'""' 符號！"
            PUB_CheckMail1 = False
        ElseIf InStr(InStr(1, D_mail(k), "@"), D_mail(k), Chr(32)) > 0 Then
            strErr = strErr & "Mail 不允許有空白字元"
            PUB_CheckMail1 = False
        Else
            PUB_CheckMail1 = True
        End If
        strComb = strComb & ";" & D_mail(k)
    Next k
    strComb = Mid(strComb, 2) '去掉第一個;
  
End Function

'限定字串長度
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function

'Add by Amy 2014/03/03
'切割郵遞區號及地址(地址字串,回傳郵遞區號)
Private Function CutAddr(ByVal p_Addr As String, ByRef p_ZipCode As String) As String
    Dim intJ As Integer
    Dim strTp As String
        
    For intJ = 1 To Len(p_Addr)
        strTp = Mid(p_Addr, intJ, 1)
        If IsNumeric(strTp) Then
            p_ZipCode = p_ZipCode & strTp
        Else
            Exit For
        End If
    Next intJ
    
    CutAddr = Trim(Right(p_Addr, Len(p_Addr) - Len(p_ZipCode)))
    If Len(p_ZipCode) = 3 Or Len(p_ZipCode) = 5 Then
        p_ZipCode = ChangeZIP_StrNum(p_ZipCode)
    Else
        p_ZipCode = ""
    End If
End Function
'end 2014/03/03

'Add by Amy 2014/03/27 Email、地址、電話、備註資料檢查
'Modify by Amy 2014/05/09 增加欄位大小、造字檢查
'Status=0(新增)/1(重轉)
Private Sub ChkErrData(ByRef sPoc01 As String, Optional ByVal oldPoc10 As String = "", Optional ByVal Status As Integer = 0)
        Dim strA(2) As String, strTp As String
       
        
        If Status = 1 Then strTp = "-重轉-"
        
        '----- 判斷Email與地址
        If Email = MsgText(601) And sPoc10 = MsgText(601) Then
            '潛在客戶維護Email與地址需擇一輸入
            intErrRow = intErrRow + 1
            Call ReadTxt1(sPoc01, sPoc03, Email & sPoc10, strTp & "Email與地址皆空白！")
        Else
            '-- 判斷國籍資料
            If Status = "1" Then
                '重轉資料
                If GetAddr <> oldPoc10 Then
                    sPoc04 = GetTWCode(Left(ZipCode, 3)) '郵遞區號轉全型取國別
                    If sPoc04 = "" Then
                        intErrRow = intErrRow + 1
                        Call ReadTxt1(sPoc01, sPoc03, sPoc10, strTp & "無國籍資料")
                    End If
                End If
            Else
                '新增資料
                If ZipCode = MsgText(601) Then
                    intErrRow = intErrRow + 1
                    Call ReadTxt1(sPoc01, sPoc03, sPoc10, "無國籍資料")
                Else
                    sPoc04 = GetTWCode(Left(ZipCode, 3)) '郵遞區號轉全型取國別
                    If sPoc04 = "" Then
                        intErrRow = intErrRow + 1
                        Call ReadTxt1(sPoc01, sPoc03, sPoc10, "無國籍資料")
                    End If
                End If
            End If
            '--end 判斷國籍資料
            
            '--判斷地址
            strA(0) = "地址": strA(1) = "": bolOverSize = False: bolIsSpecTxt = False
            If sPoc10 <> MsgText(601) Then
                '判斷欄位是否過長
                If CheckLengthIsOK(sPoc10, intPoc10, False) = False Then bolOverSize = True
                '利用TextBox判斷是否有造字
                TestSpecTxt = sPoc10
                If InStr(TestSpecTxt, "?") > 0 Then bolIsSpecTxt = True
                '未改成新制判別(台北縣->新北市)
                If CheckTaiwanAddr(sPoc10, sPoc04, strA(0), False) = False Then
                    intErrRow = intErrRow + 1
                    If bolOverSize = True Then strA(1) = "過長無法寫入！"
                    If bolIsSpecTxt = True Then strA(1) = strA(1) & "可能有造字，請確認！"
                    Call ReadTxt1(sPoc01, sPoc03, sPoc10, strTp & "地址" & strA(1) & strA(0))
                ElseIf bolOverSize = True Then
                    intErrRow = intErrRow + 1
                    Call ReadTxt1(sPoc01, sPoc03, sPoc10, strTp & IIf(bolIsSpecTxt = True, "地址過長無法寫入！並請確認是否有造字！", "地址過長無法寫入！"))
                ElseIf bolIsSpecTxt = True Then
                    intErrRow = intErrRow + 1
                    Call ReadTxt1(sPoc01, sPoc03, sPoc10, strTp & "地址可能有造字，請確認！")
                End If
            End If
            '--end 判斷地址
            
            'Add by Amy 2013/06/10 將郵遞區號寫入地址欄位中
            sPoc10 = ZipCode & IIf(bolOverSize = True, "", sPoc10)
              
            strA(1) = "": bolOverSize = False: bolIsSpecTxt = False
            If Trim(Email) <> MsgText(601) Then
                If Right(Email, 1) = ";" Then
                    Email = Left(Email, Len(Email) - 1) '去掉最後一個;
                End If
                '判斷欄位是否過長
                If CheckLengthIsOK(Email, intPoc09, False) = False Then bolOverSize = True
                '判斷是否有造字
                TestSpecTxt = Email
                If InStr(TestSpecTxt, "?") > 0 Then bolIsSpecTxt = True
                '判斷mail格式是否正確
                If PUB_CheckMail1(Email, sPoc09, strErrMsg) = False Then
                    intErrRow = intErrRow + 1
                    If bolOverSize = True Then strA(1) = "過長無法寫入！"
                    If bolIsSpecTxt = True Then strA(1) = strA(1) & "可能有造字，請確認！"
                    Call ReadTxt1(sPoc01, sPoc03, sPoc09, strTp & "Email" & strA(1) & strErrMsg)
                ElseIf bolOverSize = True Then
                    intErrRow = intErrRow + 1
                    Call ReadTxt1(sPoc01, sPoc03, Email, strTp & IIf(bolIsSpecTxt = True, "「Email資料過長，無法寫入」並請確認是否有造字", "「Email資料過長，無法寫入」"))
                ElseIf bolIsSpecTxt = True Then
                    intErrRow = intErrRow + 1
                    Call ReadTxt1(sPoc01, sPoc03, Email, strTp & "Email欄位可能有造字，請確認！")
                End If
            End If
            If bolOverSize = True Then sPoc09 = ""
        End If
        '----- end 判斷Email與地址
        
        '----- 判斷電話是否超過20個字元
        sPoc05 = PUB_StrToStr_byVal(TelNo, 20)
        If Status = 1 Then
            If TelNo <> sPoc05 Then sPoc15 = CFDate(strSrvDate(2)) & "重轉檔電話：" & TelNo
        Else
            If TelNo <> sPoc05 Then
                sPoc15 = sPoc15Source & ";" & " 電話：" & TelNo
            Else
                sPoc15 = sPoc15Source
            End If
        End If
        '----- end 判斷電話是否超過20個字元
        
        '----- 判斷備註是否過長
        bolIsSpecTxt = False
        '判斷是否有造字
        TestSpecTxt = sPoc15
        If InStr(TestSpecTxt, "?") > 0 Then bolIsSpecTxt = True
        If CheckLengthIsOK(sPoc15, intPoc15, False) = False Then
            intErrRow = intErrRow + 1
            Call ReadTxt1(sPoc01, sPoc03, sPoc15, strTp & IIf(bolIsSpecTxt = True, "備註過長無法寫入！並請確認是否有造字！", "備註過長無法寫入！"))
        ElseIf bolIsSpecTxt = True Then
            intErrRow = intErrRow + 1
            Call ReadTxt1(sPoc01, sPoc03, sPoc15, strTp & "備註欄位可能有造字，請確認！")
        End If
        If bolOverSize = True Then sPoc15 = sPoc15Source
        '----- end 判斷備註是否過長
End Sub
'end 2014/03/27

'Add by Amy 2015/03/23 +檢查是否已有資料,回傳客戶編號
Private Function GetHasDataList(stCusStatus As String, stSales As String) As String
    Dim adoquery As New ADODB.Recordset
    Dim strQ As String, stCusList As String
    
    stCusList = "": stCusStatus = "": stSales = "": GetHasDataList = ""
    strQ = "Select CU04 ,NVL(ST02,CU01) ,CU01||CU02||Decode(CU02,'0','','＊')||decode(CU111,'Y','$','')||decode(CU121,'Y','●','') as CusNo,'客戶基本檔' as StartDate,'' as CP10 From Customer,STAFF WHERE rtrim(ltrim(CU04))='" & sPoc03 & "' And CU13=ST01(+) "
    strQ = strQ & "Union Select PCU08,NVL(ST02,PCU01),PCU01||PCU02||Decode(PCU02,'0','','＊') as CusNo,'國外潛在客戶檔' as StartDate,'' as CP10 From PotCustomer,staff Where rtrim(ltrim(PCU08))='" & sPoc03 & "' And substr(LTrim(PCU38),1,5)=ST01(+) "
    strQ = strQ & "Union Select POC03,NVL(ST02,POC01),POC01||POC02||Decode(POC02,'0','','＊') as CusNo,''||POC12 as StartDate,'' as CP10 From PotCustomer1,staff Where rtrim(ltrim(POC03))='" & sPoc03 & "' And poc13=ST01(+) "
    strQ = strQ & "Union Select FA04,NVL(ST02,FA01||FA02),FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') as CusNo,'代理人基本檔' as StartDate,'' as CP10 From FAGENT,STAFF Where  rtrim(ltrim(FA04))='" & sPoc03 & "' And substr(LTrim(FA94),1,5)=ST01(+) "
    strQ = strQ & "Union Select NT02,NVL(ST02,NT01),NT01||Decode(NT21,null,'⊕','') as CusNo,'不得代理基本檔' as StartDate,'' as CP10 From Notagent,STAFF Where rtrim(ltrim(NT02))='" & sPoc03 & "' And NT18=ST01(+) "
    '聯絡人
    strQ = strQ & "Union Select PCC05,NVL(ST02,PCC01),PCC01||'0-'||PCC02 as CusNo,'聯絡人檔' as StartDate,'' as CP10 From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & sPoc03 & "') A,CUSTOMER,STAFF Where CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
    strQ = strQ & "Union Select PCC05,NVL(ST02,PCC01),PCC01||'0-'||PCC02 as CusNo,'聯絡人檔' as StartDate,'' as CP10 From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & sPoc03 & "' ) A,PotCustomer,STAFF Where PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
    strQ = strQ & "Union Select PCC05,PCC01,PCC01||'0-'||PCC02 as CusNo,'聯絡人檔' as StartDate,'' as CP10 From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & sPoc03 & "' ) A,FAGENT Where FA01(+)=PCC01 AND FA02='0' "
    strQ = strQ & "Union Select PCC05,NVL(ST02,PCC01),PCC01||'0-'||PCC02 as CusNo,'聯絡人檔' as StartDate,'' as CP10 From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & sPoc03 & "') A,PotCustomer1,STAFF Where POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
    '開拓客戶
    strQ = strQ & "Union Select ecd03||' '||ecd04,'外法開拓',ecd02||'-'||LPAD(ecd01,6,'0') as CusNo,'外法開拓' as StartDate,'' as CP10 From ExPandCusDetail Where (ltrim(rtrim(ECD03))='" & sPoc03 & "' Or ltrim(rtrim(ECD04))='" & sPoc03 & "') "
    strQ = strQ & "Union Select ecd11||' '||ecd12,'外法開拓',ecd02||'-'||LPAD(ecd01,6,'0') as CusNo,'外法開拓' as StartDate,'' as CP10 From ExPandCusDetail Where (ltrim(rtrim(ECD11))='" & sPoc03 & "' Or ltrim(rtrim(ECD12))='" & sPoc03 & "') "
    '對造 -中英日(同案號只出現一筆）
    strQ = strQ & "Union Select CP40,'對造',CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CusNo,'CP40' as StartDate,CP10 From (Select * From CaseProgress Where CP40>' ') Where rtrim(ltrim(CP40))='" & sPoc03 & "' "
    strQ = strQ & "Union Select CP50,'其他相關人',CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CusNo,'CP50' as StartDate,CP10 From (Select * From CaseProgress Where CP50>' ') Where rtrim(ltrim(CP50))='" & sPoc03 & "' "
    strQ = strQ & "Union Select CP41,'對造',CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CusNo,'CP41' as StartDate,CP10 From (Select * From CaseProgress Where CP41>' ') Where rtrim(ltrim(CP41))='" & sPoc03 & "' "
    strQ = strQ & "Union Select CP51,'其他相關人',CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CusNo,'CP51' as StartDate,CP10 From (Select * From CaseProgress Where CP51>' ') Where rtrim(ltrim(CP51))='" & sPoc03 & "' "
    strQ = strQ & "Union Select CP42,'對造',CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CusNo,'CP42' as StartDate,CP10 From (Select * From CaseProgress Where CP42>' ') Where rtrim(ltrim(CP42))='" & sPoc03 & "' "
    strQ = strQ & "Union Select CP52,'其他相關人',CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CusNo,'CP52' as StartDate,CP10 From (Select * From CaseProgress Where CP52>' ') Where rtrim(ltrim(CP52))='" & sPoc03 & "' "
    strQ = strQ & "Order by CusNo"
    
    adoquery.CursorLocation = adUseClient
    adoquery.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If adoquery.RecordCount > 0 Then
        Do While Not adoquery.EOF
            stCusList = stCusList & "," & adoquery.Fields("CusNo")
            If adoquery.Fields(1) = "業務助理" Then
                stSales = stSales & "," & adoquery.Fields(1) & adoquery.Fields("CusNo")
            Else
                stSales = stSales & "," & adoquery.Fields(1)
            End If
            stCusStatus = stCusStatus & "," & adoquery.Fields("StartDate") '資料表名or日期
            adoquery.MoveNext
        Loop
    End If
    If stCusList <> "" Then
        GetHasDataList = Right(stCusList, Len(stCusList) - 1)
        stCusStatus = Right(stCusStatus, Len(stCusStatus) - 1)
        stSales = Right(stSales, Len(stSales) - 1)
    End If
    adoquery.Close
End Function

Private Sub UpdRecord(stPoc01 As String, ByVal oldPoc04 As String, ByVal oldPoc05 As String, ByVal oldPoc09 As String, ByVal oldPoc10 As String, ByVal oldPoc15 As String)
    Dim strUpdField As String, NowTime As String
    
    intU = 0
    
    If sPoc05 <> "" & oldPoc05 Then strUpdField = ",poc05 = " & CNULL(ChgSQL(sPoc05))
    If sPoc09 = "" And sPoc10 = "" Then
        'Email與地址都為空不更新資料(txt會產生資料)
    ElseIf sPoc10 <> "" & oldPoc10 Then
        strUpdField = strUpdField & ",poc10=" & CNULL(ChgSQL(sPoc10))
        If sPoc04 = MsgText(601) And sPoc10 = MsgText(601) Then
            '地址改為空,國籍改999(其他)並於備註註原本國籍
            strUpdField = strUpdField & ",poc04='999',poc15='" & oldPoc15 & ";" & CFDate(strSrvDate(2)) & "重轉檔改國籍(原:" & oldPoc04 & ")' "
        ElseIf sPoc04 <> oldPoc04 Then
            '國籍有變需更新
            strUpdField = ",poc04='" & sPoc04 & "'" & strUpdField
        End If
    ElseIf sPoc09 <> "" & oldPoc09 Then
            strUpdField = strUpdField & ",poc09='" & ChgSQL(sPoc09) & "'"
    End If
                
    '重轉電話超過20字寫備註
    If sPoc15 <> "" Then strUpdField = strUpdField & ",poc15='" & oldPoc15 & ";" & sPoc15 & "' "
    
    If strUpdField = "" Then
        intU = 9999
    Else
        strUpdField = Mid(strUpdField, 2) '去掉第一個,
                
        NowTime = CStr(ServerTime)
        strUpdField = "Update PotCustomer1 Set " & strUpdField & _
                    ",poc20='" & strUserNum & "',poc21='" & strSrvDate(1) & "',poc22='" & IIf(Len(NowTime) = 6, Left(NowTime, 4), Left(NowTime, 3)) & "' " & _
                    " Where poc13='001-1' And poc01='" & stPoc01 & "' And poc02='0' "
        cnnConnection.Execute strUpdField, intU
    End If
End Sub
