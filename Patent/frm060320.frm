VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060320 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專未完成核稿明細查詢/列印"
   ClientHeight    =   1820
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1820
   ScaleWidth      =   5430
   Begin VB.TextBox txt2 
      Height          =   270
      Left            =   990
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   1350
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Left            =   990
      MaxLength       =   1
      TabIndex        =   0
      Top             =   630
      Width           =   435
   End
   Begin VB.TextBox txtEP04 
      Height          =   270
      Left            =   990
      MaxLength       =   6
      TabIndex        =   1
      Top             =   990
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3420
      TabIndex        =   3
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4215
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin MSForms.Label lblEP04T 
      Height          =   300
      Left            =   1890
      TabIndex        =   9
      Top             =   990
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "列印別："
      Height          =   180
      Left            =   225
      TabIndex        =   8
      Top             =   1395
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "(1.查詢  2.印表)"
      Height          =   180
      Left            =   1350
      TabIndex        =   7
      Top             =   1395
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "組別：                  ( 1 電子電機 2 化學 3 日文 4 機械設計 5 其他 )"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   675
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核稿人："
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   5
      Top             =   1035
      Width           =   720
   End
End
Attribute VB_Name = "frm060320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; lblEP04T ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'Modified by Morgan 2012/5/15 將查詢與列印功能合併(加列印別選項)
Option Explicit
Dim PLeft(0 To 12) As Integer, iPrint As Integer, iPage As Integer, strTemp(1 To 12) As String
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciPageHeight = 11000, ciLineHeight = 300
Dim m_NameID As String, m_Group As String
Dim m_Orientation As Integer

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub FormPrint()

   Dim stSQL As String, stCon As String, stCon1 As String, iCol As Integer, stSubID As String, iSubCount As Integer
   Dim iOrientation As Integer
   Dim stVTB As String
   
On Error GoTo flgErr
    
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/10 清除查詢印表記錄檔欄位
   If txtEP04 <> "" Then
      'Modify by Morgan 2007/8/6 加外譯編號
      strExc(1) = PUB_GetMapID(txtEP04, 0)
      If strExc(1) <> "" Then
         stCon = stCon & " AND EP04 IN ('" & txtEP04 & "','" & strExc(1) & "')"
      Else
         stCon = stCon & " AND EP04='" & txtEP04 & "'"
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txtEP04 & lblEP04T 'Add By Sindy 2010/12/10
   End If
   
   'Add by Morgan 2007/8/6
   If txt1 <> "" Then
      stCon1 = stCon1 & " AND S1.ST16='" & txt1 & "'"
      pub_QL05 = pub_QL05 & ";" & Left(Label1(0), 3) & txt1 & "( 1 電子電機 2 化學 3 日文 4 機械設計 5 其他 )" 'Add By Sindy 2010/12/10
   End If
   
   'Add by Morgan 2007/9/4 加控制核稿人或承辦人為外專人員的案件
   stCon1 = stCon1 & " AND ( SUBSTR(S1.ST15,1,1)='F' OR SUBSTR(S2.ST15,1,1)='F')"
   
   'Modify by Morgan 2007/8/1 核稿人若為外譯編號時改抓員工編號
   'Modify by Morgan 2007/9/4 加控制外專收文
   'Modify by Morgan 2010/5/4 外專收文改判斷CP12為F部門
   'Modified by Morgan 2013/11/6 +235核對中說格式
   'Modify by Amy 2016/08/24 cp27/cp57 is null 改抓cp158/cp159;cp05+0
   stVTB = "SELECT CP01,CP02,CP03,CP04,CP06,CP07,CP09,CP10,CP14,CP64,NVL(SIM01,EP04) EP04,EP08,EP09" & _
      " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF_IDMAP" & _
      " Where CP05+0>20060000 AND cp158=0 AND cp159=0 and substr(cp12,1,1)='F'" & _
      " AND CP10 IN ('201','235','209','210') AND CP14 IS NOT NULL" & _
      " AND EP02(+)=CP09 AND SIM02(+)=EP04" & stCon
   
   'Modify by Morgan 2005/4/21 加稿人ID
   'Modify by Morgan 2005/4/20 翻譯,檢視中說若有延期加 * 號
   'FCP 未發文、未取消收文、未閉卷且案件性質為{翻譯(201)、製作中說(209)、檢視中說(210)}
   'Modify by Morgan 2005/9/8 核稿期限改抓ep08(原cp48)
   'Modify by Morgan 2007/5/30 改排序為1.部門2.組別3.員工編號
   '2012/1/2 modify by sonia 員工姓名只抓4個字
   strExc(0) = "Select substr(S1.ST02,1,4) R01" & _
            ", substr(S2.ST02,1,4) R02" & _
            ", DECODE(PA10,NULL,'',SUBSTR(PA10,1,4)-1911||'/'||SUBSTR(PA10,5,2)||'/'||SUBSTR(PA10,7,2)) R03" & _
            ", DECODE(INSTR('201,210',X.CP10),0,NULL,DECODE(F.CP01,NULL,NULL,'*'))||X.CP01||'-'||X.CP02||'-'||X.CP03||'-'||X.CP04 R04" & _
            ", RPAD(PA05,16,'　') R05" & _
            ", DECODE(EP09,NULL,'',SUBSTR(EP09,1,4)-1911||'/'||SUBSTR(EP09,5,2)||'/'||SUBSTR(EP09,7,2)) R06" & _
            ", DECODE(EP08,NULL,DECODE(EP09,NULL,'   **',''),SUBSTR(EP08,1,4)-1911||'/'||SUBSTR(EP08,5,2)||'/'||SUBSTR(EP08,7,2)) R07" & _
            ", DECODE(X.CP06,NULL,'',SUBSTR(X.CP06,1,4)-1911||'/'||SUBSTR(X.CP06,5,2)||'/'||SUBSTR(X.CP06,7,2)) R08" & _
            ", DECODE(X.CP07,NULL,'',SUBSTR(X.CP07,1,4)-1911||'/'||SUBSTR(X.CP07,5,2)||'/'||SUBSTR(X.CP07,7,2)) R09" & _
            ", DECODE(PA09,'000',PTM03,PTM04) R10" & _
            ", CPM03 R11" & _
            ", RPAD(X.CP64,10,'　') R12, S1.ST01 R13,S1.ST16,S1.ST15,EP04, X.CP14, EP08, X.CP01, X.CP02, X.CP03, X.CP04, X.CP09" & _
            " From (" & stVTB & ") X, PATENT C, CASEPROPERTYMAP D, PATENTTRADEMARKMAP E, STAFF S1, STAFF S2,CASEPROGRESS F" & _
            " Where PA01(+)=X.CP01 AND PA02(+)=X.CP02 AND PA03(+)=X.CP03 AND PA04(+)=X.CP04 AND PA57 IS NULL AND PA01 IS NOT NULL" & _
            " AND CPM01(+)=X.CP01 AND CPM02(+)=X.CP10" & _
            " AND PTM01(+)='1' AND PTM02(+)=PA08" & _
            " AND S1.ST01(+)=EP04 AND S2.ST01(+)=X.CP14" & stCon1 & _
            " AND F.CP43(+)=X.CP09 AND F.CP10(+)='404'"
            
   'Add by Morgan 2007/8/6 加FG
   '2012/1/2 modify by sonia 員工姓名只抓4個字
   strExc(0) = strExc(0) & " UNION Select substr(S1.ST02,1,4) R01" & _
            ", substr(S2.ST02,1,4) R02" & _
            ", sqldatet(SP10) R03" & _
            ", DECODE(INSTR('201,210',X.CP10),0,NULL,DECODE(F.CP01,NULL,NULL,'*'))||X.CP01||'-'||X.CP02||'-'||X.CP03||'-'||X.CP04 R04" & _
            ", RPAD(SP05,16,'　') R05" & _
            ", sqldatet(EP09) R06" & _
            ", DECODE(EP08,NULL,DECODE(EP09,NULL,'   **',''),sqldatet(EP08)) R07" & _
            ", sqldatet(X.CP06) R08" & _
            ", sqldatet(X.CP07) R09" & _
            ", '' R10" & _
            ", CPM03 R11" & _
            ", RPAD(X.CP64,10,'　') R12, S1.ST01 R13,S1.ST16,S1.ST15,EP04, X.CP14, EP08, X.CP01, X.CP02, X.CP03, X.CP04, X.CP09" & _
            " From (" & stVTB & ") X, SERVICEPRACTICE C, CASEPROPERTYMAP D, STAFF S1, STAFF S2,CASEPROGRESS F" & _
            " Where SP01(+)=X.CP01 AND SP02(+)=X.CP02 AND SP03(+)=X.CP03 AND SP04(+)=X.CP04 AND SP15 IS NULL AND SP01 IS NOT NULL" & _
            " AND CPM01(+)=X.CP01 AND CPM02(+)=X.CP10" & _
            " AND S1.ST01(+)=EP04 AND S2.ST01(+)=X.CP14" & stCon1 & _
            " AND F.CP43(+)=X.CP09 AND F.CP10(+)='404'"
            
   strExc(0) = strExc(0) & " ORDER BY ST15,ST16,EP04,CP14,EP08,CP01,CP02,CP03,CP04,CP09"

   intI = 1

   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/10
      
      If txt2 = "2" Then
         iOrientation = Printer.Orientation
         Printer.Orientation = 2
         stSubID = ""
         
         With adoRecordset
         .MoveFirst
         Do While Not .EOF
             If stSubID <> "" & .Fields("R01") Then
               If stSubID <> "" Then
                  Call PrintSubTotal(iSubCount)
                  Printer.NewPage
               End If
               iPage = 1
               stSubID = "" & .Fields("R01")
               m_NameID = stSubID & "(" & .Fields("R13") & ")"
               '2010/1/8 MODIFY BY SONIA
               'm_Group = PUB_GetFCPGrpName("" & .Fields("ST16"))
               m_Group = PUB_GetFCPGrpName("" & .Fields("ST16"), True)
               '2010/1/8 END
               PrintPageHeader
               PrintPageHeader1
               iSubCount = 0
             End If
             iSubCount = iSubCount + 1
             For iCol = LBound(strTemp) To UBound(strTemp)
                 strTemp(iCol) = "" & .Fields(iCol - 1)
             Next
             PrintDetail
             .MoveNext
         Loop
         End With
         Call PrintSubTotal(iSubCount)
         'Modify by Morgan 2005/4/21 不用印合計
         'Call PrintReportFooter(.RecordCount)
         Printer.EndDoc
         MsgBox "列印完成！"
         Printer.Orientation = iOrientation
         
      Else
         SetGrid adoRecordset
      End If
      
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/10
        ShowNoData
    End If
    
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
    
End Sub

Private Sub SetGrid(p_Rst As ADODB.Recordset)
   With frm060320_1
      .Show
      .grdDataList.Visible = False
      Set .grdDataList.Recordset = p_Rst.Clone
      .grdDataList.FormatString = .grdDataList.FormatString
      For intI = 0 To .grdDataList.Cols - 1
         Select Case intI
            '日期置中
            Case 2, 5, 6, 7, 8
               .grdDataList.ColAlignment(intI) = 4
            '其他靠左
            Case Else
               .grdDataList.ColAlignment(intI) = 1
         End Select
         If intI > 13 Then
            .grdDataList.ColWidth(intI) = 0
         End If
      Next
      .grdDataList.Visible = True
   End With
End Sub

Private Function TxtValidate() As Boolean

   Dim bolCancel As Boolean
   
   bolCancel = False
   Call txtEP04_Validate(bolCancel)
   If bolCancel Then GoTo flgFail
   
   'Added by Morgan 2012/5/15
   If txt2 = "" Then
      MsgBox "列印別不可空白！"
      txt2.SetFocus
      GoTo flgFail
   End If
   'end 2012/5/15
   
   TxtValidate = True

flgFail:

End Function

Private Sub cmdOK_Click()
    Screen.MousePointer = vbHourglass
    If TxtValidate Then FormPrint
    Screen.MousePointer = vbDefault
End Sub

Sub GetPleft()

    Erase PLeft
    PLeft(0) = 500
    '核稿人(1050)
    PLeft(1) = 500
    '承辦人(1050)
    PLeft(2) = PLeft(1) + 1050
    '申請日(1200)
    PLeft(3) = PLeft(2) + 1050
    '本所案號(2000)
    PLeft(4) = PLeft(3) + 1200
    '案件名稱(2100)
    PLeft(5) = PLeft(4) + 2000
    '完稿日(1200)
    PLeft(6) = PLeft(5) + 2100
    '核稿期限(1200)
    PLeft(7) = PLeft(6) + 1200
    '本所期限(1200)
    PLeft(8) = PLeft(7) + 1200
    '法定期限(1200)
    PLeft(9) = PLeft(8) + 1200
    '種類(900)
    PLeft(10) = PLeft(9) + 1200
    '案件性質(1100)
    PLeft(11) = PLeft(10) + 900
    '進度備註
    PLeft(12) = PLeft(11) + 1100
    
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 0)

    iPrint = iPrint + ciLineHeight
    If iPrint >= (ciPageHeight - iExtraLines * ciLineHeight) Then
        iPage = iPage + 1
        Printer.NewPage
        PrintPageHeader
        If bolSubtotal Then
            PrintPageHeader1
            iPrint = iPrint + ciLineHeight
        End If
    End If
    
End Sub

Sub PrintDetail()

    Dim iCol As Integer

    PrintNewLine
    For iCol = LBound(strTemp) To UBound(strTemp)
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(iCol)
    Next
    
End Sub

Sub PrintPageHeader()
    iPrint = ciStartY
    Printer.FontName = "細明體"
    Printer.Font.Size = ciTitleFontSize
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.CurrentX = 5800
    Printer.CurrentY = iPrint
    Printer.Print GetTitleNick & "未完成核稿明細表"
    iPrint = iPrint + 500
    Printer.Font.Size = ciFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    
    Printer.CurrentX = 6500
    Printer.CurrentY = iPrint
    If m_NameID <> "()" Then
      Printer.Print "核稿人：" & m_NameID
    End If
    
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "列印人：" & strUserName
    
    'Add by Morgan 2007/5/30
    If m_Group <> "" Then
      Printer.CurrentX = 6500
      Printer.CurrentY = iPrint
      Printer.Print "　組別：" & m_Group
    End If
    'end 2007/5/30
    
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
    PrintNewLine
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "頁    次：" & str(iPage)
    PrintNewLine
    'Add by Morgan 2005/4/20
    Printer.CurrentX = 6000
    Printer.CurrentY = iPrint
    Printer.Print "* 為已延期, ** 為未完稿無核稿期限"
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    '2005/4/20 end
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "核稿人"
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "承辦人"
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = iPrint
    Printer.Print "申請日"
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = iPrint
    Printer.Print "本所案號"
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = iPrint
    Printer.Print "案件名稱"
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "完稿日"
    Printer.CurrentX = PLeft(7)
    Printer.CurrentY = iPrint
    Printer.Print "核稿期限"
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = iPrint
    Printer.Print "本所期限"
    Printer.CurrentX = PLeft(9)
    Printer.CurrentY = iPrint
    Printer.Print "法定期限"
    Printer.CurrentX = PLeft(10)
    Printer.CurrentY = iPrint
    Printer.Print "種類"
    Printer.CurrentX = PLeft(11)
    Printer.CurrentY = iPrint
    Printer.Print "案件性質"
    Printer.CurrentX = PLeft(12)
    Printer.CurrentY = iPrint
    Printer.Print "進度備註"
    
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub
'列印小計
Private Sub PrintSubTotal(Optional ByVal iRecCount As Integer = 0)
    Call PrintNewLine(True, 2)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "小計： " & iRecCount & " 筆"
    PrintNewLine
End Sub

Private Sub Form_Load()
    Dim strST05 As String
    
    MoveFormToCenter Me
    GetPleft
    
   'Add by Morgan 2007/9/21
   '外專工程師要控管權限
   strST05 = PUB_GetST05(strUserNum)
   Select Case strST05
      Case "39" '外專工程師中級主管只可查該組
         txt1 = PUB_GetStaffST16(strUserNum)
         txt1.Locked = True
      Case "40", "49" '外專工程師只可查本人  'modify by sonia 2024/8/15 加入等級49日外專海外工程師
         txtEP04 = strUserNum
         lblEP04T = strUserName
         txtEP04.Locked = True
   End Select
   
   '輸出
   If Pub_StrUserSt15 = "F21" Then
      txt2 = "1"
      
      'Modified by Morgan 2012/5/15
      'txt1(14).Enabled = False
      '各組主管可列印
      'Removed by Morgan 2012/6/1 改用權限控管
      'If strST05 <> "38" And strST05 <> "42" Then
      '   txt2.Enabled = False
      'End If
      'end 2012/6/1
      'end 2012/5/15
   Else
      txt2 = "2"
   End If
   
   'Added by Morgan 2012/6/1 改用權限控管
   If IsUserHasRightOfFunction(Me.Name, strPrint, False) = False Then
      txt2 = "1"
      txt2.Enabled = False
   End If
   'end 2012/6/1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060320 = Nothing
End Sub

Private Sub txt1_GotFocus()
   TextInverse txt1
   CloseIme
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)    '2008/2/22加4德文組by sonia
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") And KeyAscii <> Asc("5") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txt2_GotFocus()
   TextInverse txt2
   CloseIme
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtEP04_GotFocus()
   TextInverse txtEP04
   CloseIme
End Sub

Private Sub txtEP04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtEP04_Validate(Cancel As Boolean)

    Dim strName As String

    lblEP04T = ""
    If txtEP04 <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetStaff(txtEP04, strName) Then
        '2011/6/21 MODIFY BY SONIA 靜芳說離職人員也可查 A0016
        'If ClsPDGetStaff(txtEP04, strName) Then
        If ClsPDGetStaffN(txtEP04, strName) Then
            lblEP04T = strName
        Else
            lblEP04T = ""
            Cancel = True
            Call txtEP04_GotFocus
        End If
    End If

End Sub
