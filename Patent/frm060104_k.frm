VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_k 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文後產生Outlook草稿"
   ClientHeight    =   2960
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2960
   ScaleWidth      =   8470
   Begin VB.CommandButton cmdOK 
      Caption         =   "產生Outlook"
      Height          =   350
      Index           =   0
      Left            =   6060
      TabIndex        =   0
      Top             =   60
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7215
      TabIndex        =   1
      Top             =   60
      Width           =   1080
   End
   Begin MSForms.Label lblSales 
      Height          =   255
      Left            =   1710
      TabIndex        =   21
      Top             =   1545
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4233;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseProperty 
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   1215
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4233;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   1710
      TabIndex        =   22
      Top             =   900
      Width           =   2445
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4313;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   330
      Left            =   990
      TabIndex        =   2
      Top             =   480
      Width           =   7350
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12965;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   3615
      TabIndex        =   20
      Top             =   180
      Width           =   1965
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請號："
      Height          =   255
      Left            =   2865
      TabIndex        =   19
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   195
      Index           =   0
      Left            =   4275
      TabIndex        =   18
      Top             =   1215
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本所號："
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   16
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblIssue 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   195
      Left            =   2385
      TabIndex        =   15
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   1545
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   1005
      TabIndex        =   13
      Top             =   1545
      Width           =   675
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   3315
      TabIndex        =   12
      Top             =   1215
      Width           =   885
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   900
      TabIndex        =   11
      Top             =   1215
      Width           =   1365
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   870
      TabIndex        =   10
      Top             =   180
      Width           =   1920
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   900
      TabIndex        =   9
      Top             =   900
      Width           =   780
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   900
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   570
      Width           =   915
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   5265
      TabIndex        =   6
      Top             =   1215
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   195
      Left            =   4275
      TabIndex        =   5
      Top             =   900
      Width           =   915
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   9
      Left            =   5265
      TabIndex        =   4
      Top             =   900
      Width           =   645
   End
   Begin VB.Label lblNation 
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   900
      Width           =   2400
   End
End
Attribute VB_Name = "frm060104_k"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created By Sindy 2022/5/7
Option Explicit

Public m_CP09 As String '總收文號
Public m_strRecDate As String '當天報告
Dim strCP65 As String
Dim strPA75 As String
Dim cp() As String, pa() As String
Public strTo As String, strCC As String
Public strSubject As String, strContent As String
Dim strUserNumST52 As String, strUserNumST52Name As String
Dim strSalesST52 As String, strSalesST52Name As String
Dim strSales As String, strFCPHandler As String, strFCPHandlerST52 As String

Private Sub cmdok_Click(Index As Integer)
Dim adoRst As ADODB.Recordset
Dim objOutLook As Object
Dim objMail As Object
Dim nFrm As Form
   
   If Index = 1 Then '結束
      Unload Me
      Exit Sub
   End If
   
   strSubject = "": strContent = "": strTo = "": strCC = ""
   
'   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
'      Exit Sub
'   End If
   
   '程序組發文時Outlook草稿:
   '案件上發文時若是1.經發文室 或 2.電子送件"Y"則自動產出Outlook草稿以供程序人員通知報告或請款
   If cp(118) = "" And (cp(123) = "" Or cp(123) = "N") Then
      Exit Sub
   End If
   
   strFCPHandler = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4)) 'FCP程序(管制)人員
   strFCPHandlerST52 = GetST52(strFCPHandler) '程序人員的二級主管
   
   '操作人員的二級主管
   strUserNumST52 = GetST52(strUserNum)
   strUserNumST52Name = GetPrjSalesNM(strUserNumST52)
   '智權人員
   strSales = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
   strSalesST52 = GetST52(strSales)
   strSalesST52Name = GetPrjSalesNM(strSalesST52)
      
   Call ReadF22Data
   
   '檢查人員是否有休假,若有,副本帶職代
   
   '有內容及收件者才需要呼叫Outlook草稿
   If strContent <> "" And strTo <> "" Then
      '呼叫新郵件：
      Set objOutLook = CreateObject("Outlook.Application")
      Set objMail = objOutLook.CreateItem(0)
      
      '轉HTML格式
      strContent = Replace(strContent, "新細明體", "Times New Roman")
      '&nbsp; 不換行空格
      '&thinsp; 窄空格
      '單純只是想要輸入空白？ &nbsp; 就對了
      '&emsp; 全形空格
      '&ensp; 半形空格
      'strContent = Replace(strContent, "　", "&emsp;") '&emsp; 全形空格
      strContent = Replace(strContent, " ", "&thinsp;") '&ensp; 半形空格
      strContent = Replace(strContent, vbCrLf, "<BR>")
      
      'Added by Morgan 2024/3/5 機械組案件主旨都加【機械設計組】--Sharon
      If pa(1) = "FCP" Then 'Add By Sindy 2024/10/1 +if
         If pa(150) = "4" Then
            strSubject = "【機械設計組】" & strSubject
         End If
      End If
      'end 2024/3/5
                        
'      If TypeName(objOutLook.Assistant) <> "Nothing" Then
'         objOutLook.ActiveWindow.WindowState = 1 '0.最大化 1.視窗小點
'      End If
      With objMail
         '.BodyFormat = 2 '2=olFormatHTML 1=olFormatPlain 3=olFormatRichText
         .To = strTo
         .cc = strCC
         .Subject = strSubject
         .HTMLBody = strContent
         .Display
      End With
      
      Set objMail = Nothing
      Set objOutLook = Nothing
      
      '發Outlook之後請進入中說請款(frm060306-frm060306_1新案)畫面，供程序人員產生定稿:
      '新案翻譯(201)、檢視中說(209)、核對中說格式(235)、製作中說(210)
      If cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
         '檢查表單是否已開啟，若是，則關閉
         For Each nFrm In Forms
            If StrComp(nFrm.Name, "frm060306_1", vbTextCompare) = 0 Then
               Unload frm060306_1
            End If
            If StrComp(nFrm.Name, "frm060306", vbTextCompare) = 0 Then
               Unload frm060306
               Exit For
            End If
         Next
         frm060306.Show
         frm060306.Text1.Text = cp(1)
         frm060306.Text2.Text = cp(2)
         frm060306.Text3.Text = cp(3)
         frm060306.Text4.Text = cp(4)
         frm060306.m_quyNewCase = True
         frm060306.Command1_Click
         If frm060306.MSHFlexGrid1.Rows >= 2 Then
            If frm060306.MSHFlexGrid1.TextMatrix(1, 2) <> "" Then
               frm060306.MSHFlexGrid1.TextMatrix(1, 0) = "v"
               Call frm060306.cmdok_Click(1)
               frm060306_1.Show
            End If
         End If
      ElseIf cp(10) = 讓與 Or cp(10) = 變更 Then
         '檢查表單是否已開啟，若是，則關閉
         For Each nFrm In Forms
            If StrComp(nFrm.Name, "frm060306_4", vbTextCompare) = 0 Then
               Unload frm060306_4
            End If
            If StrComp(nFrm.Name, "frm060306", vbTextCompare) = 0 Then
               Unload frm060306
               Exit For
            End If
         Next
         frm060306.Show
         frm060306.Text1.Text = cp(1)
         frm060306.Text2.Text = cp(2)
         frm060306.Text3.Text = cp(3)
         frm060306.Text4.Text = cp(4)
         frm060306.m_quyAnyCP10 = cp(10)
         frm060306.Command1_Click
         If frm060306.MSHFlexGrid1.Rows >= 2 Then
            If frm060306.MSHFlexGrid1.TextMatrix(1, 2) <> "" Then
               frm060306.MSHFlexGrid1.TextMatrix(1, 0) = "v"
               Call frm060306.cmdok_Click(1)
               frm060306_3.cmdOK(0).Value = True
               Unload frm060306
            End If
         End If
      End If
   End If
   
   Set adoRst = Nothing
End Sub

'**********************************
'程序組
'**********************************
Private Sub ReadF22Data()
Dim m_notREC As String, bolHad203 As Boolean
Dim bolMoney As Boolean '是否已有提供金額
Dim bolCP20isN As Boolean '不向客戶收款
Dim strCP14 As String, strCP09 As String
Dim strCP10 As String, strCP06 As String, strCP07 As String, strNP23 As String, strNP09 As String
Dim intQ As Integer
Dim strNP As String 'Add By Sindy 2022/7/13
Dim strEP04 As String 'Add By Sindy 2022/9/27
Dim strSPecTitle As String 'Added by Lydia 2022/09/27 註記(Murgitroyd案優先); 請在組合主旨strSubject記得串在前面, ex.strSubject = strSPecTitle &
Dim strCP14spec As String, strContEx01 As String 'Added by Lydia 2024/04/18

   'Add By Sindy 2022/9/27
   strExc(0) = "select * from engineerprogress where ep02='" & m_CP09 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strEP04 = "" & RsTemp.Fields("EP04") '核稿工程師
   End If
   '2022/9/27 END
   
   '主旨
   'Modified by Lydia 2022/09/27 拿掉(Murgitroyd案優先),改到前面
   'strSubject = " Our Ref: " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " [INCOM." & Trim(lblCaseField(6)) & "]" & IIf(Left(ChangeCustomerL(strPA75), 8) = "Y2099001", "(Murgitroyd案優先)", "")
   strSubject = " Our Ref: " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " [INCOM." & Trim(lblCaseField(6)) & "]"
   'Added by Lydia 2022/09/27 請將(Murgitroyd案優先)移至主旨前面，以便承辦可以快速辨認(by Bobbie); 判斷改用模組
   If pa(1) = "FCP" Or pa(1) = "P" Then
      strSPecTitle = PUB_GetSetMailSubF2(pa(75))
   ElseIf pa(1) = "FG" Or pa(1) = "PS" Then
      strSPecTitle = PUB_GetSetMailSubF2(pa(26))
   End If
   'end 2022/09/27
   
   Select Case cp(10)
      Case "101", "102", "103", "125" '新案發文
         strTo = strUserNumST52 '發文人員之主管(二級主管)
         strCC = strUserNum '發文人員
         'Modified by Lydia 2022/09/27 +strSPecTitle
         strSubject = strSPecTitle & "【新案已提申" & IIf(m_strRecDate = "Y", "_當天報告", "") & "】請承辦告申日號" & strSubject
         'Added by Lydia 2023/05/19 待實審發文後再單獨進入通知申請案號，所以另外讀取
         'If m_strContactSheetA4 = "" Then 'Mark by Lydia 2023/05/24 原本新案發文會先預設m_strContactSheetA4只有新案的內容，考慮到實審可能一併發文，改抓最新內容; ex.FCP-69595
           m_strContactSheetA4 = PUB_FCPPrintContactSheetA4(False, cp(9), cp(1), cp(2), cp(3), cp(4), cp(10), True)
         'End If
         'end 2023/05/19
         strContent = "To " & strUserNumST52Name & ":" & vbCrLf & _
                      "檢查期限後請轉寄 " & Trim(GetPrjSalesNM(strSales)) & " 及CC: " & strSalesST52Name & "; " & strUserName & "; backup" & vbCrLf & vbCrLf & _
                      "------------------------------------------------------------------------" & vbCrLf & _
                      "受 文 者：" & GetPrjSalesNM(strSales) & vbCrLf & _
                      cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & vbCrLf & _
                      IIf(m_strRecDate = "Y", "※送件當天簡單報告" & vbCrLf, "") & _
                      m_strContactSheetA4 & vbCrLf
      
         'Added by Lydia 2023/05/03 因為先發文新案，所以同時提醒該案亦有同日發文的規費金額
         'Mark by Lydia 2023/05/19 如果新案同時發文實審，則在發文新案時，key通知申請案號時應該要先跳出，待實審發文後再單獨進入通知申請案號=>所以Email有帶出,不用彈訊息
         'If strTo <> "" And strContent <> "" And PUB_ChkCPExist(cp, "416", 1) = True Then
         '    MsgBox "有實審未發文，若一併發文請填入outlook規費金額，若否請刪除。", vbInformation + vbOKOnly
         'End If
         'end 2023/05/19
         'end 2023/05/03
      Case "201", "209", "235", "210" '新案翻譯(201)、檢視中說(209)、核對中說格式(235)、製作中說(210)
         '下一程序有未收文的補文件
         strExc(10) = ""
         strSql = "select * from nextprogress" & _
                 " where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'" & _
                 " and np06 is null AND np07='" & 補文件 & "'"
         CheckOC
         With adoRecordset
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                  strExc(10) = strExc(10) & "  " & Trim(.Fields("np15")) & "(約定期限：" & ChangeWStringToTDateString(.Fields("np23")) & ")" & vbCrLf
                  .MoveNext
               Loop
            End If
         End With
         
         '中說發文時若其他進度尚未請款時要提醒(不管CP16是否有值)
         m_notREC = ""
         strSql = "select cp09,cp10,cpm03 from caseprogress,casepropertymap" & _
                 " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
                 " and cp20 is null AND CP159=0 and cp60 is null and cp01=cpm01(+) and cp10=cpm02(+) " & _
                 " order by cp09"
         CheckOC
         With adoRecordset
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                  If m_notREC <> "" Then m_notREC = m_notREC & "＋"
                  m_notREC = m_notREC & .Fields("cpm03")
                  .MoveNext
               Loop
               bolHad203 = False: bolMoney = False: bolCP20isN = False
               If PUB_ChkCPExist(cp, "203", , strCP09, strCP14) = True Then
                  bolHad203 = True
                  'Add By Sindy 2025/2/6 敏莉提 ex: FCP-072662
                  '若有主動修正（是否向客戶收款欄位"N"）or 已有請款單號 cp60 is not null
                  '走下列 有主動修正:"有"提供金額 的規則
                  '且不用詢問「主動修正『有』提供金額嗎？」
                  strExc(0) = "select * from caseprogress where cp09='" & strCP09 & "'" & _
                              " and (cp20='N' or cp60 is not null)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolCP20isN = True
                     bolMoney = True '已有提供金額
                  End If
                  If bolMoney = False Then
                  '2025/2/6 END
                     If MsgBox("主動修正『有』提供金額嗎？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
                        bolMoney = True
                     End If
                  End If
               End If
               
               'Added by Lydia 2024/05/16 中間程序由內專工程師所處理的案件
               If Mid(strCP14, 4, 1) = "9" Then
                  '收件人改為外專工程師
                  strCP14spec = PUB_GetFCPEngSup(strCP14, , , True)
                  strContEx01 = "承辦人為：" & GetStaffName(strCP14, True) & "，"
               ElseIf strCP14 <> "" Then
                  strCP14spec = strCP14
                  strContEx01 = ""
               End If
               'end 2024/05/16
               
               '無主動修正 或 有主動修正"有"提供金額
               If bolHad203 = False Or bolMoney = True Then
                  strTo = strSales '智權人員
                  strCC = strSalesST52 & ";backup" '智權人員之主管;backup
                  'Modified by Lydia 2022/09/27 +strSPecTitle
                  strSubject = strSPecTitle & "【已送件完成_中說" & IIf(m_strRecDate = "Y", "_當天報告", "") & "】請進行請款" & strSubject
                  'Modify By Sindy 2025/2/6 + And bolCP20isN = False
                  If bolMoney = True And bolCP20isN = False Then
                     strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                                  "1.請進行請款，送件附檔可參卷宗區" & vbCrLf & _
                                  "2.未請款案件性質：" & m_notREC & vbCrLf & _
                                  "3.附上主動修正金額" & vbCrLf & _
                                  "4." & IIf(strExc(10) = "", "文件已齊備", "本案尚缺：" & vbCrLf & strExc(10)) & vbCrLf & _
                                  "5.有/無修正頁" & vbCrLf & _
                                  "6.定稿已產生" & vbCrLf
                                  If Val(cp(152)) > 0 Then
                                    'Modified by Lydia 2023/05/02 +規費金額
                                    strContent = strContent & "7.收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                                  End If
                  Else
                     strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                                  "1.請進行請款，送件附檔可參卷宗區" & vbCrLf & _
                                  "2.未請款案件性質：" & m_notREC & vbCrLf & _
                                  "3." & IIf(strExc(10) = "", "文件已齊備", "本案尚缺：" & vbCrLf & strExc(10)) & vbCrLf & _
                                  "4.有/無修正頁" & vbCrLf & _
                                  "5.定稿已產生" & vbCrLf
                                  If Val(cp(152)) > 0 Then
                                    'Modified by Lydia 2023/05/02 +規費金額
                                    strContent = strContent & "6.收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                                  End If
                  End If
               '有主動修正"無"提供金額
               Else
                  'Modified by Lydia 2024/05/16 strCP14改為strCP14spec
                  strTo = strCP14spec & ";" & strSales '主動修正工程師; 智權人員
                  strCC = PUB_GetFCPEngSup(strCP14spec) & ";" & strSalesST52 & ";backup" '主動修正工程師之主管;智權人員之主管;backup
                  'end 2024/05/16
                  'Modified by Lydia 2022/09/27 +strSPecTitle
                  strSubject = strSPecTitle & "【已送件完成_中說" & IIf(m_strRecDate = "Y", "_當天報告", "") & "】1.請工程師提供金額 2.請承辦進行請款" & strSubject
                  'Modifed by Lydia 2024/05/16 strCP14改為strCP14spec ; "   請提供主動修正金額給承辦">> "   " & strContEx01 & "請提供主動修正金額給承辦"
                  strContent = "To " & GetPrjSalesNM(strCP14spec) & ":" & vbCrLf & _
                               "   " & strContEx01 & "請提供主動修正金額給承辦" & vbCrLf & vbCrLf & _
                               "To " & GetPrjSalesNM(strSales) & ":" & vbCrLf & _
                               IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                               "   待工程師提供主動修正金額，請進行請款流程" & vbCrLf & vbCrLf & _
                               "1.送件附檔可參卷宗區" & vbCrLf & _
                               "2.未請款案件性質：" & m_notREC & vbCrLf & _
                               "3." & IIf(strExc(10) = "", "文件已齊備", "本案尚缺：" & vbCrLf & strExc(10)) & vbCrLf & _
                               "4.有/無修正頁" & vbCrLf & _
                               "5.定稿已產生" & vbCrLf
                               If Val(cp(152)) > 0 Then
                                 'Modified by Lydia 2023/05/02 +規費金額
                                 strContent = strContent & "6.收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                               End If
               End If
            Else
               '此案皆已請款，可直接寄中說
               strTo = strSales '智權人員
               strCC = strSalesST52 & ";backup" '智權人員之主管;backup
               'Modified by Lydia 2022/09/27 +strSPecTitle
               strSubject = strSPecTitle & "【已送件完成_中說" & IIf(m_strRecDate = "Y", "_當天報告", "") & "】請寄中說" & strSubject
               strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                            "1.請承辦寄中說，送件附檔可參卷宗區" & vbCrLf & _
                            "2." & IIf(strExc(10) = "", "文件已齊備", "本案尚缺：" & vbCrLf & strExc(10)) & vbCrLf & _
                            "3.定稿已產生" & vbCrLf
                            If Val(cp(152)) > 0 Then
                              'Modified by Lydia 2023/05/02 +規費金額
                              strContent = strContent & "4.收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                            End If
            End If
         End With
         
      Case 延期
         'If Not IsNull(cp(30)) Then
         If Trim(cp(30)) <> "" Then
            '下一程序檔
            strSql = "select np01,np07 as cp10,np08 as cp06,np09 as cp07,np23,np15 from nextprogress" & _
                     " where np01='" & cp(43) & "' and np22='" & cp(30) & "'"
         Else
            '案件進度檔
            strSql = "select caseprogress.*,'' as np23,'' as np15 from caseprogress" & _
                     " where cp09='" & cp(43) & "'"
         End If
         CheckOC
         strCP14 = cp(14)
         With adoRecordset
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
               strCP10 = .Fields("cp10")
               strCP06 = "" & .Fields("cp06")
               strCP07 = "" & .Fields("cp07")
               strNP23 = "" & .Fields("np23")
               If Trim(cp(30)) = "" Then
                  strCP14 = "" & .Fields("cp14")
               End If
            End If
         End With
         'Modify By Sindy 2022/7/25 新案翻譯(201)、檢視中說(209)、核對中說格式(235)、製作中說(210) ex: FCP-066846 (延期-檢視中說)
         If strCP10 = "201" Or strCP10 = "209" Or strCP10 = "235" Or strCP10 = "210" Then
            'Add By Sindy 2022/9/27 工程師:201.新案翻譯抓核稿工程師
            If strCP10 = "201" And strEP04 <> "" Then
               strTo = IIf(m_strRecDate = "Y", strSales & ";", "") & strEP04 & ";" & strUserNumST52 '智權人員(若勾當天報告時才帶)、核稿工程師、程序主管
               strCC = IIf(m_strRecDate = "Y", strSalesST52 & ";", "") & PUB_GetFCPEngSup(strEP04) & ";backup" '智權人員主管(若勾當天報告時才帶)、核稿工程師之主管;backup
            Else
            '2022/9/27 END
               strTo = IIf(m_strRecDate = "Y", strSales & ";", "") & strCP14 & ";" & strUserNumST52 '智權人員(若勾當天報告時才帶)、工程師、程序主管
               strCC = IIf(m_strRecDate = "Y", strSalesST52 & ";", "") & PUB_GetFCPEngSup(strCP14) & ";backup" '智權人員主管(若勾當天報告時才帶)、工程師之主管;backup
            End If
            'Modified by Lydia 2022/09/27 +strSPecTitle
            strSubject = strSPecTitle & "【已送件完成_延期-" & GetPrjState4(lblCaseField(0), strCP10) & IIf(m_strRecDate = "Y", "_當天報告", "") & "】" & IIf(m_strRecDate = "Y", "承辦請報告", "") & strSubject
            strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                        "1.已向智慧局申請延期" & vbCrLf & _
                        "2.補呈中說之期限：(不得再延)" & vbCrLf
         Else
         '2022/7/25 END
            strTo = strUserNumST52 & ";" & strSales '程序主管、智權人員
            strCC = strSalesST52 & ";backup" '智權人員之主管;backup
            'Modified by Lydia 2022/09/27 +strSPecTitle
            strSubject = strSPecTitle & "【已送件完成_延期-" & GetPrjState4(lblCaseField(0), strCP10) & IIf(m_strRecDate = "Y", "_當天報告", "") & "】請承辦報告新期限" & strSubject
            strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                        "1.告已向智慧局申請延期" & vbCrLf
         End If
         If Val(strNP23) > 0 Then
            strExc(10) = "    約定期限：" & ChangeWStringToTDateString(strNP23) & vbCrLf
         Else
            strExc(10) = "    本所期限：" & ChangeWStringToTDateString(strCP06) & vbCrLf
         End If
         strExc(10) = strExc(10) & "    法定期限：" & ChangeWStringToTDateString(strCP07) & vbCrLf
         If strCP10 = "202" Then '補文件
            strSql = "select np01,np06,np08,np09,np23,np15 from nextprogress" & _
                     " where np01='" & cp(43) & "' and np07='" & 補文件 & "' and np06 is null" & _
                     " order by np23 asc"
            CheckOC
            strNP23 = ""
            strContent = strContent & _
                         "2.本案尚缺文件之期限：" & vbCrLf
            With adoRecordset
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
               If .RecordCount > 0 Then
                  .MoveFirst
                  Do While Not .EOF
                     If strNP23 <> .Fields("np23") Then
                        If strNP23 <> "" Then
                        strContent = strContent & "  約定期限：" & ChangeWStringToTDateString(strNP23) & vbCrLf
                        strContent = strContent & "  法定期限：" & ChangeWStringToTDateString(strNP09) & vbCrLf
                        End If
                        strContent = strContent & "  ■" & convForm(CheckStr(.Fields("np15")), 50) & vbCrLf
                     Else
                        strContent = strContent & "    " & convForm(CheckStr(.Fields("np15")), 50) & vbCrLf
                     End If
                     strNP23 = "" & .Fields("np23")
                     strNP09 = "" & .Fields("np09")
                     .MoveNext
                  Loop
                  strContent = strContent & "  約定期限：" & ChangeWStringToTDateString(strNP23) & vbCrLf
                  strContent = strContent & "  法定期限：" & ChangeWStringToTDateString(strNP09) & vbCrLf
               End If
            End With
            strContent = strContent & "不得再延" & vbCrLf
            
         ElseIf strCP10 = "205" Then '申復
            strContent = strContent & _
                         "2.補呈下列事項之期限：" & vbCrLf & _
                         "  ■申復理由 不得再延" & vbCrLf & strExc(10)
         ElseIf strCP10 = "107" Then '再審
            strContent = strContent & _
                         "2.補呈下列事項之期限：" & vbCrLf & _
                         "  ■再審理由" & vbCrLf & strExc(10)
         ElseIf strCP10 = "501" Then '訴願
            strContent = strContent & _
                         "2.補呈下列事項之期限：" & vbCrLf & _
                         "  ■訴願理由及委任狀" & vbCrLf & strExc(10)
         ElseIf strCP10 = "804" Then '舉發答辨
            strContent = strContent & _
                         "2.補呈下列事項之期限：" & vbCrLf & _
                         "  ■舉發答辨理由" & vbCrLf & strExc(10)
         'Modify By Sindy 2022/7/25
         ElseIf strCP10 = "201" Or strCP10 = "209" Or strCP10 = "235" Or strCP10 = "210" Then
            strContent = strContent & strExc(10)
         '2022/7/25 END
         Else
            strContent = strContent & _
                         "2.補呈下列事項之期限：" & vbCrLf & _
                         "  ■" & GetPrjState4(lblCaseField(0), strCP10) & vbCrLf & _
                         strExc(10)
         End If
   End Select
   
   '其他狀況:
   If strContent = "" And strTo = "" Then
      '承辦人是掛程序人員時
      If PUB_GetST03(cp(14)) = "F22" Then
         intQ = MsgBox("要【通知報告】請按【是】，若要【通知請款】請按【否】？", vbYesNoCancel + vbExclamation + vbDefaultButton1)
         If intQ = vbCancel Then
            Exit Sub
         End If
         If intQ = vbYes Then
            '通知報告
            strTo = strSales '智權人員
            strCC = strSalesST52 & ";backup" '智權人員之主管;backup
            'Modified by Lydia 2022/09/27 +strSPecTitle
            strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(m_strRecDate = "Y", "_當天報告", "") & "】請報告客戶" & strSubject
            strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                         "請承辦報告客戶，送件附檔可參卷宗區" & vbCrLf
         Else
            '通知請款
            strTo = strSales '智權人員
            strCC = strSalesST52 & ";backup" '智權人員之主管;backup
            'Modified by Lydia 2022/09/27 +strSPecTitle
            strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(m_strRecDate = "Y", "_當天報告", "") & "】請進行請款" & strSubject
            strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                         "1.請承辦處理請款，送件附檔可參卷宗區" & vbCrLf & _
                         "2.有/無定稿" & vbCrLf 'Modify By Sindy 2024/9/23 定稿已產生 請改成 有/無定稿
                         If Val(cp(152)) > 0 Then
                           'Modified by Lydia 2023/05/02 +規費金額
                           strContent = strContent & "3.收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                         End If
         End If
         
      '承辦人是掛工程師時
      ElseIf PUB_GetST03(cp(14)) = "F21" Then
         intQ = MsgBox("要【通知報告】請按【是】，若要【通知請款】請按【否】？", vbYesNoCancel + vbExclamation + vbDefaultButton1)
         If intQ = vbCancel Then
            Exit Sub
         End If
         'Added by Lydia 2024/04/18 中間程序由內專工程師所處理的案件
         If Mid(cp(14), 4, 1) = "9" Then
            '收件人改為外專工程師
            strCP14spec = PUB_GetFCPEngSup(cp(14), , , True)
            strContEx01 = "承辦人為：" & GetStaffName(cp(14), True) & "，"
         Else
            strCP14spec = cp(14)
            strContEx01 = ""
         End If
         'end 2024/04/18
         
         If intQ = vbYes Then
            '通知報告(不分工程師組別)
            'Modified by Lydia 2024/04/18 cp(14)改為strCP14spec
            strTo = IIf(m_strRecDate = "Y", strSales & ";", "") & strCP14spec '工程師
            strCC = IIf(m_strRecDate = "Y", strSalesST52 & ";", "") & PUB_GetFCPEngSup(strCP14spec) & ";backup" '工程師之主管;backup
            'end 2024/04/18
            'Modified by Lydia 2022/09/27 +strSPecTitle
            strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(m_strRecDate = "Y", "_當天報告", "") & "】請報告客戶" & strSubject
            'Modified by Lydia 2024/04/18 + strContEx01
            strContent = IIf(m_strRecDate = "Y", "※送件當天報告" & vbCrLf, "") & _
                         "1." & strContEx01 & "請工程師報告客戶，送件附檔可參卷宗區" & vbCrLf & _
                         "2.完成後請通知程序人員：" & GetPrjSalesNM(strFCPHandler) & vbCrLf
         Else
            'Modify By Sindy 2022/7/12
            '若是分割發文，則多判斷下一程序有無(416)實審期限，若有，請加帶備註”5.實審期限　約定期限: 111/07/22　法定期限: 111/07/31”
            'Modify By Sindy 2022/9/27 + 435.續行母案再審
            strSql = "select np01,np06,np08,np09,np23,np15,np07 from nextprogress" & _
                     " where np01='" & cp(9) & "' and np07 in('" & 實體審查 & "','435') and np06 is null"
            CheckOC
            strNP = ""
            With adoRecordset
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
               If .RecordCount > 0 Then
                  If .Fields("np07") = 實體審查 Then
                     strNP = "實審期限　"
                  Else
                     strNP = "續行母案再審　"
                  End If
                  strNP = strNP & "約定期限：" & IIf("" & .Fields("np23") <> "", ChangeWStringToTDateString(.Fields("np23")), "") & "　法定期限：" & IIf("" & .Fields("np09") <> "", ChangeWStringToTDateString(.Fields("np09")), "")
               End If
            End With
            
            '通知請款
            If PUB_GetStaffST16(cp(14)) = "3" Then '日文組
               '當天報告
               If m_strRecDate = "Y" Then
                  'Modified by Lydia 2024/04/18 cp(14)改為strCP14spec
                  strTo = strSales & ";" & strCP14spec '智權人員;工程師
                  strCC = strSalesST52 & ";" & PUB_GetFCPEngSup(strCP14spec) & ";backup" '智權人員之主管;工程師之主管;backup
                  'end 2024/04/18
                  'Modified by Lydia 2022/09/27 +strSPecTitle
                  strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(Val(cp(7)) > 0, "(法限:" & Right(ChangeWStringToTDateString(cp(7)), 5) & ")", "") & "】1.請承辦當天報告 2.請工程師進行請款" & strSubject
                  strContent = "1.請承辦當天報告" & vbCrLf
                              'Add By Sindy 2022/9/27
                              If strNP <> "" Then
                                strContent = strContent & "  " & strNP & vbCrLf
                              End If
                              '2022/9/27 END
                  'Modified by Lydia 2024/04/18 + strContEx01
                  strContent = strContent & _
                               "2." & strContEx01 & "請工程師處理請款，送件附檔可參卷宗區" & vbCrLf & _
                               "3.請款金額請提供給承辦（" & GetPrjSalesNM(strSales) & "）" & vbCrLf
                               If Val(cp(152)) > 0 Then
                                 'Modified by Lydia 2023/05/02 +規費金額
                                 strContent = strContent & "4.收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                               End If
               Else
                  'Modified by Lydia 2024/04/18 cp(14)改為strCP14spec
                  strTo = strCP14spec '工程師
                  strCC = PUB_GetFCPEngSup(strCP14spec) & ";backup" '工程師之主管;backup
                  'end 2024/04/18
                  'Modified by Lydia 2022/09/27 +strSPecTitle
                  strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(Val(cp(7)) > 0, "(法限:" & Right(ChangeWStringToTDateString(cp(7)), 5) & ")", "") & "】請進行請款" & strSubject
                  'Modified by Lydia 2024/04/18 + strContEx01
                  strContent = "1." & strContEx01 & "請工程師處理請款，送件附檔可參卷宗區" & vbCrLf & _
                               "2.請款金額請提供給承辦（" & GetPrjSalesNM(strSales) & "）" & vbCrLf
                               If Val(cp(152)) > 0 Then
                                 'Modified by Lydia 2023/05/02 +規費金額
                                 strContent = strContent & "3.收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                               End If
               End If
                            
            Else '英文組
               If MsgBox("【有附請款信】請按【是】，若【無附請款信】請按【否】？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
                  '有附請款信
                  '當天報告
                  If m_strRecDate = "Y" Then
                     strTo = strSales '智權人員
                     strCC = strSalesST52 & ";backup"  '智權人員之主管;backup
                     'Modified by Lydia 2022/09/27 +strSPecTitle
                     strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(Val(cp(7)) > 0, "(法限:" & Right(ChangeWStringToTDateString(cp(7)), 5) & ")", "") & "】1.當天報告 2.請進行請款" & strSubject
                     strContent = "1.請當天報告" & vbCrLf & _
                                  "2.工程師已備妥請款信及其附件如附檔" & vbCrLf & _
                                  "3.請處理請款，送件附檔可參卷宗區" & vbCrLf
                                  intI = 3
                                  If Val(cp(152)) > 0 Then
                                    intI = intI + 1
                                    'Modified by Lydia 2023/05/02 +規費金額
                                    strContent = strContent & intI & ".收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                                  End If
                                  'Add By Sindy 2022/7/13
                                  If strNP <> "" Then
                                    intI = intI + 1
                                    strContent = strContent & intI & "." & strNP & vbCrLf
                                  End If
                                  '2022/7/13 END
                  Else
                     strTo = strSales '智權人員
                     strCC = strSalesST52 & ";backup"  '智權人員之主管;backup
                     'Modified by Lydia 2022/09/27 +strSPecTitle
                     strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(Val(cp(7)) > 0, "(法限:" & Right(ChangeWStringToTDateString(cp(7)), 5) & ")", "") & "】請進行請款" & strSubject
                     strContent = "1.工程師已備妥請款信及其附件如附檔" & vbCrLf & _
                                  "2.請處理請款，送件附檔可參卷宗區" & vbCrLf
                                  intI = 2
                                  If Val(cp(152)) > 0 Then
                                    intI = intI + 1
                                    'Modified by Lydia 2023/05/02 +規費金額
                                    strContent = strContent & intI & ".收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                                  End If
                                  'Add By Sindy 2022/7/13
                                  If strNP <> "" Then
                                    intI = intI + 1
                                    strContent = strContent & intI & "." & strNP & vbCrLf
                                  End If
                                  '2022/7/13 END
                  End If
               Else
                  '無附請款信
                  '當天報告
                  If m_strRecDate = "Y" Then
                     'Modified by Lydia 2024/04/18 cp(14)改為strCP14spec
                     strTo = strSales & ";" & strCP14spec '智權人員;工程師
                     strCC = strSalesST52 & ";" & PUB_GetFCPEngSup(strCP14spec) & ";backup"  '智權人員之主管;工程師之主管;backup
                     'end 2024/04/18
                     'Modified by Lydia 2022/09/27 +strSPecTitle
                     strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(Val(cp(7)) > 0, "(法限:" & Right(ChangeWStringToTDateString(cp(7)), 5) & ")", "") & "】1.請承辦當天報告 2.請工程師進行請款" & strSubject
                     'Modified by Lydia 2024/04/18 + strContEx01
                     strContent = "1.請承辦當天報告" & vbCrLf & _
                                  "2." & strContEx01 & "請工程師處理請款，送件附檔可參卷宗區" & vbCrLf & _
                                  "3.請款信、金額、附件請提供給承辦（" & GetPrjSalesNM(strSales) & "）及程序人員（" & GetPrjSalesNM(strFCPHandler) & "）" & vbCrLf
                                  intI = 3
                                  If Val(cp(152)) > 0 Then
                                    intI = intI + 1
                                    'Modified by Lydia 2023/05/02 +規費金額
                                    strContent = strContent & intI & ".收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                                  End If
                                  'Add By Sindy 2022/7/13
                                  If strNP <> "" Then
                                    intI = intI + 1
                                    strContent = strContent & intI & "." & strNP & vbCrLf
                                  End If
                                  '2022/7/13 END
                  Else
                     'Modified by Lydia 2024/04/18 cp(14)改為strCP14spec
                     strTo = strCP14spec '工程師
                     strCC = PUB_GetFCPEngSup(strCP14spec) & ";backup"  '工程師之主管;backup
                     'end 2024/04/18
                     'Modified by Lydia 2022/09/27 +strSPecTitle
                     strSubject = strSPecTitle & "【已送件完成_" & GetPrjState4(lblCaseField(0), cp(10)) & IIf(Val(cp(7)) > 0, "(法限:" & Right(ChangeWStringToTDateString(cp(7)), 5) & ")", "") & "】請進行請款" & strSubject
                     'Modified by Lydia 2024/04/18 + strContEx01
                     strContent = "1." & strContEx01 & "請工程師處理請款，送件附檔可參卷宗區" & vbCrLf & _
                                  "2.請款信、金額、附件請提供給承辦（" & GetPrjSalesNM(strSales) & "）及程序人員（" & GetPrjSalesNM(strFCPHandler) & "）" & vbCrLf
                                  intI = 2
                                  If Val(cp(152)) > 0 Then
                                    intI = intI + 1
                                    'Modified by Lydia 2023/05/02 +規費金額
                                    strContent = strContent & intI & ".收據下載日：" & ChangeWStringToTDateString(cp(152)) & IIf(Val(cp(84)) > 0, "，規費金額：NTD " & cp(84), "") & vbCrLf
                                  End If
                                  'Add By Sindy 2022/7/13
                                  If strNP <> "" Then
                                    intI = intI + 1
                                    strContent = strContent & intI & "." & strNP & vbCrLf
                                  End If
                                  '2022/7/13 END
                  End If
               End If
            End If
         End If
         
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ReadAllData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_k = Nothing
End Sub

Private Function ReadAllData() As Boolean
   ReDim cp(TF_CP)
   cp(9) = m_CP09
   Call PUB_ReadCaseProgressDatabase(cp(), 1)
   
   ReDim pa(TF_PA) As String
   pa(1) = cp(1)
   pa(2) = cp(2)
   pa(3) = cp(3)
   pa(4) = cp(4)
   If cp(1) = "FCP" Then
      If PUB_ReadPatentDatabase(pa(), 國外_FC) Then
      End If
   ElseIf cp(1) = "FG" Then
      If PUB_ReadServicePracticeDatabase(pa(), 國外_FC) Then
      End If
   End If
   
   strExc(0) = "select * from caseprogress,patent where cp09='" & m_CP09 & "'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
''         strCP07 = "" & .Fields("cp07")
''         If strCP07 = "" Then strCP07 = "" & .Fields("cp142")
         strCP65 = .Fields("cp65")
         lblCaseField(0) = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
         lblCaseField(1) = "" & .Fields("pa11")
         SetNameToCombo cboCaseName, "" & .Fields("pa05"), "" & .Fields("pa06"), "" & .Fields("pa07")
         lblCaseField(2) = "" & .Fields("pa26")
         If ClsPDGetCustomer(lblCaseField(2), strExc(1)) Then
            lblAgent.Caption = strExc(1)
         End If
         lblCaseField(9) = "" & .Fields("pa09")
         If ClsPDGetNation(lblCaseField(9), strExc(1)) Then
            lblNation.Caption = strExc(1)
         End If
         lblCaseField(4) = m_CP09
         lblCaseField(5) = TransDate(.Fields("cp05"), 1)
         lblCaseField(6) = .Fields("cp10")
         If ClsPDGetCaseProperty(.Fields("cp01"), lblCaseField(6), strExc(1)) Then
            lblCaseProperty = strExc(1)
         End If
         lblCaseField(7) = .Fields("cp13")
         lblSales = GetStaffName(lblCaseField(7), True)
         strPA75 = "" & .Fields("PA75")
      End With
      
      ReadAllData = True
   Else
      MsgBox "無法讀取案件資料，請確認收文號是否正確！", vbExclamation
   End If
End Function

Private Function GetST52(strUser As String) As String
   GetST52 = ""
   strExc(0) = "SELECT ST52 FROM staff WHERE ST01='" & strUser & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields("ST52")) Then
         GetST52 = RsTemp.Fields("ST52")
      End If
   End If
End Function


