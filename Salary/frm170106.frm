VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170106 
   BorderStyle     =   1  '單線固定
   Caption         =   "補充保費明細申報檔"
   ClientHeight    =   3324
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7908
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3324
   ScaleWidth      =   7908
   Begin VB.CheckBox Check1 
      Caption         =   "已申報更正"
      Height          =   180
      Left            =   216
      TabIndex        =   8
      Top             =   612
      Width           =   1560
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   6660
      TabIndex        =   2
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "轉出(&T)"
      Height          =   405
      Left            =   5445
      TabIndex        =   1
      Top             =   120
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1185
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "102"
      Top             =   187
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1308
      Left            =   135
      TabIndex        =   3
      Top             =   1584
      Width           =   7665
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   288
      Left            =   132
      TabIndex        =   4
      Top             =   888
      Width           =   7692
      _ExtentX        =   13568
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "資料年度：           年"
      Height          =   180
      Left            =   225
      TabIndex        =   7
      Top             =   232
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      Height          =   252
      Left            =   180
      TabIndex        =   6
      Top             =   1224
      Width           =   7620
   End
   Begin VB.Label Label3 
      Caption         =   "匯出檔案將存於桌面!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2115
      TabIndex        =   5
      Top             =   150
      Width           =   2940
   End
End
Attribute VB_Name = "frm170106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/26 Form2.0已檢查 (無需修改的物件)
'Create by Morgan 2013/3/11
Option Explicit
Dim strErrList As String


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   Screen.MousePointer = vbHourglass
   If Text1 = "" Then
      MsgBox "請輸入資料年度！", vbInformation
   Else
      Me.Enabled = False
      If Process = True Then
         If strErrList <> "" Then
            strExc(1) = "匯出結束但下列資料有誤!!" & vbCrLf & vbCrLf & strErrList
            MsgBox strExc(1), vbExclamation
         Else
            MsgBox "成功!", vbInformation
         End If
      End If
      Me.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text1 = strSrvDate(2) \ 10000 - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170306 = Nothing
End Sub

Private Sub Text1_GotFocus()
   CloseIme
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function Process() As Boolean
   Dim stCon As String, stCon1 As String
   Dim strDesktop As String, strAppUnit As String
   Dim ff As Integer, strFileName As String, strContent As String
   Dim ColName As String, stValue As String, iSize As Integer
   Dim strErrDesc As String, iErr As Integer
   Dim lngCount As Long, dblAmount As Double, dblFee As Double, lngSNo As Long
   
   List1.Clear
   strErrList = ""
   iErr = 0
   strDesktop = PUB_Getdesktop
   stCon = " and nhi02>=" & (Text1 + 1911) & "0101 and nhi02<=" & (Text1 + 1911) & "1231"
   
   strExc(0) = "select x1,x2,a0807,a0802,a0821 from (select nhi11 X1,'62' X2" & _
      " From nhi2nd Where nhi05>0" & stCon & _
      " union select nhi11,decode(nhi03,'50','63','9A','65','9B','65'" & _
      ",'54','66','52','67','5A','67','5B','67','5C','67','51','68',nhi03)" & _
      " From nhi2nd Where nhi06>0 And nhi05=0" & stCon & _
      "),acc080 where a0801(+)=x1 order by 1,2"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strAppUnit = ""
      With RsTemp
      Do While Not .EOF
         If strAppUnit <> .Fields("x1") Then
            If ff > 0 Then Close #ff
            ff = FreeFile
            strAppUnit = .Fields("x1")
            strFileName = PUB_Getdesktop & "\DPR" & .Fields("a0807") & strSrvDate(2) & "001"
            Open strFileName For Output As ff
         End If
         
         List1.AddItem time & " --> 匯出 " & .Fields("x1") & " 公司 " & .Fields("x2") & " 類別資料開始...", 0
         
         '申報單位資料 200
         strContent = "1" '1  資料識別碼 1-1
         strContent = strContent & .Fields("a0807") '2  申報單位統一編號 2-9
         strContent = strContent & .Fields("x2") '3  所得(收入)類別 10-11
         strContent = strContent & Text1 & "01" '4  所得給付起始年月  12-16
         strContent = strContent & Text1 & "12" '5  所得給付結束年月  17-21
         strContent = strContent & strSrvDate(2) '6  檔案製作日期   22-28
         strContent = strContent & String(8, " ") '7  總機構統一編號 29-36
         strContent = strContent & PUB_StrToStr("account@taie.com.tw", 30, True) '8  申報單位電子郵件信箱帳號   37-66
         strContent = strContent & Left(.Fields("a0802") & String(25, "　"), 25) '9  扣費義務人名稱 67-116
         strContent = strContent & String(84, " ") '10 保留欄位 117-200
         
         'Modified by Morgan 2017/1/24 非投保單位薪資所得沒有補充保費不必申報(+ nhi06>0 條件)--辜
         If GetTextLength(strContent) = 200 Then
            Print #ff, strContent
            
            stCon1 = ""
            Select Case .Fields("x2")
            Case "62"
               stCon1 = " and nhi05>0"
            Case "63"
               stCon1 = " and nhi06>0 And nhi05=0 and nhi03='50'"
            Case "65"
               stCon1 = " and nhi06>0 And nhi03 in ('9A','9B')"
            Case "66"
               stCon1 = " and nhi06>0 And nhi03='54'"
            Case "67"
               stCon1 = " and nhi06>0 And nhi03 in ('52','5A','5B','5C')"
            Case "68"
               stCon1 = " and nhi06>0 And nhi03='51'"
            Case Else
               strErrDesc = "公司別:" & .Fields("x1") & ",統編:" & .Fields("a0807") & ",類別:" & .Fields("x2") & "(類別不符)"
               List1.AddItem time & " --> 【申報單位資料:" & strErrDesc & "】", 0
               strErrList = strErrList & strErrDesc & vbCrLf
               iErr = iErr + 1
            End Select
            If stCon1 <> "" Then
               If .Fields("x2") = "62" Then
                  strExc(0) = "select id,nvl(st02,name) name,nhi02,keyno,nhi07,nhi06,nhi05,nhi13" & _
                     " from (select nhi02,nvl(st26,oi02) id,nvl(st02,oi04) name,nhi01||nhi02||nhi03||nhi04||nhi11||substr(nhi14,5) keyno,nhi07,nhi06,nhi05,nhi13" & _
                     " from nhi2nd,staff,otherincomer where nhi11='" & .Fields("x1") & "'" & stCon & stCon1 & " and st01(+)=nhi01 and oi01(+)=nhi01" & _
                     "),(select sm01 x1,st26 x2,sm02 x3 from salarymonth,staff where sm42>0 and st01(+)=sm01 and substr(sm02,1,4)=" & (Text1 + 1911) & _
                     ") x,staff where x2(+)=id and x3(+)=substr(nhi02,1,6) and st01(+)=x1 order by 1,2"
               ElseIf .Fields("x2") = "66" Then
                  'Added by Morgan 2015/1/15 +判斷個人才有補充保費(身分證號10碼)--輸入程式已修改，異常資料已刪除應該不會再有這種資料1/24
                  strExc(0) = "select nvl(st26,oi02) id,nvl(st02,oi04) name,nhi02,nhi01||nhi02||nhi03||nhi04||nhi11||substr(nhi14,5) keyno,nhi07,nhi06,nhi05,nhi13" & _
                     ",br09,br07" & _
                     " from nhi2nd, staff, otherincomer,BonusRetire where nhi11='" & .Fields("x1") & "'" & stCon & stCon1 & " and st01(+)=nhi01 and oi01(+)=nhi01" & _
                     " and br02(+)=nhi01 and br03(+)=nhi11 and br04(+)=nhi03 and br20(+)=nhi02 and length(nvl(st26,oi02))=10 order by 1,2"
               Else
                  'Added by Morgan 2017/1/23 +判斷個人才有補充保費(身分證號10碼) Ex.理律法律事務所(102007) 1050216 其他各類所得資料(平日)--輸入程式已修改，異常資料已刪除應該不會再有這種資料1/24
                  strExc(0) = "select nvl(st26,oi02) id,nvl(st02,oi04) name,nhi02,nhi01||nhi02||nhi03||nhi04||nhi11||substr(nhi14,5) keyno,nhi07,nhi06,nhi05,nhi13" & _
                     " from nhi2nd, staff, otherincomer where nhi11='" & .Fields("x1") & "'" & stCon & stCon1 & " and st01(+)=nhi01 and length(nvl(st26,oi02))=10 and oi01(+)=nhi01 order by 1,2"
               End If
               intI = 1
               Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  With adoRecordset
                  ProgressBar1.max = .RecordCount
                  ProgressBar1.Value = 0
                  Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                  DoEvents
                  
                  lngSNo = 0
                  dblAmount = 0
                  dblFee = 0
                  Do While Not .EOF
                     lngSNo = lngSNo + 1
                     dblAmount = dblAmount + .Fields("nhi07")
                     dblFee = dblFee + .Fields("nhi06")
                     '扣費明細資料 200
                     strContent = "2" '1  資料識別碼 1-1(C1)
                     strContent = strContent & RsTemp.Fields("a0807") '2  申報單位統一編號 2-9(C8)
                     strContent = strContent & RsTemp.Fields("x2") '3  所得(收入)類別 10-11(C2)
                     strContent = strContent & Format(lngSNo, String(9, "0")) '4 流水序號 12-20(N9)
                     'Modified by Morgan 2018/1/25 +R
                     If Check1.Value = vbChecked Then
                        strContent = strContent & "R"
                     Else
                        strContent = strContent & "I" '5 資料處理方式 21-21(C1)
                     End If
                     'end 2018/1/25
                     strContent = strContent & (.Fields("nhi02") - 19110000) '6 所得給付日期 22-28(C7)
                     strContent = strContent & .Fields("id") '7 所得人身分證號 29-38(C10)
                     strContent = strContent & Left(.Fields("keyno") & String(30, " "), 30) '8 申報編號 39-68(C30)
                     strContent = strContent & Format(.Fields("nhi07"), String(14, "0")) '9 所得(收入)給付金額 69-82(N14)
                     strContent = strContent & Format(.Fields("nhi06"), String(10, "0")) '10 扣繳補充保險費金額 83-92(N10)
                     '11 共用欄位區 93-132(C40)
                     '獎金(62)
                     If RsTemp.Fields("x2") = "62" Then
                        strContent = strContent & RsTemp.Fields("a0821") '投保單位代號 93-101(C9)
                        strContent = strContent & Format(.Fields("nhi05"), String(6, "0")) '扣費當月投保金額 102-107(N6)
                        strContent = strContent & Format(.Fields("nhi13"), String(10, "0")) '同年度累計獎金金額 108-117(N10)
                        strContent = strContent & String(15, " ") '保留欄位 118-132(C15)
                     '股利(66)
                     ElseIf RsTemp.Fields("x2") = "66" Then
                        'Modified by Morgan 2019/1/29 已取消可扣抵稅額，會是Null
                        strContent = strContent & Format(Val("" & .Fields("br09")), String(10, "0")) '扣取時可扣抵稅額 93-102(N10)
                        strContent = strContent & Format(Val("" & .Fields("br09")), String(10, "0")) '年度確定可扣抵稅額 103-112(N10)
                        strContent = strContent & String(10, "0") '已列入投保金額計算保險費之股利金額 113-122(N10)
                        strContent = strContent & (.Fields("br07") - 19110000) '除權(息)基準日期 123-129(C7)
                        strContent = strContent & "2" '股利註記 130-130(C1)
                        strContent = strContent & String(2, " ") '保留欄位 131-132(C2)
                     Else
                        strContent = strContent & String(40, " ")
                     End If
                     strContent = strContent & String(1, " ") '12 信託註記 133-133(C1)
                     'Modified by Morgan 2025/2/4 若有Unicode要先轉Big5，否則存檔後才變?號上傳時發生長度錯誤
                     'strContent = strContent & Left(.Fields("name") & String(25, "　"), 25) '13 所得人姓名 134-183(C50)
                     strContent = strContent & Left(fnUniToBig5(.Fields("name")) & String(25, "　"), 25) '13 所得人姓名 134-183(C50)
                     'end 2025/2/4
                     strContent = strContent & String(17, " ") '14 保留欄位 184 - 200
                     If GetTextLength(strContent) = 200 Then
                        Print #ff, strContent
                     Else
                        strErrDesc = "公司別:" & RsTemp.Fields("x1") & ",統編:" & RsTemp.Fields("a0807") & ",類別:" & RsTemp.Fields("x2") & ",所得人:" & .Fields("name") & .Fields("id") & "(長度不符)"
                        List1.AddItem time & " --> 【扣費明細資料:" & strErrDesc & "】", 0
                        strErrList = strErrList & strErrDesc & vbCrLf
                        iErr = iErr + 1
                     End If
                     
                     ProgressBar1.Value = ProgressBar1.Value + 1
                     Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                     DoEvents
         
                     .MoveNext
                  Loop
                  End With
               End If
            End If
         Else
            strErrDesc = "公司別:" & .Fields("x1") & ",統編:" & .Fields("a0807") & ",類別:" & .Fields("x2") & "(長度不符)"
            List1.AddItem time & " --> 【申報單位資料:" & strErrDesc & "】", 0
            strErrList = strErrList & strErrDesc & vbCrLf
            iErr = iErr + 1
         End If
         
         '申報單位資料 200
         strContent = "3" '1  資料識別碼 1-1
         strContent = strContent & .Fields("a0807") '2  申報單位統一編號 2-9
         strContent = strContent & .Fields("x2") '3  所得(收入)類別 10-11
         strContent = strContent & Right(String(9, "0") & lngSNo, 9) '4  申報總筆數  12-20
         strContent = strContent & Right(String(20, "0") & dblAmount, 20)  '5  所得(收入)給付總額   21-40 20
         strContent = strContent & Right(String(16, "0") & dblFee, 16) '6  扣繳補充保險費總額   41-56
         strContent = strContent & PUB_StrToStr("0225061023#542", 15, True)  '7  聯絡電話
         strContent = strContent & Left("吳婉莘" & String(25, "　"), 25) '8  聯絡人姓名  72-121  'modify by sonia 2022/10/25 吳婧瑄->吳婉莘
         strContent = strContent & String(79, " ")   '9  保留欄位 122-200
         
         If GetTextLength(strContent) = 200 Then
            Print #ff, strContent
         Else
            strErrDesc = "公司別:" & .Fields("x1") & ",統編:" & .Fields("a0807") & ",類別:" & .Fields("x2") & "(長度不符)"
            List1.AddItem time & " --> 【扣費明細總計:" & strErrDesc & "】", 0
            strErrList = strErrList & strErrDesc & vbCrLf
            iErr = iErr + 1
         End If
         DoEvents
         .MoveNext
      Loop
      If ff > 0 Then Close #ff
      If iErr > 0 Then
         strErrDesc = "(失敗 " & iErr & " 筆)"
      Else
         strErrDesc = ""
      End If
      List1.AddItem time & " --> 匯出資料結束,共 " & .RecordCount & " 筆" & strErrDesc, 0
      End With
      Process = True
   Else
      MsgBox "無資料可匯出!", vbExclamation
   End If
   
End Function

'Added by Morgan 2025/2/4
Private Function fnUniToBig5(ByVal pText As String, Optional pRepChar As String = "　") As String
   Dim strTemp As String, strTemp2 As String, strChar As String, ii As Integer
   strTemp = StrConv(StrConv(pText, vbFromUnicode), vbUnicode)
   If InStr(strTemp, "?") > 0 Then
      strTemp2 = strTemp
      strTemp = ""
      For ii = 1 To Len(strTemp2)
         strChar = Mid(strTemp2, ii, 1)
         If strChar = "?" And Mid(pText, ii, 1) <> strChar Then
            strTemp = strTemp & pRepChar
         Else
            strTemp = strTemp & strChar
         End If
      Next
   End If
   fnUniToBig5 = strTemp
End Function
