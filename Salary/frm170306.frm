VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170306 
   BorderStyle     =   1  '單線固定
   Caption         =   "各類所得轉入媒體申報套裝軟體"
   ClientHeight    =   3230
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3230
   ScaleWidth      =   6330
   Begin VB.ListBox List1 
      Height          =   1120
      Left            =   90
      TabIndex        =   4
      Top             =   1560
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "96"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "轉出(&T)"
      Height          =   405
      Left            =   2340
      TabIndex        =   1
      Top             =   30
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   3510
      TabIndex        =   0
      Top             =   30
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   930
      Width           =   6165
      _ExtentX        =   10866
      _ExtentY        =   512
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label4 
      Caption         =   "註明：轉出的檔案已改為UTF-8格式。"
      Height          =   285
      Left            =   330
      TabIndex        =   8
      Top             =   2850
      Width           =   4275
   End
   Begin VB.Label Label3 
      Caption         =   "匯出檔案將存於桌面!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2025
      TabIndex        =   7
      Top             =   570
      Width           =   3030
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      Height          =   255
      Left            =   135
      TabIndex        =   6
      Top             =   1260
      Width           =   6090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "資料年度：           年"
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   630
      Width           =   1575
   End
End
Attribute VB_Name = "frm170306"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/26 Form2.0已檢查 (待需修改:檢查長度的問題)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/1/14
Option Explicit

Dim strErrList As String


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   Screen.MousePointer = vbHourglass
   If Text1 <> "" Then
      Me.Enabled = False
      'Modify By Sindy 2023/2/4
      If Process_New = True Then
      'If Process = True Then
      '2023/2/4
         If strErrList <> "" Then
            strExc(1) = "匯出結束但下列資料有誤!!" & vbCrLf & vbCrLf & strErrList
            MsgBox strExc(1), vbExclamation
         Else
            MsgBox "成功!", vbInformation
         End If
      Else
         MsgBox "失敗!", vbCritical
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
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Modify By Sindy 2022/11/28 轉出的檔案請改為UTF-8格式
Private Function Process_New() As Boolean
   Dim strDesktop As String, strAppUnit As String
   Dim ff As Integer, strFileName As String, strContent As String
   Dim ColName As String, stValue As String, iSize As Integer
   Dim strErrDesc As String, iErr As Integer
   Dim strText As String 'Add By Sindy 2022/11/28
   Dim bolShowMsg_3 As Boolean
   
   bolShowMsg_3 = False
   List1.Clear
   strErrList = ""
   iErr = 0
   strDesktop = PUB_Getdesktop
   strExc(0) = "select * from incomedata where id14=" & (Val(Text1) + 1911) & " order by id03,id02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      List1.AddItem time & " --> 匯出資料開始...", 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      DoEvents
      strAppUnit = .Fields("id03")
      strFileName = PUB_Getdesktop & "\" & strAppUnit & "." & Format(Text1, "000") & ".U8"
      strText = "" 'Add By Sindy 2022/11/28
      '刪除舊檔
      If Dir(strFileName) <> "" Then
         Kill strFileName
      End If
'      If ff > 0 Then
'         Close #ff
'      End If
'      ff = FreeFile
'      Open strFileName For Output As ff
      Do While Not .EOF
         If strAppUnit <> .Fields("id03") Then
            'Add By Sindy 2022/11/28
            If strText <> "" Then
               Call PUB_SaveTextAsUTF8(strFileName, strText)
               strText = ""
            End If
'            If ff > 0 Then
'               Close #ff
'            End If
'            ff = FreeFile
            strAppUnit = .Fields("id03")
            strFileName = PUB_Getdesktop & "\" & strAppUnit & "." & Format(Text1, "000") & ".U8"
            '刪除舊檔
            If Dir(strFileName) <> "" Then
               Kill strFileName
            End If
'            Open strFileName For Output As ff
         End If
         strContent = ""
         For intI = 1 To 21
            ColName = "id" & Right(100 + intI, 2)
            stValue = "" & .Fields(ColName)
            
            'Added by Morgan 2011/12/20
            '共用欄位二”前面”加”所得所屬期間”X(10),年月(起)+年月(迄)
            If intI = 17 Then
               strContent = strContent & Format(Val("" & .Fields("id14")) - 1911, "000") & Format(Val("" & .Fields("id22")), "00")
               strContent = strContent & Format(Val("" & .Fields("id14")) - 1911, "000") & Format(Val("" & .Fields("id23")), "00") & "|"
            End If
            
            '年度要轉民國年
            If intI = 14 Then
               strContent = strContent & Format(Val(stValue) - 1911, "000") & "|"
            ElseIf intI = 6 And stValue = "" Then
               List1.AddItem time & " --> 【" & .Fields("id25") & "】(沒有身分證號)!!", 0
            'Add By Sindy 2023/2/9
            ElseIf intI = 11 Then
               'Add By Sindy 2024/1/25
               stValue = Trim(stValue) 'id11; 12碼; 所得人代號(或帳號)\租賃房屋稅籍編號\執行業務別\稿費必要費用別\其他所得給付項目代號
               If stValue <> "" Then
                  strContent = strContent & stValue & "|"
               Else
                  strContent = strContent & "|"
               End If
               If "" & .Fields("id05") = "54" Then '54股利
                  strContent = strContent & "1" & "|" '分配次數預帶1
                  strContent = strContent & "|"
               Else
               '2024/1/25 END
                  strContent = strContent & "||"
               End If
            ElseIf intI = 18 Then
               '檢查外國人要填”居住地國或地區代碼”
               If .Fields("id07") = "3" Then
                  bolShowMsg_3 = True
                  strContent = strContent & "ZZ|"
               Else
                  strContent = strContent & "|"
               End If
               '2023/2/9 END
            Else
               '數字欄位左邊捕0
               If .Fields(ColName).Type = adNumeric Then
                  iSize = .Fields(ColName).Precision
                  If intI = 2 Or intI = 3 Or intI = 21 Then
                     strContent = strContent & Right(String(iSize, "0") & stValue, iSize) & "|"
                  Else
                     strContent = strContent & stValue & "|"
                  End If
               '文字欄位右邊補空白 ==> Modify By Sindy 2023/2/8 不用補0欄位值後面加|
               Else
                  'Modified by Morgan 2019/1/19 O12改用CHAR長度變4Bytes
                  'iSize = .Fields(ColName).DefinedSize
                  'Modified by Morgan 2022/3/23 O12的Provider已改
                  'iSize = .Fields(ColName).DefinedSize / 4
                  iSize = .Fields(ColName).DefinedSize
                  'end 2022/3/23
                  'end 2019/1/19
                  '2014/1/14 ADD BY SONIA 共用欄位二的最後一碼改為憑單填發方式'3'(第49碼)
                  'Modify By Sindy 2023/2/9
                  'If intI = 17 Then iSize = 48
                  'strContent = strContent & PUB_StrToStr(stValue & String(iSize, Chr(32)), str(iSize)) & "|"
                  If intI = 17 Then '共用欄位二(ex:退休金額)
                     stValue = Trim(stValue)
                     If stValue <> "" Then
                        stValue = Left(stValue, 10) 'Add By Sindy 2024/1/25 取前10碼; 共用欄位二：薪資所得應為勞退自提金額數字10位及28位NULL、其他全部NULL。
                        stValue = Trim(CDbl(stValue))
                        If stValue = 0 Then stValue = ""
                     End If
                     strContent = strContent & stValue & "|"
                  Else
                  '2014/1/14 END
                     strContent = strContent & stValue & "|"
                  End If
                  '2014/1/14 ADD BY SONIA 共用欄位二的最後一碼改為憑單填發方式'3'(第49碼)
                  If intI = 17 Then strContent = strContent & "|||||3" & "|"
                  '2014/1/14 END
                  '2023/2/9 END
               End If
            End If
            
            'Added by Morgan 2011/12/20
            '共用欄位二”後面”加”是否滿183天”X(1)
            If intI = 17 Then
               '證號別為5,6,7,8的才要填Y或N(目前本所沒有)
               If .Fields("id07") = "5" Or .Fields("id07") = "7" Then
                  strContent = strContent & "N" & "|"
               ElseIf .Fields("id07") = "6" Or .Fields("id07") = "8" Then
                  List1.AddItem time & " --> 【" & .Fields("id25") & "】(無法判斷是否滿183天)!!"
               Else
                  'strContent = strContent & Chr(32) & "|"
                  strContent = strContent & "|"
               End If
            End If
         Next
         
         'Modified by Morgan 2011/12/20
         'If GetTextLength(strContent) = 200 Then
'         If GetTextLength(strContent) = 250 Then
            'Add By Sindy 2022/11/28
            strText = strText & strContent & vbCrLf
            'Print #ff, strContent
            '2022/11/28 END
'         Else
'            strErrDesc = "代號:" & .Fields("id25") & ",統編:" & .Fields("id03") & ",格式:" & .Fields("id05")
'            List1.AddItem time & " --> 【" & strErrDesc & "】(長度不符)", 0
'            strErrList = strErrList & strErrDesc & vbCrLf
'            iErr = iErr + 1
'         End If
            
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         .MoveNext
      Loop
      If strText <> "" Then
         Call PUB_SaveTextAsUTF8(strFileName, strText) 'Add By Sindy 2022/11/28
         strText = ""
      End If
'      If ff > 0 Then
'         Close #ff
'      End If
      If iErr > 0 Then
         strErrDesc = "(失敗 " & iErr & " 筆)"
      Else
         strErrDesc = ""
      End If
      List1.AddItem time & " --> 匯出資料結束,共 " & .RecordCount & " 筆" & strErrDesc, 0
      End With
   End If
   
   If bolShowMsg_3 = True Then
      MsgBox "請至匯出的媒體檔裡，補輸外籍人士的國籍！謝謝~", vbInformation
   End If
   
   Process_New = True
End Function

'Modify By Sindy 2023/2/9 改版停用
Private Function Process() As Boolean
   Dim strDesktop As String, strAppUnit As String
   Dim ff As Integer, strFileName As String, strContent As String
   Dim ColName As String, stValue As String, iSize As Integer
   Dim strErrDesc As String, iErr As Integer

   List1.Clear
   strErrList = ""
   iErr = 0
   strDesktop = PUB_Getdesktop
   strExc(0) = "select * from incomedata where id14=" & (Val(Text1) + 1911) & " order by id03,id02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      List1.AddItem time & " --> 匯出資料開始...", 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      DoEvents
      strAppUnit = .Fields("id03")
      strFileName = PUB_Getdesktop & "\" & strAppUnit & "." & Format(Text1, "000")
      If ff > 0 Then
         Close #ff
      End If
      ff = FreeFile
      Open strFileName For Output As ff
      Do While Not .EOF
         If strAppUnit <> .Fields("id03") Then
            If ff > 0 Then
               Close #ff
            End If
            ff = FreeFile
            strAppUnit = .Fields("id03")
            strFileName = PUB_Getdesktop & "\" & strAppUnit & "." & Format(Text1, "000")
            Open strFileName For Output As ff
         End If
         strContent = ""
         For intI = 1 To 21
            ColName = "id" & Right(100 + intI, 2)
            stValue = "" & .Fields(ColName)

            'Added by Morgan 2011/12/20
            '共用欄位二前面加所得所屬期限X(10),年月(起)+年月(迄)
            If intI = 17 Then
               strContent = strContent & Format(Val("" & .Fields("id14")) - 1911, "000") & Format(Val("" & .Fields("id22")), "00")
               strContent = strContent & Format(Val("" & .Fields("id14")) - 1911, "000") & Format(Val("" & .Fields("id23")), "00")
            End If

            '年度要轉民國年
            If intI = 14 Then
               strContent = strContent & Format(Val(stValue) - 1911, "000")
            ElseIf intI = 6 And stValue = "" Then
               List1.AddItem time & " --> 【" & .Fields("id25") & "】(沒有身分證號)!!", 0
            Else
               '數字欄位左邊捕0
               If .Fields(ColName).Type = adNumeric Then
                  iSize = .Fields(ColName).Precision
                  strContent = strContent & Right(String(iSize, "0") & stValue, iSize)
               '文字欄位右邊補空白
               Else
                  'Modified by Morgan 2019/1/19 O12改用CHAR長度變4Bytes
                  'iSize = .Fields(ColName).DefinedSize
                  'Modified by Morgan 2022/3/23 O12的Provider已改
                  'iSize = .Fields(ColName).DefinedSize / 4
                  iSize = .Fields(ColName).DefinedSize
                  'end 2022/3/23
                  'end 2019/1/19
                  '2014/1/14 ADD BY SONIA 共用欄位二的最後一碼改為憑單填發方式'3'(第49碼)
                  If intI = 17 Then iSize = 48
                  '2014/1/14 END
                  strContent = strContent & PUB_StrToStr(stValue & String(iSize, Chr(32)), str(iSize))
                  '2014/1/14 ADD BY SONIA 共用欄位二的最後一碼改為憑單填發方式'3'(第49碼)
                  If intI = 17 Then strContent = strContent & "3"
                  '2014/1/14 END
               End If
            End If

            'Added by Morgan 2011/12/20
            '共用欄位二後面加是否滿183天X(1)
            If intI = 17 Then
               '證號別為5,6,7,8的才要填Y或N(目前本所沒有)
               If .Fields("id07") = "5" Or .Fields("id07") = "7" Then
                  strContent = strContent & "N"
               ElseIf .Fields("id07") = "6" Or .Fields("id07") = "8" Then
                  List1.AddItem time & " --> 【" & .Fields("id25") & "】(無法判斷是否滿183天)!!"
               Else
                  strContent = strContent & Chr(32)
               End If
            End If
         Next

         'Modified by Morgan 2011/12/20
         'If GetTextLength(strContent) = 200 Then
         If GetTextLength(strContent) = 250 Then
            Print #ff, strContent
         Else
            strErrDesc = "代號:" & .Fields("id25") & ",統編:" & .Fields("id03") & ",格式:" & .Fields("id05")
            List1.AddItem time & " --> 【" & strErrDesc & "】(長度不符)", 0
            strErrList = strErrList & strErrDesc & vbCrLf
            iErr = iErr + 1
         End If

         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         .MoveNext
      Loop
      If ff > 0 Then
         Close #ff
      End If
      If iErr > 0 Then
         strErrDesc = "(失敗 " & iErr & " 筆)"
      Else
         strErrDesc = ""
      End If
      List1.AddItem time & " --> 匯出資料結束,共 " & .RecordCount & " 筆" & strErrDesc, 0
      End With
   End If
   Process = True
End Function
