VERSION 5.00
Begin VB.Form frm040335 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限通知檢核及報表"
   ClientHeight    =   3708
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5988
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   5988
   Begin VB.TextBox TxtCnt 
      BackColor       =   &H8000000F&
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   660
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   25
      Top             =   2580
      Width           =   450
   End
   Begin VB.ListBox List2 
      BackColor       =   &H8000000F&
      Height          =   1128
      ItemData        =   "frm040335.frx":0000
      Left            =   3960
      List            =   "frm040335.frx":0002
      TabIndex        =   23
      Top             =   1980
      Width           =   1935
   End
   Begin VB.Frame FrameEmp 
      Height          =   280
      Left            =   180
      TabIndex        =   20
      Top             =   240
      Width           =   2530
      Begin VB.ComboBox Combo3 
         Height          =   260
         Left            =   990
         TabIndex        =   21
         Text            =   "Combo3"
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "程序人員："
         Height          =   180
         Index           =   4
         Left            =   0
         TabIndex        =   22
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   260
      Left            =   1170
      Style           =   2  '單純下拉式
      TabIndex        =   19
      Top             =   600
      Width           =   1395
   End
   Begin VB.CheckBox Check1 
      Caption         =   "副本"
      Height          =   375
      Left            =   3990
      TabIndex        =   17
      Top             =   1560
      Width           =   720
   End
   Begin VB.ComboBox Combo1 
      Height          =   260
      ItemData        =   "frm040335.frx":0004
      Left            =   1170
      List            =   "frm040335.frx":0006
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   1260
      Width           =   2760
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "清除(&U)"
      Enabled         =   0   'False
      Height          =   400
      Index           =   4
      Left            =   3150
      TabIndex        =   15
      Top             =   2460
      Width           =   756
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   260
      Left            =   1170
      TabIndex        =   13
      Top             =   3330
      Width           =   3630
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "報表(&R)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   3150
      TabIndex        =   8
      Top             =   120
      Width           =   756
   End
   Begin VB.ListBox List1 
      Height          =   1128
      ItemData        =   "frm040335.frx":0008
      Left            =   1170
      List            =   "frm040335.frx":000A
      TabIndex        =   12
      Top             =   1950
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "新增(&A)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3150
      TabIndex        =   9
      Top             =   1620
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "檢核(&C)"
      Enabled         =   0   'False
      Height          =   400
      Index           =   1
      Left            =   3150
      TabIndex        =   11
      Top             =   2040
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   930
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2715
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1650
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2475
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1650
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1635
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1650
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "P"
      Top             =   1650
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(待確認案件)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   6
      Left            =   4800
      TabIndex        =   26
      Top             =   1770
      Width           =   1150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "件數："
      Height          =   180
      Index           =   5
      Left            =   60
      TabIndex        =   24
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "報表所別："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   18
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "通知性質："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   3390
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "通知日期："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   980
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   1700
      Width           =   900
   End
End
Attribute VB_Name = "frm040335"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2015/9/2
Option Explicit

Dim strKey1 As String, StrKey2 As String
'列印用
Dim strPrinter As String
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Public m_ProState As String
Dim m_strCP01 As String 'Add By Sindy 2025/4/17


Private Function CheckCase() As Boolean
   Dim stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String, stCon As String
   Dim ii As Integer
   Dim strConSql As String 'Add By Sindy 2025/4/21
   Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2025/4/21
   
   If Text1(2) = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Text1(2).SetFocus
      Exit Function
   End If
   
   stCP01 = Text1(1)
   stCP02 = Right("00000" & Text1(2), 6)
   stCP03 = Right("0" & Text1(3), 1)
   stCP04 = Right("00" & Text1(4), 2)
   
   If List1.ListCount > 0 Then
      For ii = 0 To List1.ListCount - 1
         If List1.List(ii) = stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04 Then
            MsgBox "本所案號重複！", vbExclamation
            Text1(2).SetFocus
            Text1_GotFocus 2
            Exit Function
         End If
      Next
   End If
   
   If Text1(0) <> "" Then
      'Modified by Morgan 2019/2/20 CFP案要確認報價,發文日可能會不同,改以收文日判斷
      If m_ProState = "CFP" Then
         stCon = stCon & " and c1.cp05>=" & DBDATE(Text1(0)) & " and c1.cp27>=c1.cp05"
      Else
         stCon = stCon & " and c1.cp27=" & DBDATE(Text1(0))
      End If
   End If
   
   If Combo1.ItemData(Combo1.ListIndex) <> 0 Then
      stCon = stCon & " and c1.cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'"
   End If
   
   'Add By Sindy 2025/4/16
   If FrameEmp.Visible = True Then
   '2025/4/16 END
      'Added by Morgan 2020/4/16
      '程序人員(承辦人)
      If Combo3.Visible = True And Trim(Combo3.Text) <> "" Then
         stCon = stCon & " and c1.cp14='" & Left(Combo3, 5) & "'"
      End If
      'end 2020/4/16
   End If
   
   'Add By Sindy 2025/4/17
   If m_ProState = "T" And m_strCP01 <> "" Then
      stCon = stCon & " and c1.cp01='" & m_strCP01 & "'"
   End If
   Screen.MousePointer = vbHourglass
   '2025/4/17 END
   
   'Modified by Morgan 2015/10/8 +判斷大對台案件
   'Modified by Morgan 2016/7/7 年費逾期通知只限台灣案--玲玲
   'Modified by Morgan 2016/11/16 +副本語法
   'Modified by Morgan 2019/10/21 +本所信函1999
   'Modify By Sindy 2025/4/21 加入strConSQL
   If Check1.Value = vbChecked Then
      strConSql = "select decode(c1.cp01,'CFP',c1.cp05,c2.cp27)-19110000 dt,lp33 Key1,lp34 Key2,c2.cp09,c1.cp10,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04" & _
         " from caseprogress c1,caseprogress c2,letterprogress" & _
         " where c1.cp10 in ('1913','1605','1999')" & stCon & _
         " and c2.cp43(+)=c1.cp09 and c2.cp10='990' and lp01(+)=c2.cp09 and lp10='Y' and lp32='Y' and lp15='N'"
      
      strExc(0) = strConSql & " and c1.cp01='" & stCP01 & "' and c1.cp02='" & stCP02 & "' and c1.cp03='" & stCP03 & "' and c1.cp04='" & stCP04 & "'" & _
         " order by c2.cp27 desc"
   Else
      'Modify By Sindy 2025/4/16 +內商
      If m_ProState = "T" Then
         strConSql = "select c1.cp27-19110000 dt,decode(lp31,'Y',tm44,tm23) Key1,decode(lp31,'Y','',nvl(tm123,cu127)) Key2,c1.cp09,c1.cp10,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04" & _
            " from caseprogress c1,letterprogress,trademark,customer" & _
            " where c1.cp10 in ('1725','1717')" & stCon & _
            " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N'" & _
            " and tm01(+)=c1.cp01 and tm02(+)=c1.cp02 and tm03(+)=c1.cp03 and tm04(+)=c1.cp04 and tm01 is not null" & _
            " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) TMReplaceSQL"
         strConSql = strConSql & " union " & _
            "select c1.cp27-19110000 dt,decode(lp31,'Y',sp26,sp08) Key1,decode(lp31,'Y','',nvl(sp78,cu127)) Key2,c1.cp09,c1.cp10,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04" & _
            " from caseprogress c1,letterprogress,servicepractice,customer" & _
            " where c1.cp10 in ('1725','1717')" & stCon & _
            " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N'" & _
            " and sp01(+)=c1.cp01 and sp02(+)=c1.cp02 and sp03(+)=c1.cp03 and sp04(+)=c1.cp04 and sp01 is not null" & _
            " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) SPReplaceSQL"
         
         strExc(0) = Replace(strConSql, "TMReplaceSQL", " and c1.cp01='" & stCP01 & "' and c1.cp02='" & stCP02 & "' and c1.cp03='" & stCP03 & "' and c1.cp04='" & stCP04 & "'")
         strExc(0) = Replace(strExc(0), "SPReplaceSQL", " and c1.cp01='" & stCP01 & "' and c1.cp02='" & stCP02 & "' and c1.cp03='" & stCP03 & "' and c1.cp04='" & stCP04 & "'") & _
              " order by dt desc"
      Else
      '2025/4/16 END
         'Modified by Morgan 2019/2/20 CFP案要確認報價,發文日可能會不同,改以收文日判斷
         strConSql = "select decode(c1.cp01,'CFP',c1.cp05,c1.cp27)-19110000 dt,decode(lp31,'Y',pa75,pa26) Key1,decode(lp31,'Y','',nvl(pa149,cu127)) Key2,c1.cp09,c1.cp10,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04" & _
            " from caseprogress c1,letterprogress,patent,customer" & _
            " where c1.cp10 in ('1913','1605','1999')" & stCon & _
            " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N' and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04 and not (c1.cp10='1605' and pa09<>'000')" & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
            
         strExc(0) = strConSql & " and c1.cp01='" & stCP01 & "' and c1.cp02='" & stCP02 & "' and c1.cp03='" & stCP03 & "' and c1.cp04='" & stCP04 & "'" & _
            " order by c1.cp27 desc"
      End If
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Adde by Morgan 2019/1/8
      '若一案多期限時再確認筆數 Ex:CFP-015828
      If RsTemp.RecordCount > 1 Then
         Screen.MousePointer = vbDefault 'Add By Sindy 2025/4/22
         If MsgBox(stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04 & "案有 " & RsTemp.RecordCount & " 筆期限要通知，是否確定？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
         Screen.MousePointer = vbHourglass 'Add By Sindy 2025/4/22
      End If
      'end 2019/1/8
      If List1.ListCount = 0 Then
         If Text1(0) = "" Then
            Text1(0) = RsTemp("dt")
         End If
         Text1(0).Enabled = False
         For ii = 0 To Combo1.ListCount - 1
            If Combo1.ItemData(ii) = RsTemp("cp10") Then
               Combo1.ListIndex = ii
               Exit For
            End If
         Next
         Combo1.Enabled = False
         
         strKey1 = RsTemp("Key1")
         StrKey2 = "" & RsTemp("Key2")
         List1.AddItem stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04, 0
         List1.ItemData(0) = PUB_DocNo2Num(RsTemp("cp09"))  '收文號
         'Added by Morgan 2019/1/8
         If RsTemp.RecordCount > 1 Then
            RsTemp.MoveNext
            Do While Not RsTemp.EOF
               List1.AddItem stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04, 0
               List1.ItemData(0) = PUB_DocNo2Num(RsTemp("cp09"))  '收文號
               RsTemp.MoveNext
            Loop
         End If
         'end 2019/1/8
         
         'Add By Sindy 2025/4/21
         '將其他案號寫入待確認案件ListBox裡
         If InStr(strConSql, "TMReplaceSQL") > 0 Or InStr(strConSql, "SPReplaceSQL") > 0 Then
            If InStr(strConSql, "TMReplaceSQL") > 0 Then
               strExc(10) = " and substr(decode(lp31,'Y',tm44,tm23),1,8)='" & Mid(strKey1, 1, 8) & "'"
               If StrKey2 = "" Then
                  strExc(10) = strExc(10) & " and decode(lp31,'Y','',nvl(tm123,cu127)) is null"
               Else
                  strExc(10) = strExc(10) & " and decode(lp31,'Y','',nvl(tm123,cu127))='" & StrKey2 & "'"
               End If
               strConSql = Replace(strConSql, "TMReplaceSQL", strExc(10))
            End If
            If InStr(strConSql, "SPReplaceSQL") > 0 Then
               strExc(10) = " and substr(decode(lp31,'Y',sp26,sp08),1,8)='" & Mid(strKey1, 1, 8) & "'"
               If StrKey2 = "" Then
                  strExc(10) = strExc(10) & " and decode(lp31,'Y','',nvl(sp78,cu127)) is null"
               Else
                  strExc(10) = strExc(10) & " and decode(lp31,'Y','',nvl(sp78,cu127))='" & StrKey2 & "'"
               End If
               strConSql = Replace(strConSql, "SPReplaceSQL", strExc(10))
            End If
         Else
            If Check1.Value = vbChecked Then
               strConSql = strConSql & _
                  " and substr(lp33,1,8)='" & Mid(strKey1, 1, 8) & "'"
               If StrKey2 = "" Then
                  strConSql = strConSql & " and lp34 is null"
               Else
                  strConSql = strConSql & " and lp34='" & StrKey2 & "'"
               End If
            Else
               strConSql = strConSql & _
                  " and substr(decode(lp31,'Y',pa75,pa26),1,8)='" & Mid(strKey1, 1, 8) & "'"
               If StrKey2 = "" Then
                  strConSql = strConSql & " and decode(lp31,'Y','',nvl(pa149,cu127)) is null"
               Else
                  strConSql = strConSql & " and decode(lp31,'Y','',nvl(pa149,cu127))='" & StrKey2 & "'"
               End If
            End If
         End If
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strConSql)
         If intI = 1 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               If rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & "-" & rsTmp.Fields("cp03") & "-" & rsTmp.Fields("cp04") <> stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04 Then
                  List2.AddItem rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & "-" & rsTmp.Fields("cp03") & "-" & rsTmp.Fields("cp04"), 0
               End If
               rsTmp.MoveNext
            Loop
         End If
         '2025/4/21 END
         
         Text1(2) = ""
         Text1(3) = ""
         Text1(4) = ""
         'Add By Sindy 2025/4/21
         TxtCnt.Text = List1.ListCount '記錄筆數
         Text1(2).SetFocus
         '2025/4/21 END
         CheckCase = True
      Else
         'Modified by Morgan 2019/2/20
         If Text1(0) <> RsTemp(0) And m_ProState <> "CFP" Then
            Screen.MousePointer = vbDefault 'Add By Sindy 2025/4/22
            MsgBox "通知日期不同！", vbExclamation
            Text1(2).SetFocus
            Text1_GotFocus 2
            Exit Function
         'Modified by Morgan 2016/4/14
         '只需抓前8碼判斷
         'ElseIf strKey1 <> RsTemp(1) Then
         ElseIf Left(strKey1, 8) <> Left(RsTemp(1), 8) Then
            'Modified by Morgan 2016/11/16
            'MsgBox "申請人不同！", vbExclamation
            Screen.MousePointer = vbDefault 'Add By Sindy 2025/4/22
            MsgBox "收受人不同！", vbExclamation
            Text1(2).SetFocus
            Text1_GotFocus 2
            Exit Function
         ElseIf StrKey2 <> "" & RsTemp(2) Then
            Screen.MousePointer = vbDefault 'Add By Sindy 2025/4/22
            MsgBox "接洽人不同！", vbExclamation
            Text1(2).SetFocus
            Text1_GotFocus 2
            Exit Function
         ElseIf Combo1.ItemData(Combo1.ListIndex) <> RsTemp("cp10") Then
            Screen.MousePointer = vbDefault 'Add By Sindy 2025/4/22
            MsgBox "期限通知性質不同！", vbExclamation
            Text1(2).SetFocus
            Text1_GotFocus 2
            Exit Function
         Else
            List1.AddItem stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04, 0
            List1.ItemData(0) = PUB_DocNo2Num(RsTemp("cp09"))  '收文號
            'Added by Morgan 2019/1/8
            If RsTemp.RecordCount > 1 Then
               RsTemp.MoveNext
               Do While Not RsTemp.EOF
                  List1.AddItem stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04, 0
                  List1.ItemData(0) = PUB_DocNo2Num(RsTemp("cp09"))  '收文號
                  RsTemp.MoveNext
               Loop
            End If
            'end 2019/1/8
            
            'Add By Sindy 2025/4/21
            '檢查待確認案件有資料時,要移除
            If List2.ListCount > 0 Then
               For ii = 0 To List2.ListCount - 1
                  If List2.List(ii) = stCP01 & "-" & stCP02 & "-" & stCP03 & "-" & stCP04 Then
                     List2.RemoveItem ii
                     Exit For
                  End If
               Next
            End If
            '2025/4/21 END
            
            Text1(2) = ""
            Text1(3) = ""
            Text1(4) = ""
            'Add By Sindy 2025/4/21
            TxtCnt.Text = List1.ListCount '記錄筆數
            Text1(2).SetFocus
            '2025/4/21 END
            CheckCase = True
         End If
      End If
   Else
      MsgBox "該案號沒有通知信函！", vbInformation
      Text1(2).SetFocus 'Add By Sindy 2025/4/22
      Text1_GotFocus 2
   End If
   Screen.MousePointer = vbDefault 'Add By Sindy 2025/4/22
   If CheckCase = True Then Check1.Value = vbUnchecked
End Function

Private Function CheckList() As Boolean
   Dim ii As Integer, stCon As String, stCon1 As String, bFound As Boolean, strMissList As String
   Dim stConS As String 'Add By Sindy 2025/4/16
   
   If StrKey2 <> "" Then
      'Modify By Sindy 2025/4/16 +內商
      If m_ProState = "T" Then
         stCon = " and nvl(tm123,cu127)='" & StrKey2 & "'"
         stConS = " and nvl(sp78,cu127)='" & StrKey2 & "'"
      Else
      '2025/4/16 END
         stCon = " and nvl(pa149,cu127)='" & StrKey2 & "'"
      End If
      stCon1 = " and lp34='" & StrKey2 & "'"
      
   'Added by Morgan 2021/3/3
   Else
      'Modify By Sindy 2025/4/16 +內商
      If m_ProState = "T" Then
         stCon = " and nvl(tm123,cu127) is null"
         stConS = " and nvl(sp78,cu127) is null"
      Else
      '2025/4/16 END
         stCon = " and nvl(pa149,cu127) is null"
      End If
      stCon1 = " and lp34 is null"
   'end 2021/3/3
   End If
   
   'Add By Sindy 2025/4/17
   If m_ProState = "T" And m_strCP01 <> "" Then
      stCon = stCon & " and cp01='" & m_strCP01 & "'"
      stConS = stConS & " and cp01='" & m_strCP01 & "'"
   End If
   '2025/4/17 END
   
   'Modified by Morgan 2016/4/14
   '客戶編號只需抓前8碼判斷
   'Modified by Morgan 2016/11/8 +lp32='Y'
   'Modified by Morgan 2019/2/20 CFP案要確認報價,發文日可能會不同,改以收文日判斷
   'Modify By Sindy 2025/4/16 +內商
   If m_ProState = "T" Then
      strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo,cp09,cp10" & _
         " from caseprogress,letterprogress,trademark,customer" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05" & _
         " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='" & Text1(1) & "'" & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N'" & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9)" & _
         " and substr(decode(lp31,'Y',tm44,tm23),1,8)='" & Left(strKey1, 8) & "'" & stCon
      strExc(0) = strExc(0) & " union " & _
         "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo,cp09,cp10" & _
         " from caseprogress,letterprogress,servicepractice,customer" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05" & _
         " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='" & Text1(1) & "'" & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N'" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9)" & _
         " and substr(decode(lp31,'Y',sp26,sp08),1,8)='" & Left(strKey1, 8) & "'" & stConS
   Else
   '2025/4/16 END
      strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo,cp09,cp10 from caseprogress,letterprogress,patent,customer" & _
         " where " & IIf(m_ProState = "CFP", "cp05>=", "cp27=") & DBDATE(Text1(0)) & " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='" & Text1(1) & "'" & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and substr(decode(lp31,'Y',pa75,pa26),1,8)='" & Left(strKey1, 8) & "'" & stCon
   End If
   'Added by Morgan 2016/11/16
   strExc(0) = strExc(0) & " union select c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) CaseNo,c1.cp09,c1.cp10 from caseprogress c1,caseprogress c2,letterprogress" & _
      " where " & IIf(m_ProState = "CFP", "c1.cp05>=", "c1.cp27=") & DBDATE(Text1(0)) & " and c1.cp27>=c1.cp05 and c1.cp01='" & Text1(1) & "' and c1.cp10='990' and c2.cp09(+)=c1.cp43 and c2.cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
      " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N' and substr(lp33,1,8)='" & Left(strKey1, 8) & "'" & stCon1
   'end 2016/11/16
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         bFound = False
         For ii = 0 To List1.ListCount - 1
            If List1.ItemData(ii) = PUB_DocNo2Num(RsTemp("cp09")) Then
               bFound = True
               Exit For
            End If
         Next
         If bFound = False Then
            strMissList = strMissList & vbCrLf & .Fields("CaseNo") & IIf(.Fields("cp10") = "990", "(副本)", "")
         End If
         .MoveNext
      Loop
      End With
      If strMissList = "" Then
         If UpdateList = True Then
            CheckList = True
         End If
      Else
         MsgBox "尚缺下列案號！" & vbCrLf & strMissList, vbExclamation, "檢核失敗"
         Text1(2).SetFocus
         Text1_GotFocus 2
      End If
   End If
End Function

'更新檢核日期
Private Function UpdateList() As Boolean
   Dim ii As Integer, strNo As String
   
On Error GoTo ErrHnd
   
   strNo = "'" & PUB_Num2DocNo(List1.ItemData(0)) & "'"
   For ii = 1 To List1.ListCount - 1
      strNo = strNo & ",'" & PUB_Num2DocNo(List1.ItemData(ii)) & "'"
   Next
   strSql = "update letterprogress set lp27='" & strUserNum & "',lp28=sysdate where lp01 in (" & strNo & ")"
   cnnConnection.Execute strSql, intI
   UpdateList = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbExclamation
End Function

Private Sub setForm(Index As Integer)
   Select Case Index
   Case 0
      If List1.ListCount > 0 Then
         Text1(0).Enabled = False
         Combo1.Enabled = False
         cmdOK(4).Enabled = True
      'Else
         cmdOK(1).Enabled = True
      End If
      
   Case 1, 4
      List1.Clear
      List2.Clear: TxtCnt.Text = "" 'Add By Sindy 2025/4/21
      Text1(0).Enabled = True
      Combo1.Enabled = True
      cmdOK(1).Enabled = False
      cmdOK(4).Enabled = False
      strKey1 = ""
      StrKey2 = ""
   End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
   'Add By Sindy 2025/4/17
   m_strCP01 = ""
   If InStr(Trim(Combo1.Text), "-") > 0 Then
      m_strCP01 = SystemNumber(Trim(Combo1.Text), 1)
   End If
   '2025/4/17 END
   Select Case Index
   Case 0 '新增
      If CheckCase = True Then
         setForm Index
      End If
      
   Case 1 '檢核
      If CheckList = True Then
         setForm Index
      End If
      
   Case 2 '報表
      PrintReport
      
   Case 3 '結束
      Unload Me
      
   Case 4 '清除
      setForm Index
   End Select
End Sub

'Add By Sindy 2025/4/17
Private Sub Combo1_Click()
   If InStr(Trim(Combo1.Text), "-") > 0 Then
      Text1(1) = SystemNumber(Trim(Combo1.Text), 1)
   ElseIf m_ProState = "T" Then
      Text1(1) = "T"
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   
   'Add By Sindy 2025/4/16
   Combo1.Clear
   FrameEmp.Visible = True
   If m_ProState = "T" Then
      Text1(1).Text = "T"
      Text1(1).Enabled = True
      Check1.Visible = False '副本隱藏
      FrameEmp.Visible = False '不分程序人員
      Combo1.AddItem "TM-通知期限", 0
      Combo1.ItemData(0) = 1725
      Combo1.AddItem "TD-通知期限", 0
      Combo1.ItemData(0) = 1725
      Combo1.AddItem "TF-通知期限", 0
      Combo1.ItemData(0) = 1725
      Combo1.AddItem "T-通知期限", 0
      Combo1.ItemData(0) = 1725
      Combo1.AddItem "智慧局通知延展", 0
      Combo1.ItemData(0) = 1717
'      Combo1.AddItem "通知繳納註冊費", 0 'Add By Sindy 2025/4/23 件數不多,取消為大宗
'      Combo1.ItemData(0) = 1720
      
      Combo1.AddItem "", 0
      Combo1.ItemData(0) = 0
   Else
   '2025/4/16 END
      'Modified by Morgan 2018/10/25 +CFP
      If m_ProState = "CFP" Then
         Text1(1) = m_ProState
         Label1(1) = "輸入日期>="
         Combo1.AddItem "通知期限", 0
         Combo1.ItemData(0) = 1913
         Combo1.Enabled = False
      Else
         Combo1.AddItem "通知年費逾期", 0
         Combo1.ItemData(0) = 1605
         Combo1.AddItem "通知期限", 0
         Combo1.ItemData(0) = 1913
         'Modified by Morgan 2019/10/21 +本所信函1999
         Combo1.AddItem "本所信函", 0
         Combo1.ItemData(0) = 1999
         
         Combo1.AddItem "", 0
         Combo1.ItemData(0) = 0
      End If
      'end 2018/10/25
   End If
   Combo1.ListIndex = 0
   
   'Added by Lydia 2019/12/16
   Combo2.Clear
   Combo2.AddItem "0. 全所"
   Combo2.AddItem "1. 北所"
   Combo2.AddItem "2. 非北所"
   Combo2.ListIndex = 0
      
   'Added by Morgan 2020/4/16
   If m_ProState = "CFP" Then
      Combo2.Enabled = False
      Label1(4).Visible = True
      Combo3.Visible = True
      Call SetPatentP12Combo(Combo3, m_ProState, Label1(4))
   
   'Added by Morgan 2025/1/16
   ElseIf strSrvDate(1) >= P業務區劃分啟用日 Then
      Label1(4).Visible = True
      Combo3.Visible = True
      Call SetPatentP12Combo(Combo3, "P", Label1(4))
   'end 2025/1/16
   Else
      Label1(4).Visible = False
      Combo3.Visible = False
   End If
   'end 2020/4/16
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040335 = Nothing
End Sub

'Add By Sindy 2025/4/22
Private Sub List2_Click()
Dim ii As Integer

   If List2.ListCount > 0 Then
      For ii = 0 To List2.ListCount - 1
         If List2.Selected(ii) = True Then
            If Text1(1).Enabled = True Then Text1(1) = SystemNumber(Trim(List2.List(ii)), 1)
            Text1(2) = SystemNumber(Trim(List2.List(ii)), 2)
            Text1(3) = SystemNumber(Trim(List2.List(ii)), 3)
            Text1(4) = SystemNumber(Trim(List2.List(ii)), 4)
            Exit For
         End If
      Next ii
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   CloseIme
   TextInverse Text1(Index)
End Sub

'Modified by Morgan 2016/11/8 所有語法+lp32='Y' 條件
'Modified by Morgan 2017/10/13 所有語法+lp15='N'條件,因為台灣大陸會同一天催期限(Ex:106/10/12)
Private Sub PrintReport()
   Dim iNoPaperCount As Integer 'Added by Morgan 2019/3/8
   Dim stCon As String
   
   If Text1(0) = "" Then
      MsgBox "請輸入通知日期！", vbExclamation
      Text1(0).SetFocus
      Text1_GotFocus 0
      Exit Sub
   ElseIf Not ChkDate(Text1(0)) Then
      Exit Sub
   End If
   
   If Combo1.ItemData(Combo1.ListIndex) = 0 Then
      MsgBox "請選擇通知性質", vbInformation
      Combo1.SetFocus
      Exit Sub
   End If
   
   'Modified by Morgan 2020/1/3 從下面印報表前移上來(檢核也要控制所別)
   'Added by Lydia 2019/12/16
   strExc(1) = ""
   If Left(Combo2.Text, 1) = "1" Then '北所
        strExc(1) = " and st06='1' "
   ElseIf Left(Combo2.Text, 1) = "2" Then '非北所
        strExc(1) = " and st06<>'1' "
   End If
   'end 2020/1/3
   
   'Add By Sindy 2025/4/16
   If FrameEmp.Visible = True Then
   '2025/4/16 END
      'Added by Morgan 2020/4/16
      '程序人員(承辦人)
      If Combo3.Visible = True And Trim(Combo3.Text) <> "" Then
         stCon = stCon & " and c1.cp14='" & Left(Combo3, 5) & "'"
      End If
      'end 2020/4/16
   End If
   
   '剔除FMP
   'Modified by Morgan 2015/10/8 +判斷大對台案件
   'Modified by Morgan 2016/4/14
   '客戶編號只需抓前8碼判斷
   'Modified by Morgan 2016/7/7 年費逾期通知只限台灣案--玲玲
   'Modified by Morgan 2016/11/16 +副本語法,已加lp32判斷大批發文不必再判斷國家
   'Modified by Morgan 2019/2/20 CFP案要確認報價,發文日可能會不同,改以收文日判斷
   'Modified by Morgan 2020/4/16 +stCon 程序人員條件
   strExc(0) = "select max(LP01) from ("
   'Modify By Sindy 2025/4/16 +內商
   If m_ProState = "T" Then
      strExc(0) = strExc(0) & " select lp01, decode(lp31,'Y',substr(tm44,1,8),substr(tm23,1,8)||nvl(tm123,cu127)) key" & _
         " From caseprogress c1, trademark, letterprogress, CUSTOMER" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N'" & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9)"
      strExc(0) = strExc(0) & " union select lp01, decode(lp31,'Y',substr(sp26,1,8),substr(sp08,1,8)||nvl(sp78,cu127)) key" & _
         " From caseprogress c1, servicepractice, letterprogress, CUSTOMER" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N'" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9)"
   Else
   '2025/4/16 END
      strExc(0) = strExc(0) & " select lp01, decode(lp31,'Y',substr(pa75,1,8),substr(pa26,1,8)||nvl(pa149,cu127)) key" & _
         " From caseprogress c1, patent, letterprogress, CUSTOMER" & _
         " where " & IIf(m_ProState = "CFP", "cp05>=", "cp27=") & DBDATE(Text1(0)) & " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "' and substr(cp12,1,1)<>'F' " & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   End If
   strExc(0) = strExc(0) & " union select lp01,substr(lp33,1,8)||lp34 from caseprogress c1,caseprogress c2,letterprogress" & _
      " where " & IIf(m_ProState = "CFP", "c1.cp05>=", "c1.cp27=") & DBDATE(Text1(0)) & " and c1.cp27>=c1.cp05 and c1.cp10='990'" & stCon & _
      " and c1.cp01='" & Text1(1) & "' and substr(c1.cp12,1,1)<>'F' and c2.cp09(+)=c1.cp43 and c2.cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
      " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N') group by key having count(*)=1"
   
   strSql = "update letterprogress set lp27='QPGMR',lp28=sysdate where lp01 in (" & strExc(0) & ") and lp28 is null"
   cnnConnection.Execute strSql, intI
   
   'modify by sonia 2016/7/12 年費逾期通知只限台灣案
   'strExc(0) = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
      " From caseprogress, letterprogress" & _
      " where cp27=" & DBDATE(Text1(0)) & " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='P' and substr(cp12,1,1)<>'F'" & _
      " and lp01(+)=cp09 and lp10='Y' and lp28 is null"
   'modified by Morgan 2016/11/16 +副本語法,已加lp32判斷大批發文不必再判斷國家
   'If Combo1.ItemData(Combo1.ListIndex) = "1605" Then
   '   strExc(0) = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
   '      " From caseprogress, letterprogress, patent" & _
   '      " where cp27=" & DBDATE(Text1(0)) & " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='P' and substr(cp12,1,1)<>'F'" & _
   '      " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp28 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09='000'"
   'Else
   '   strExc(0) = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
   '      " From caseprogress, letterprogress" & _
   '      " where cp27=" & DBDATE(Text1(0)) & " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='P' and substr(cp12,1,1)<>'F'" & _
   '      " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp28 is null"
   'End If
   'Modified by Morgan 2019/2/20 CFP案要確認報價,發文日可能會不同,改以收文日判斷
   'Modified by Morgan 2020/1/3 +所別條件
   'Modified by Morgan 2020/4/16 +stCon 程序人員條件
   'Add By Sindy 2025/4/16 +內商
   If m_ProState = "T" Then
      strExc(0) = "select * from (" & _
         " select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,st06" & _
         " From caseprogress c1, letterprogress, trademark, CUSTOMER, staff" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N' and lp28 is null" & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9)" & _
         " and st01(+)=cu13"
       strExc(0) = strExc(0) & " union " & _
         " select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,st06" & _
         " From caseprogress c1, letterprogress, servicepractice, CUSTOMER, staff" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N' and lp28 is null" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9)" & _
         " and st01(+)=cu13"
       strExc(0) = strExc(0) & " union " & _
         " select c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||'(副本)' CNo,st06" & _
         " from caseprogress c1,caseprogress c2,letterprogress,customer,staff" & _
         " where c1.cp27=" & DBDATE(Text1(0)) & " and c1.cp27>=c1.cp05 and c1.cp10='990'" & stCon & _
         " and c1.cp01='" & Text1(1) & "' and c2.cp09(+)=c1.cp43 and c2.cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N' and lp28 is null" & _
         " and cu01(+)=substr(lp33,1,8) and cu02(+)=substr(lp33,9)" & _
         " and st01(+)=cu13) X where 1=1" & strExc(1)
   Else
   '2025/4/16 END
      strExc(0) = "select * from (" & _
         " select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,decode(instr(';'||oMan,st01),0,st06,'1') st06" & _
         " From caseprogress c1, letterprogress, patent, CUSTOMER, staff,setspecman" & _
         " where " & IIf(m_ProState = "CFP", "cp05>=", "cp27=") & DBDATE(Text1(0)) & " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "' and substr(cp12,1,1)<>'F'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N' and lp28 is null" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
         " and st01(+)=cu13 and ocode(+)='A7'" & _
         " union select c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||'(副本)' CNo,decode(instr(';'||oMan,st01),0,st06,'1') st06" & _
         " from caseprogress c1,caseprogress c2,letterprogress,customer,staff,setspecman" & _
         " where " & IIf(m_ProState = "CFP", "c1.cp05>=", "c1.cp27=") & DBDATE(Text1(0)) & " and c1.cp27>=c1.cp05 and c1.cp10='990'" & stCon & _
         " and c1.cp01='" & Text1(1) & "' and substr(c1.cp12,1,1)<>'F' and c2.cp09(+)=c1.cp43 and c2.cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N' and lp28 is null" & _
         " and cu01(+)=substr(lp33,1,8) and cu02(+)=substr(lp33,9)" & _
         " and st01(+)=cu13 and ocode(+)='A7') X where 1=1" & strExc(1)
      'end 2016/11/16
      'end 2016/7/12
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(0) = ""
      Do While Not RsTemp.EOF
         intI = intI + 1
         If intI > 10 Then
            strExc(0) = strExc(0) & vbCrLf & "...等" & RsTemp.RecordCount & "案"
            Exit Do
         Else
            strExc(0) = strExc(0) & vbCrLf & RsTemp(0)
         End If
         RsTemp.MoveNext
      Loop
      MsgBox "下列案件尚未檢核不可列印清單！" & vbCrLf & strExc(0), vbExclamation
      Exit Sub
   End If
   
   'Added by Morgan 2019/3/8
   'Memo by Morgan 2022/2/15 改說明
   'E化無紙本案件數(半E及全E)
   iNoPaperCount = 0
   'Modified by Morgan 2020/4/16 +stCon 程序人員條件
   'Modify By Sindy 2025/4/16 +內商
   If m_ProState = "T" Then
      strExc(0) = "select nvl(count(*),0) from caseprogress c1,letterprogress" & _
         " where cp27=" & DBDATE(Text1(0)) & _
         " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "' and cp154='QPGMR'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y'"
   Else
   '2025/4/16 END
      strExc(0) = "select nvl(count(*),0) from caseprogress c1,letterprogress" & _
         " where " & IIf(m_ProState = "CFP", "cp05>=", "cp27=") & DBDATE(Text1(0)) & _
         " and cp27>=cp05 and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and cp01='" & Text1(1) & "' and substr(cp12,1,1)<>'F' and cp154='QPGMR'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      iNoPaperCount = RsTemp(0)
   End If
   'end 2019/3/8
   
   'Modified by Morgan 2015/10/8 +判斷大對台案件
   'Modified by Morgan 2016/4/14
   '客戶編號只需抓前8碼判斷
   'modify by sonia 2016/7/12 年費逾期通知只限台灣案,故加入LP27 IS NOT NULL 的條件
   'Modified by Morgan 2016/11/16 +副本語法
   'Modified by Morgan 2019/2/20 CFP案要確認報價,發文日可能會不同,改以收文日判斷
   'Modified by Lydia 2019/12/16 傳入條件strExc(1) : 北所/非北所
   'Modified by Morgan 2020/4/16 +stCon 程序人員條件
   'Modify By Sindy 2025/4/16 +內商
   If m_ProState = "T" Then
      strExc(0) = "select decode(st06,'2','中','3','南','4','高','北') 所別 ,sum(cnt) 件數,count(*) 封數,sum(cnt1) 非直寄件數,sum(decode(cnt1,0,0,1)) 非直寄封數,sum(cnt2) 直寄件數,sum(decode(cnt2,0,0,1)) 直寄封數" & _
         " from (select st06,ToNum,count(*) cnt,sum(decode(lp11,'Y',0,1)) cnt1,sum(decode(lp11,'Y',1,0)) cnt2" & _
         " from (select st06,decode(lp31,'Y',substr(tm44,1,8),substr(tm23,1,8)||nvl(tm123,cu127)) ToNum,cp09,LP11" & _
         " From caseprogress c1, letterprogress, trademark, CUSTOMER, staff" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05" & _
         " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='" & Text1(1) & "'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N' AND lp27 IS NOT NULL " & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9)" & _
         " and st01(+)=cu13"
      strExc(0) = strExc(0) & " union " & _
         "select st06,decode(lp31,'Y',substr(sp26,1,8),substr(sp08,1,8)||nvl(sp78,cu127)) ToNum,cp09,LP11" & _
         " From caseprogress c1, letterprogress, servicepractice, CUSTOMER, staff" & _
         " where cp27=" & DBDATE(Text1(0)) & " and cp27>=cp05" & _
         " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='" & Text1(1) & "'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N' AND lp27 IS NOT NULL " & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9)" & _
         " and st01(+)=cu13"
      strExc(0) = strExc(0) & " union " & _
         "select st06,substr(lp33,1,8)||lp34 ToNum,c1.cp09,lp11" & _
         " from caseprogress c1,caseprogress c2,letterprogress,customer,staff,setspecman" & _
         " where c1.cp27=" & DBDATE(Text1(0)) & " and c1.cp27>=c1.cp05 and c1.cp10='990'" & stCon & _
         " and c1.cp01='" & Text1(1) & "' and c2.cp09(+)=c1.cp43 and c2.cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N' AND lp27 IS NOT NULL " & _
         " and cu01(+)=substr(lp33,1,8) and cu02(+)=substr(lp33,9)" & _
         " and st01(+)=cu13) group by st06, ToNum" & _
         " ) where 1=1 " & strExc(1) & " group by st06"
   Else
   '2025/4/16 END
      strExc(0) = "select decode(st06,'2','中','3','南','4','高','北') 所別 ,sum(cnt) 件數,count(*) 封數,sum(cnt1) 非直寄件數,sum(decode(cnt1,0,0,1)) 非直寄封數,sum(cnt2) 直寄件數,sum(decode(cnt2,0,0,1)) 直寄封數" & _
         " from (select st06,ToNum,count(*) cnt,sum(decode(lp11,'Y',0,1)) cnt1,sum(decode(lp11,'Y',1,0)) cnt2" & _
         " from (select decode(instr(';'||oMan,st01),0,st06,'1') st06,decode(lp31,'Y',substr(pa75,1,8),substr(pa26,1,8)||nvl(pa149,cu127)) ToNum,cp09,LP11" & _
         " From caseprogress c1, letterprogress, patent, CUSTOMER, staff,setspecman" & _
         " where " & IIf(m_ProState = "CFP", "cp05>=", "cp27=") & DBDATE(Text1(0)) & " and cp27>=cp05" & _
         " and cp10='" & Combo1.ItemData(Combo1.ListIndex) & "' and cp01='" & Text1(1) & "' and substr(cp12,1,1)<>'F'" & stCon & _
         " and lp01(+)=cp09 and lp10='Y' and lp32='Y' and lp15='N' AND lp27 IS NOT NULL " & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
         " and st01(+)=cu13 and ocode(+)='A7'" & _
         " union select decode(instr(';'||oMan,st01),0,st06,'1') st06,substr(lp33,1,8)||lp34 ToNum,c1.cp09,lp11" & _
         " from caseprogress c1,caseprogress c2,letterprogress,customer,staff,setspecman" & _
         " where " & IIf(m_ProState = "CFP", "c1.cp05>=", "c1.cp27=") & DBDATE(Text1(0)) & " and c1.cp27>=c1.cp05 and c1.cp10='990'" & stCon & _
         " and c1.cp01='" & Text1(1) & "' and substr(c1.cp12,1,1)<>'F' and c2.cp09(+)=c1.cp43 and c2.cp10='" & Combo1.ItemData(Combo1.ListIndex) & "'" & _
         " and lp01(+)=c1.cp09 and lp10='Y' and lp32='Y' and lp15='N' AND lp27 IS NOT NULL " & _
         " and cu01(+)=substr(lp33,1,8) and cu02(+)=substr(lp33,9)" & _
         " and st01(+)=cu13 and ocode(+)='A7') group by st06, ToNum" & _
         " ) where 1=1 " & strExc(1) & " group by st06"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      PUB_RestorePrinter cmbPrinter
      DoPrint RsTemp, iNoPaperCount
      PUB_RestorePrinter strPrinter
   Else
      'Added by Morgan 2025/11/10
      If Left(Combo3, 5) <> strUserNum And Pub_StrUserSt03 <> "M51" Then
         MsgBox "若為代行期限通知，承辦人會設定為操作人員，程序人員請點選自己！", vbExclamation
      Else
      'end 2025/11/10
         MsgBox "無資料！"
         
      End If
   End If
End Sub

Private Sub DoPrint(pRst As ADODB.Recordset, Optional pNoPaperCount As Integer)
   Dim iOrientation As Integer
   Dim strTemp(7) As String
   Dim iSubTot(7) As Integer
   Dim ii As Integer
   
On Error GoTo ErrHnd
   
   iOrientation = Printer.Orientation
   Printer.PaperSize = 9
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   GetPleft
   
   For ii = 0 To 1
      Erase iSubTot
      Erase strTemp
      
      iPage = 1
      PrintPageHeader ii * lngPageHeight / 2
      PrintPageHeader1
      
      With pRst
      .MoveFirst
      Do While Not .EOF
         strTemp(1) = .Fields("所別")
         strTemp(2) = .Fields("件數")
         strTemp(3) = .Fields("封數")
         strTemp(4) = .Fields("非直寄件數")
         strTemp(5) = .Fields("非直寄封數")
         strTemp(6) = .Fields("直寄件數")
         strTemp(7) = .Fields("直寄封數")
         
         iSubTot(2) = iSubTot(2) + Val("" & .Fields("件數"))
         iSubTot(3) = iSubTot(3) + Val("" & .Fields("封數"))
         iSubTot(4) = iSubTot(4) + Val("" & .Fields("非直寄件數"))
         iSubTot(5) = iSubTot(5) + Val("" & .Fields("非直寄封數"))
         iSubTot(6) = iSubTot(6) + Val("" & .Fields("直寄件數"))
         iSubTot(7) = iSubTot(7) + Val("" & .Fields("直寄封數"))
         PrintDetail strTemp
         .MoveNext
      Loop
      End With
      Call PrintReportFooter(iSubTot, pNoPaperCount)
   Next
   Printer.EndDoc
   Printer.Orientation = iOrientation
   MsgBox "列印完成！"
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description
End Sub

Private Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   ReDim PLeft(1 To 8)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(2, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(3, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(3, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      PrintLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
End Sub

Private Sub PrintLine()
   Dim iNo As Integer
   iNo = (Printer.ScaleWidth - Printer.CurrentX - 500) \ Printer.TextWidth("-")
   Printer.Print String(iNo, "-")
End Sub

Sub PrintPageHeader(Optional pBaseY As Long = 0)
   Dim strPTmp As String
   iPrint = ciStartY + pBaseY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   
   strPTmp = ChangeTStringToTDateString(Text1(0)) & " " & Combo1 & "統計清單"
   
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintPageHeader1()
   Call PrintNewLine(False, 1)
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "所別"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "件數"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "封數"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "非直寄件數"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "非直寄封數"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "直寄件數"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "直寄封數"
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintDetail(strData() As String)
   PrintNewLine
   For intI = 1 To UBound(strData)
      If intI = 1 Then
         Printer.CurrentX = PLeft(intI)
      Else
         Printer.CurrentX = PLeft(intI + 1) - ciColGap - Printer.TextWidth(strData(intI))
      End If
      Printer.CurrentY = iPrint
      Printer.Print strData(intI)
   Next
End Sub

'列印表尾
Private Sub PrintReportFooter(ByRef pSubTot() As Integer, Optional pNoPaperCount As Integer)
    Call PrintNewLine(True, 1)
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    PrintLine
    PrintNewLine
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "合計"
    
    For intI = 2 To UBound(pSubTot)
      If intI = 1 Then
         Printer.CurrentX = PLeft(intI)
      Else
         Printer.CurrentX = PLeft(intI + 1) - ciColGap - Printer.TextWidth(Format(pSubTot(intI)))
      End If
      Printer.CurrentY = iPrint
      Printer.Print Format(pSubTot(intI))
   Next
   
   'Added by Morgan 2019/3/8
   PrintNewLine
   PrintNewLine
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   'Modified by Morgan 2022/2/15
   'Printer.Print "E化不直寄(無紙本)案件數：" & pNoPaperCount
   Printer.Print "E化無紙本案件數：" & pNoPaperCount
   'end 2019/3/8
End Sub
