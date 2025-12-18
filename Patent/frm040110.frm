VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040110 
   BorderStyle     =   1  '單線固定
   Caption         =   "P案各式書表"
   ClientHeight    =   3636
   ClientLeft      =   900
   ClientTop       =   1056
   ClientWidth     =   6708
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3636
   ScaleWidth      =   6708
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   1932
      Left            =   2808
      TabIndex        =   15
      Top             =   1680
      Width           =   3804
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   2616
         MaxLength       =   2
         TabIndex        =   23
         Top             =   60
         Width           =   375
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   2376
         MaxLength       =   1
         TabIndex        =   22
         Top             =   60
         Width           =   255
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1536
         MaxLength       =   6
         TabIndex        =   21
         Top             =   60
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1056
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "P"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdCase 
         Caption         =   "新增"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   3036
         TabIndex        =   19
         Top             =   48
         Width           =   600
      End
      Begin VB.CommandButton cmdCase 
         Caption         =   "刪除"
         Height          =   400
         Index           =   1
         Left            =   3036
         TabIndex        =   18
         Top             =   468
         Width           =   600
      End
      Begin VB.ListBox List1 
         Height          =   1488
         ItemData        =   "frm040110.frx":0000
         Left            =   1056
         List            =   "frm040110.frx":0002
         TabIndex        =   17
         Top             =   384
         Width           =   1932
      End
      Begin VB.Label Label4 
         Caption         =   "多案合併："
         Height          =   252
         Left            =   96
         TabIndex        =   16
         Top             =   120
         Width           =   996
      End
   End
   Begin VB.TextBox txtKind3 
      Height          =   300
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1632
      Width           =   420
   End
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   5520
      TabIndex        =   10
      Top             =   45
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Word編輯(&W)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4320
      TabIndex        =   9
      Top             =   45
      Width           =   1200
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   2745
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   6
      Top             =   405
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   2505
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   405
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1545
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   4
      Top             =   405
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1065
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "P"
      Top             =   405
      Width           =   495
   End
   Begin VB.Label lblKind3 
      Caption         =   "委任書種類：　　　(1.個案  2.總委任書)"
      Height          =   252
      Left            =   105
      TabIndex        =   14
      Top             =   1680
      Width           =   3588
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   324
      Left            =   1068
      TabIndex        =   7
      Top             =   756
      Width           =   5460
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9631;572"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   204
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
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
      Height          =   250
      Left            =   3240
      TabIndex        =   13
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label12 
      Height          =   250
      Left            =   1695
      TabIndex        =   12
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Label Label11 
      Height          =   250
      Left            =   1125
      TabIndex        =   11
      Top             =   1125
      Width           =   525
   End
   Begin VB.Label Label3 
      Caption         =   "申請國家："
      Height          =   250
      Left            =   105
      TabIndex        =   8
      Top             =   1125
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   250
      Left            =   105
      TabIndex        =   2
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   250
      Left            =   105
      TabIndex        =   1
      Top             =   450
      Width           =   975
   End
End
Attribute VB_Name = "frm040110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/09/22 解聘書改成「P案各式書表」，由frm040103_1進入
'Memo by Morgan 2021/12/21 改成Form2.0 (Combo1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim pa() As String
Dim intWhere As Integer
Dim strTemp As String
'Mark by Lydia 2023/09/22 改成下載Word範本套印
'Dim bolRetry As Boolean '是否已發生錯誤且重試
'Dim AppNo As String
'Dim m_Dept As String
'Dim AppDate As String
'Dim m_CP44 As String
'Dim m_CP12 As String
'Dim m_CP13 As String
'Dim m_CP27 As String
'Dim m_FA04 As String
'end 2023/09/22
'Added by Lydia 2023/09/22
Dim m_PrevForm As Form, m_strKind As String  '前一畫面、種類
Dim m_DefPath As String, m_FileName As String  '預設下載檔案路徑、檔名
Dim m_CP09 As String, m_CP10 As String, m_CP31 As String 'Added by Lydia 2024/03/21 前一畫面收文號,案件性質, 是否為新案

'Added by Lydia 2023/09/22
'Modified by Lydia 2024/03/21 +pCP09
Public Sub SetParent(ByRef pForm As Form, ByVal pKind As String, ByVal pCP09 As String, ByRef pData() As String)
   Set m_PrevForm = pForm
   m_strKind = pKind
   pa = pData
   m_CP09 = pCP09
End Sub

'讀取案件資料
Private Function CaseNoCheck() As Boolean
   Dim Cancel As Boolean
   
   If Combo1.ListCount = 0 Then
      Text3_Validate Cancel
      Text4_LostFocus
   End If
   CaseNoCheck = True
End Function

'加所有申請人選項
''Mark by Lydia 2023/09/22
'Private Function GetCustName(strCaseNo As String, strLang As String, Optional ByVal bolAll As Boolean = False, Optional ByVal preStr As String, Optional ByVal bolNum As Boolean = False) As String
'On Error GoTo ErrHnd
'
'   If bolAll = True Then
'      strSql = "select pa26,pa27,pa28,pa29,pa30 from patent where " & ChgPatent(strCaseNo)
'      strSql = strSql & " Union Select tm23,tm78,tm79,tm80,tm81 From Trademark Where " & ChgTradeMark(strCaseNo)
'      strSql = strSql & " Union Select LC11,null,null,null,null From Lawcase Where " & ChgLawcase(strCaseNo)
'      strSql = strSql & " Union Select HC11,null,null,null,null From Hirecase Where " & ChgHirecase(strCaseNo)
'      strSql = strSql & " Union Select sp08,sp58,sp59,sp65,sp66 From ServicePractice Where " & ChgService(strCaseNo)
'
'      strSql = "select C1.CU04, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90, C1.CU06" & _
'         ",C2.CU04, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90, C2.CU06" & _
'         ",C3.CU04, C3.CU05||' '||C3.CU88||' '||C3.CU89||' '||C3.CU90, C3.CU06" & _
'         ",C4.CU04, C4.CU05||' '||C4.CU88||' '||C4.CU89||' '||C4.CU90, C4.CU06" & _
'         ",C5.CU04, C5.CU05||' '||C5.CU88||' '||C5.CU89||' '||C5.CU90, C5.CU06" & _
'         " from (" & strSql & ") X,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5" & _
'         " where C1.CU01(+)=substr(PA26,1,8) And C1.CU02(+)=substr(PA26,9,1)" & _
'         " and C2.CU01(+)=substr(PA27,1,8) And C2.CU02(+)=substr(PA27,9,1)" & _
'         " and C3.CU01(+)=substr(PA28,1,8) And C3.CU02(+)=substr(PA28,9,1)" & _
'         " and C4.CU01(+)=substr(PA29,1,8) And C4.CU02(+)=substr(PA29,9,1)" & _
'         " and C5.CU01(+)=substr(PA30,1,8) And C5.CU02(+)=substr(PA30,9,1)"
'   Else
'      strSql = "Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Patent Where substr(PA26,1,8)=CU01 And substr(PA26,9,1)=CU02 And " & ChgPatent(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Trademark Where substr(TM23,1,8)=CU01 And substr(TM23,9,1)=CU02 And " & ChgTradeMark(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Lawcase Where substr(LC11,1,8)=CU01 And substr(LC11,9,1)=CU02 And " & ChgLawcase(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Hirecase Where substr(HC05,1,8)=CU01 And substr(HC11,9,1)=CU02 And " & ChgHirecase(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, ServicePractice Where substr(SP08,1,8)=CU01 And substr(SP08,9,1)=CU02 And " & ChgService(strCaseNo)
'   End If
'   CheckOC
'   With adoRecordset
'      .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If .RecordCount > 0 Then
'         Select Case strLang
'            Case "1" '中文
'               GetCustName = "" & .Fields(0).Value
'               If bolAll = True Then
'                  For intI = 1 To 4
'                     strExc(1) = Trim("" & .Fields(3 * intI).Value)
'                     If strExc(1) <> "" Then
'                        GetCustName = GetCustName & vbCrLf & preStr & strExc(1)
'                     End If
'                  Next
'               End If
'            Case "2" '英文
'               If Not IsNull(.Fields(1)) Then
'                  If bolAll = False Then
'                     GetCustName = SplitTitle(.Fields(1), GetLen(preStr))
'                  Else
'                     If bolNum = True And Trim("" & .Fields(1 + 3 * 1)) <> "" Then
'                        GetCustName = SplitTitle("1." & .Fields(1), GetLen(preStr))
'                     Else
'                        GetCustName = SplitTitle(.Fields(1), GetLen(preStr))
'                     End If
'                     For intI = 1 To 4
'                        strExc(1) = Trim("" & .Fields(1 + 3 * intI).Value)
'                        If strExc(1) <> "" Then
'                           If bolNum = True Then
'                              GetCustName = GetCustName & vbCrLf & preStr & SplitTitle(intI + 1 & "." & strExc(1), GetLen(preStr) + 1)
'                           Else
'                              GetCustName = GetCustName & vbCrLf & preStr & SplitTitle(strExc(1), GetLen(preStr))
'                           End If
'                        End If
'                     Next
'                  End If
'               End If
'
'            Case "3" '日文
'               GetCustName = "" & .Fields(2).Value
'               If bolAll = True Then
'                  For intI = 1 To 4
'                     strExc(1) = Trim("" & .Fields(2 + 3 * intI).Value)
'                     If strExc(1) <> "" Then
'                        GetCustName = GetCustName & vbCrLf & preStr & strExc(1)
'                     End If
'                  Next
'               End If
'         End Select
'      End If
'   End With
'
'ErrHnd:
'   If Err.NUMBER <> 0 Then MsgBox Err.Description
'End Function
'end 2023/09/22

Private Sub Command1_Click()
   '檢查本所案號
   If Text1.Text = "" Or Text2.Text = "" Then
       MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
       Text2.SetFocus
       Text2_GotFocus
       Exit Sub
   End If

   If CaseNoCheck = False Then Exit Sub
   'Added by Lydia 2023/09/22
   If m_strKind = "3" And Trim(txtKind3) = "" Then
      MsgBox "請輸入委任書種類！", vbExclamation
      Exit Sub
   End If
   'end 2023/09/22
   
   'm_Dept = GetStaffDepartment(strUserNum) 'Mark by Lydia 2023/09/22
   
   'Memo by Lydia 2023/09/22 刪除列印聯絡單 PrintContactSheet
   
   Screen.MousePointer = vbHourglass
   'Modified by Lydia 2023/09/22　原本表單只有解聘書並且限制大陸案可以使用，現在有1.解聘書2.大陸案委託書3.委任書4.讓與契約書5.簽章切結書
   'WordChinese
   'Modified by Lydia 2023/10/03 改放在桌面
   'm_DefPath = App.path & "\" & strUserNum
   m_DefPath = strExcelPath
   Pub_ChkExcelPath m_DefPath

   Select Case m_strKind
      Case "1"  '(大陸案)解聘書
         runWordProc1
      Case "2"  '大陸案委託書
         'Added by Lydia 2024/03/26
         If pa(9) = "056" Then
            runWordProc2_1 'PCT案委托書
         Else
         'end 2024/03/26
            runWordProc2
         End If
      Case "3"  '委任書
         runWordProc3
      Case "4"  '讓與契約書
         If List1.ListCount = 0 Then
            If pa(9) = "000" Then
               runWordProc4_1
            Else
               runWordProc4_3
            End If
         Else
            runWordProc4_2
         End If
      Case "5"  '簽章切結書
         If List1.ListCount = 0 Then
            runWordProc5_1
         Else
            runWordProc5_2
         End If
   End Select
   'end 2023/09/22
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
   'ReDim pa(TF_PA) As String 'Mark by Lydia 2023/09/22
End Sub

Private Sub Form_Load()
   'Added by Lydia 2023/09/22
   strTemp = ""
   Select Case m_strKind
      Case "1": strTemp = "解聘書"
      Case "2": strTemp = "大陸案委託書"
      Case "3": strTemp = "委任書"
      Case "4": strTemp = "讓與契約書"
      Case "5": strTemp = "簽章切結書"
   End Select
   If strTemp <> "" Then
      Me.Caption = Me.Caption & "-" & strTemp
   End If
   If InStr("1,2,", m_strKind) > 0 Then
      Me.Height = 1930
   Else
      If m_strKind = "3" Then
         Me.Height = 2460
         lblKind3.Visible = True
         txtKind3.Visible = True
         Frame1.Visible = False
      Else
         Me.Height = 4080
         lblKind3.Visible = False
         txtKind3.Visible = False
         Frame1.Visible = True
         Frame1.Left = lblKind3.Left
      End If
   End If
   
   Call ReadData
   If pa(9) <> "000" Then  '大陸案沒有多案合併
      Frame1.Visible = False
      If InStr("1,2,", m_strKind) = 0 Then
         Me.Height = 2460
      End If
   End If
   'end 2023/09/22
   
   MoveFormToCenter Me
   lblClose.Caption = ""
   
   'Memo by Lydia 2023/09/22 刪除批次作業cmdBatch

End Sub

Private Sub Form_Unload(Cancel As Integer)

   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
      m_PrevForm.ClearForm
   End If
   Set frm040110 = Nothing
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
    CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   'Mark by Lydia 2023/09/22
'   Dim strTemp1
'   Dim strTemp2
'   Dim ii As Integer
'   Dim jj As Integer
'   Dim ss As Integer
'
'   If Text1.Text = "" Then Exit Sub
'   strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
'   strTemp2 = Split(Replace(UCase(Text1.Text), ",,", ""), ",")
'   For ii = 0 To UBound(strTemp2)
'       ss = 0
'       For jj = 0 To UBound(strTemp1)
'           If strTemp2(ii) = strTemp1(jj) Then
'               ss = 1
'               Exit For
'           End If
'       Next jj
'       If ss = 0 Then
'          '開放FF案件之權限
'          m_Dept = GetStaffDepartment(strUserNum)
'          Select Case m_Dept
'             Case "F21", "F23", "F61", "F81"  '開放F21,F23使用P,PS,CFP,CPS權限
'                If Text1.Text = "P" Or Text1.Text = "PS" Or Text1.Text = "CFP" Or Text1.Text = "CPS" Then
'                   Exit For
'                End If
'             Case "F10", "F11"    '開放F10,F11使用T權限
'                If Text1.Text = "T" Then
'                   Exit For
'                End If
'          End Select
'          '檢查跨部門權限
'          If CheckSR09(strUserNum, Text1, "Y", False, Text1, Text2, Text3, Text4) = True Then
'             Exit For
'          End If
'          ss = MsgBox(strUserName & " 沒有 " & strTemp2(ii) & " 的權限!! ", , "USER 權限問題")
'          Text1.SetFocus
'          Text1_GotFocus
'          Cancel = True
'       End If
'   Next ii
    'end 2023/09/22
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
    CloseIme
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
End Sub

Private Sub Text3_GotFocus()
    TextInverse Text3
    CloseIme
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
End Sub

Private Sub Text4_GotFocus()
    TextInverse Text4
    CloseIme
End Sub

Private Sub Text4_LostFocus()
    If Text3 = "" Then Text3 = "0"
    If Text4 = "" Then Text4 = "00"
    'Read 'Mark by Lydia 2023/09/22
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
End Sub

'Modified by Lydia 2023/09/22 改成從外部傳入=>ReadData
'Private Sub Read()
'   Combo1.Clear
'   lblClose.Caption = ""
'   Label11.Caption = ""
'   Label12.Caption = ""
'   pa(1) = Text1
'   pa(2) = Text2
'   If Text3 = "" Then
'      pa(3) = "0"
'   Else
'      pa(3) = Text3
'   End If
'   If Text4 = "" Then
'      pa(4) = "00"
'   Else
'      pa(4) = Text4
'   End If
'
'   AppNo = "": AppDate = ""
'   Command1.Enabled = False
'
'   Select Case pa(1) '判斷系統類別
'      Case "P", "CFP", "FCP" '專利
'          If ClsPDReadPatentDatabase(pa(), intWhere) Then
'              '若有基本資料
'              If Not IsNull(pa()) Then
'                  '案件名稱
'                  If IsNull(pa(5)) = False And pa(5) <> "" Then Combo1.AddItem pa(5)
'                  If IsNull(pa(6)) = False And pa(6) <> "" Then Combo1.AddItem pa(6)
'                  If IsNull(pa(7)) = False And pa(7) <> "" Then Combo1.AddItem pa(7)
'                  If IsNull(pa(5)) = False And pa(5) <> "" Then
'                     Combo1 = pa(5)
'                  ElseIf IsNull(pa(6)) = False And pa(6) <> "" Then
'                     Combo1 = pa(6)
'                  ElseIf IsNull(pa(7)) = False And pa(7) <> "" Then
'                     Combo1 = pa(7)
'                  End If
'                  '申請案號
'                  If IsNull(pa(11)) = False And pa(11) <> "" Then AppNo = pa(11)
'                  '申請日
'                  If IsNull(pa(10)) = False And pa(10) <> "" Then AppDate = Mid(DBDATE(pa(10)), 1, 4) & "年" & Mid(DBDATE(pa(10)), 5, 2) & "月" & Mid(DBDATE(pa(10)), 7, 2) & "日"
'                  '是否閉卷
'                  If Len("" & pa(57)) <= 0 Then
'                     lblClose.Caption = ""
'                  Else
'                     lblClose.Caption = "已閉卷"
'                  End If
'                  '抓國家名稱
'                  Label11 = pa(9)
'                  If ClsPDGetNation(pa(9), strTemp) Then Label12.Caption = strTemp
'
'                  '非台灣案才可作業
'                  If pa(1) = "P" And pa(9) <> "000" Then
'                     strSql = "select cp44,cp12,cp13,cp27,FA04 from caseprogress,fagent " & _
'                                  "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
'                                  "and cp09<='C' and cp27 is not null and cp44 is not null " & _
'                                  "and fa01(+)=substr(CP44,1,8) and fa02(+)=substr(CP44,9,1) " & _
'                                  "order by cp27 desc "
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     m_CP44 = "": m_CP12 = "": m_CP13 = "": m_CP27 = "": m_FA04 = ""
'                     If intI = 1 Then
'                        m_CP44 = "" & RsTemp("cp44")
'                        m_CP12 = "" & RsTemp("cp12")
'                        m_CP13 = "" & RsTemp("cp13")
'                        m_CP27 = "" & RsTemp("cp27")
'                        m_FA04 = "" & RsTemp("FA04")
'                     End If
'                  Else
'                     MsgBox "須為P非台灣案，才可執行！", vbExclamation + vbOKOnly
'                     Text2.SetFocus
'                     Text2_GotFocus
'                     Exit Sub
'                  End If
'
'                  Command1.Enabled = True
'              End If
'          Else
'              Text2.SetFocus
'              Text2_GotFocus
'          End If
'   End Select
'End Sub
Private Sub ReadData()

   Combo1.Clear
   lblClose.Caption = ""
   Label11.Caption = ""
   Label12.Caption = ""
   Command1.Enabled = False
   '若有基本資料
   If Not IsNull(pa()) Then
      Text1 = pa(1): Text2 = pa(2): Text3 = pa(3): Text4 = pa(4)
      '案件名稱
      If IsNull(pa(5)) = False And pa(5) <> "" Then Combo1.AddItem pa(5)
      If IsNull(pa(6)) = False And pa(6) <> "" Then Combo1.AddItem pa(6)
      If IsNull(pa(7)) = False And pa(7) <> "" Then Combo1.AddItem pa(7)
      If IsNull(pa(5)) = False And pa(5) <> "" Then
         Combo1 = pa(5)
      ElseIf IsNull(pa(6)) = False And pa(6) <> "" Then
         Combo1 = pa(6)
      ElseIf IsNull(pa(7)) = False And pa(7) <> "" Then
         Combo1 = pa(7)
      End If
      '是否閉卷
      If Len("" & pa(57)) <= 0 Then
         lblClose.Caption = ""
      Else
         lblClose.Caption = "已閉卷"
      End If
      '抓國家名稱
      Label11 = pa(9)
      If ClsPDGetNation(pa(9), strTemp) Then Label12.Caption = strTemp
      Command1.Enabled = True
      'Added by Lydia 2024/03/21
      strExc(0) = "select cp09,cp10,cp31 from caseprogress where cp09='" & m_CP09 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          m_CP10 = "" & RsTemp.Fields("cp10")
          m_CP31 = "" & RsTemp.Fields("cp31")
      End If
      'end 2024/03/21
   End If
End Sub
'end 2023/09/22

''Mark by Lydia 2023/09/22 (大陸案)解聘書=>改成下載Word範本套印runWordProc1
'Private Sub WordChinese()
'
'   bolRetry = False
'
'On Error GoTo ERRORSECTION1
'
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'   g_WordAp.Documents.add
'   With g_WordAp
'
'      .Selection.Font.Name = "標楷體"
'      .Selection.PageSetup.Orientation = wdOrientPortrait
'      .Selection.Orientation = wdTextOrientationHorizontal
'      .Selection.Font.Size = 22
'      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3)
'      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2.5)
'      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2.5)
'      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
'      .Selection.ParagraphFormat.DisableLineHeightGrid = True
'      '固定行高
'      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
'      .Selection.ParagraphFormat.LineSpacing = 17 '15
'
'      '補滿5行
'      For intI = 1 To 6
'         .Selection.TypeParagraph
'      Next
'
'      '第一頁
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
'      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceAtLeast 'Add by Morgan 2011/9/2 標題改最小行高,否則畫面顯示會不完整
'      .Selection.TypeText "解　　聘　　書"
'      .Selection.TypeParagraph
'      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly 'Add by Morgan 2011/9/2
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.Font.Size = 16
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
'      .Selection.TypeText "申　請　號：" & AppNo
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "申　請　日：" & AppDate
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "名　　　稱：" & Combo1.Text
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "申　請　人：" & GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1", True, "　　　　　　")
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　本申請人／專利權人的上述專利／專利申請原委托："
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      '底線 Star
'      .Selection.Font.Underline = wdUnderlineSingle
'      .Selection.TypeText "　　" & m_FA04 & "　　"
'      .Selection.Font.Underline = wdUnderlineNone
'      '底線 End
'      .Selection.TypeText "　代理。"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "現解除代理關係。"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "特此聲明。"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      'Modified by Morgan 2012/11/13 +申請人／
'      '.Selection.TypeText "　　　　　　　　　　　　專利權人：" & GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1", True, "　　　　　　　　　　　　　　　　　")
'      .Selection.TypeText "　　　　　　　　申請人／專利權人：" & GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1", True, "　　　　　　　　　　　　　　　　　")
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　　　　　　　　　　　　年　　月　　日"
'      .Selection.TypeParagraph
'
'      '第二頁
'      .Selection.InsertBreak Type:=wdPageBreak '分隔設定
'      .Selection.Font.Size = 16
'      '粗體 Star
'      .Selection.Font.Bold = wdToggle
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
'      .Selection.TypeText "專 利 代 理 委 托 書（中英文）"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "POWER OF ATTORNEY"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.Font.Bold = wdToggle
'      '粗體 End
'      .Selection.Font.Size = 12
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
'      .Selection.TypeText "　　我/我們是_____的公民/法人，根據中華人民共和國專利法第19條的規定，茲委托______________________________________________________（機構代碼__________），並由該機構指定其專利代理人__________、__________代為辦理名稱為___________________________________________________________________________申請號（或專利號）/國際申請號為_________________的專利申請在中華人民共和國的全部專利事宜。"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "Pursuant to the Article 19 of the Patent Law of the People's Republic of China, I/we, citizen/legal entity of _____ hereby authorize _________________________________ (Code: _________) to appoint its patent attorney(s) __________，_________ to handle all patent affairs related to the application with title as ________________________________________________ and application number(or patent number)/international application number as____________ in the PRC."
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　委托人姓名或名稱　　　　　　　" & GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1", True, "　　　　　　　　　　　　　　　　　")
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　Authorized by (Name)         ___________________________________"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　委托人簽字或蓋章"
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　Signature or Seal            ___________________________________"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　被委托專利代理機構蓋章"
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　Seal of the Authorized Agent ___________________________________"
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　委托日期"
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　Date of Authorization        ___________________________________"
'
'      '英數字改字體
'      If Mid(m_Dept, 1, 2) = "P1" Then
'         .Selection.WholeStory
'         .Selection.Font.Name = "Arial"
'      End If
'   End With
'
''   Select Case Left(m_Combo8, 1)
''      Case "2", "3", "4", "5", "6", "8"
''         PhaseIndent    '調整首行凸排
''   End Select
'
'   g_WordAp.Visible = True
'   g_WordAp.WindowState = wdWindowStateMaximize
'
'   Set g_WordAp = Nothing 'Added by Morgan 2021/12/21 用完清除，否則第二次開會多一的空白頁
'
'ERRORSECTION1:
'
'   If Err.NUMBER <> 0 Then
'      Select Case Err.NUMBER
'         Case 91, 462:
'            Set g_WordAp = New Word.Application
'            g_WordAp.Documents.add
'            If bolRetry = False Then
'               bolRetry = True
'               Resume
'            End If
'         Case Else:
'            MsgBox "錯誤 : " & Err.Description, vbCritical
'      End Select
'   End If
'End Sub
'end 2023/09/22

''Mark by Lydia 2023/09/22
'Private Function SplitTitle(stTitle As String, iReserve As Integer, Optional iMax As Integer = 82) As String
'   Dim arrWords '單字陣列
'   Dim ii As Integer
'   Dim iUsable As Integer
'   Dim iRest As Integer
'   Dim strTmp As String
'   Dim strWord As String
'   Dim iLen As Integer
'
'   iUsable = iMax - iReserve
'   stTitle = Trim(stTitle)
'   If stTitle <> "" Then
'      arrWords = Split(stTitle, " ")
'      strTmp = ""
'      iRest = iUsable
'       For ii = LBound(arrWords) To UBound(arrWords)
'         strWord = arrWords(ii)
'         Do While strWord <> ""
'            iLen = GetLen(strWord)
'            '超過最大可印長度時,接續印並斷字後跳行
'            If iLen > iUsable Then
'               strTmp = strTmp & " " & GetWord(strWord, iRest, strWord) & vbCrLf & String(iReserve, " ")
'               iRest = iUsable
'            '超過剩餘最大可印長度時,跳行後列印
'            ElseIf iLen > iRest - 1 Then
'               strTmp = strTmp & vbCrLf & String(iReserve, " ") & strWord
'               iRest = iUsable - iLen
'               strWord = ""
'            '可印
'            Else
'               If iRest = iUsable Then
'                  strTmp = strTmp & strWord
'                  iRest = iRest - iLen
'               Else
'                  strTmp = strTmp & " " & strWord
'                  iRest = iRest - iLen - 1
'               End If
'               strWord = ""
'            End If
'         Loop
'       Next
'   End If
'   SplitTitle = strTmp
'End Function

''Mark by Lydia 2023/09/22
'Private Function GetLen(strWord As String) As Integer
'   Dim stChar As String
'   Dim iLen As Integer
'   Dim ii As Integer
'   For ii = 1 To Len(strWord)
'      stChar = Mid(strWord, ii, 1)
'      '全形字 2
'      If Asc(stChar) < 0 Then
'         iLen = iLen + 2
'      '英文大寫 1.5
'      ElseIf Asc(stChar) >= 65 And Asc(stChar) <= 90 Then
'         iLen = iLen + 1.5
'      '其他 1
'      Else
'         iLen = iLen + 1
'      End If
'   Next
'   GetLen = iLen
'End Function

''Mark by Lydia 2023/09/22
'Private Function GetWord(ByVal strWord As String, ByVal iMaxLen As Integer, ByRef strWord2 As String) As String
'   Dim stChar As String
'   Dim iLen As Integer
'   Dim strTmp As String
'   Dim ii As Integer
'
'   For ii = 1 To Len(strWord)
'      stChar = Mid(strWord, ii, 1)
'      '全形字 2
'      If Asc(stChar) < 0 Then
'         iLen = iLen + 2
'      '英文大寫 1.5
'      ElseIf Asc(stChar) >= 65 And Asc(stChar) <= 90 Then
'         iLen = iLen + 1.5
'      '其他 1
'      Else
'         iLen = iLen + 1
'      End If
'      If iLen > iMaxLen Then
'         Exit For
'      Else
'         strTmp = strTmp & stChar
'      End If
'   Next
'   If strTmp = strWord Then
'      strWord2 = ""
'   Else
'      strWord2 = Mid(strWord, ii)
'   End If
'   GetWord = strTmp
'End Function
'end 2023/09/24

'Remove by Lydia 2023/09/22 刪除模組:InitPrtPosition, PrintContactSheet1, PrintContactSheet, PhaseIndent

'Added by Lydia 2023/09/22
Private Sub txtKind3_GotFocus()
   TextInverse txtKind3
End Sub

Private Sub txtKind3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2023/09/22
Private Sub cmdCase_Click(Index As Integer)
Dim strTmp As String, strMsg As String

   If Index = 0 And txtCode(2).Text <> "" Then
      strTmp = txtCode(1) & txtCode(2)
      If txtCode(3).Text = "" Then
         strTmp = strTmp & "0"
      Else
         strTmp = strTmp & txtCode(3).Text
      End If
      If txtCode(4).Text = "" Then
         strTmp = strTmp & "00"
      Else
         strTmp = strTmp & txtCode(4).Text
      End If
      If strTmp = pa(1) & pa(2) & pa(3) & pa(4) Then
         strMsg = "不可輸入相同案號！"
         GoTo EXITSUB
      End If
      If List1.ListCount > 0 Then
         For intI = 1 To List1.ListCount
            If strTmp = Trim(List1.List(intI - 1)) Then
               strMsg = "不可輸入相同案號！"
               GoTo EXITSUB
            End If
         Next intI
      End If
      
      intI = 1
      strExc(0) = "SELECT pa09 FROM PATENT WHERE " & ChgPatent(strTmp)
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If "" & RsTemp.Fields("pa09") <> "000" Then
            strMsg = "非臺灣案無多案合併！"
            GoTo EXITSUB
         Else
            List1.AddItem strTmp
            txtCode(2).Text = ""
         End If
      Else
         MsgBox "案號不存在，請重新輸入 !", vbCritical
      End If
      txtCode(2).SetFocus
   Else
      If List1.ListIndex > -1 Then List1.RemoveItem List1.ListIndex
   End If
   
   Exit Sub
   
EXITSUB:
   If strMsg <> "" Then
      MsgBox strMsg, vbExclamation
   End If
End Sub

'Added by Lydia 2023/09/22 大陸案解聘書：下載Word範本套印
Private Sub runWordProc1()
Dim strName As String
Dim strText
Dim intA As Integer
Dim tmpArr As Variant

On Error GoTo ErrHand
       
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000401-0-01 內專P大陸案解聘書.docx", "M51", "000401", "0", "01", "4", "1")

   '下載範本檔
   'Modified by Lydia 2023/10/03 改副檔名為英文
   'm_FileName = "$$" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "解聘書.docx"
   m_FileName = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & ".ATT.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000401-0-01", , m_DefPath) = False Then
      Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 6
         strName = "PS" & Format(intA, "00")
         strText = ""
         If intA = 1 Then
            '申請案號
            strText = pa(11)
         ElseIf intA = 2 Then
            '申請日
            If pa(10) <> "" Then
               strText = Mid(DBDATE(pa(10)), 1, 4) & "年" & Mid(DBDATE(pa(10)), 5, 2) & "月" & Mid(DBDATE(pa(10)), 7, 2) & "日"
            End If
         ElseIf intA = 3 Then
            '名稱
            If Combo1.Text <> "" Then
               strText = Combo1.Text
            Else
               For intI = 5 To 7
                  If pa(intI) <> "" Then
                     strText = pa(intI)
                     Exit For
                  End If
               Next intI
            End If
         ElseIf intA = 4 Then
            '申　請　人:含標題固定5行高
            strText = "申　請　人：" & Replace(GetCuXname("1", pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), "　　　　　　"), "|", vbCrLf)
            tmpArr = Empty
            tmpArr = Split(strText, vbCrLf)
            If UBound(tmpArr) < 4 Then
               For intI = UBound(tmpArr) To 3
                  strText = strText & vbCrLf
               Next intI
            End If
         ElseIf intA = 5 Then
              '代理人
              If PUB_GetCP44(pa(1), pa(2), pa(3), pa(4), strExc(1), strExc(2), , strExc(3)) = True Then
                 strText = String(2, "　") & strExc(3) & String(2, "　")
              End If
         ElseIf intA = 6 Then
              '申請人／專利權人
              strText = String(6, "　") & "申請人／專利權人：" & Replace(GetCuXname("1", pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), String(15, "　")), "|", vbCrLf)
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
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
            .Selection.Delete
            If intA = 5 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = True
            End If

            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            If intA = 5 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = False
            End If
         End If
      Next intA
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2023/09/22 大陸案委托書：下載Word範本套印
Private Sub runWordProc2()
Dim strName As String
Dim strText
Dim intA As Integer
Dim tmpArr As Variant

On Error GoTo ErrHand
       
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000401-0-02 內專P大陸案委托書.docx", "M51", "000401", "0", "02", "4", "1")

   '下載範本檔
   'Modified by Lydia 2023/10/03 改副檔名為英文
   'm_FileName = "$$" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "委托書.docx"
   m_FileName = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & ".POA.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000401-0-02", , m_DefPath) = False Then
      Exit Sub
   End If
   
   'Added by Lydia 2024/03/21 讓與案(701,708)的委任書之委任人，非新案改抓進度的受讓人，新案才抓申請人=委任人
   Dim strCuA As String '讓與人
   Dim strCuB As String '受讓人(讓與申請人)
   Dim strCuNow As String '要帶入的委任人
   If InStr("701,708", m_CP10) > 0 And m_CP31 <> "Y" Then
       Call GetXcuList(strCuA, strCuB)
       strCuNow = strCuB
   Else
       strCuNow = pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30)
   End If
   'end 2024/03/21
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 3
         strName = "PS" & Format(intA, "00")
         strText = ""
         If intA = 1 Then
            '名稱
            If Combo1.Text <> "" Then
               strText = Combo1.Text
            Else
               For intI = 5 To 7
                  If pa(intI) <> "" Then
                     strText = pa(intI)
                     Exit For
                  End If
               Next intI
            End If
            strText = String(2, " ") & strText & String(2, " ")
         ElseIf intA = 2 Then
            '申請案號
            strText = String(2, " ") & IIf(Trim(pa(11)) = "", String(10, " "), pa(11)) & String(2, " ")
         ElseIf intA = 3 Then
            '委託人(申請人)
            'Modified by Lydia 2024/03/21 改用變數
            'strText = Replace(GetCuXname("1", pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30)), "|", vbCrLf)
            strText = Replace(GetCuXname("1", strCuNow), "|", vbCrLf)
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
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
            .Selection.Delete
            If intA < 3 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = True
            End If

            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            If intA < 3 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = False
            End If
         End If
      Next intA
   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2023/09/22 委任書：下載Word範本套印
Private Sub runWordProc3()
Dim strName As String
Dim strText
Dim intA As Integer
Dim tmpArr As Variant

On Error GoTo ErrHand
       
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000401-1-03 內專P台灣案總委任書.docx", "M51", "000401", "1", "03", "4", "1")
   
   '下載範本檔
   'Modified by Lydia 2023/10/03 改副檔名為英文
   'm_FileName = "$$" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & IIf(txtKind3 = "1", "個案", "總") & "委任書.docx"
   m_FileName = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & ".POA.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, IIf(txtKind3 = "1", "M51-000401-0-03", "M51-000401-1-03"), , m_DefPath) = False Then
      Exit Sub
   End If
   
   'Added by Lydia 2024/03/21 讓與案(701,708)的委任書之委任人，非新案改抓進度的受讓人，新案才抓申請人=委任人
   Dim strCuA As String '讓與人
   Dim strCuB As String '受讓人(讓與申請人)
   Dim strCuNow As String '要帶入的委任人
   If InStr("701,708", m_CP10) > 0 And m_CP31 <> "Y" Then
       Call GetXcuList(strCuA, strCuB)
       strCuNow = strCuB
   Else
       strCuNow = pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30)
   End If
   'end 2024/03/21
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To IIf(txtKind3 = "1", 5, 4)
         strName = "PS" & Format(intA, "00")
         strText = ""
         If intA = 1 Then
            '委任人(申請人)
            'Modified by Lydia 2024/03/21 改用變數
            'strText = Replace(GetCuXname("1", pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30)), "|", vbCrLf)
            strText = Replace(GetCuXname("1", strCuNow), "|", vbCrLf)
         ElseIf intA = 2 Then
            '代表人
            'Modified by Lydia 2024/03/21 改用變數
            'strText = Replace(GetCuXname("2", pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30)), "|", vbCrLf)
            strText = Replace(GetCuXname("2", strCuNow), "|", vbCrLf)
         ElseIf intA = 3 Then
            '受任人: 林+閻
            strTemp = ""
            strExc(1) = "select decode(st01,'81040','1','2') ord1 , st01,st02,oa05 from ouragent,staff " & _
                        "where oa01='" & pa(1) & "' and instr('94007,81040',oa02)>0 and st01(+)=oa02 order by 1 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  strExc(1) = "" & RsTemp.Fields("st02")
                  strExc(2) = ""
                  For intI = 1 To Len(strExc(1)) - 1
                     strExc(2) = strExc(2) & Mid(strExc(1), intI, 1) & "　"
                  Next intI
                  strExc(2) = strExc(2) & Right(strExc(1), 1)
                  strText = strText & "," & RsTemp.AbsolutePosition & "." & strExc(2) & "專利師"
                  strTemp = strTemp & "," & RsTemp.AbsolutePosition & "." & RsTemp.Fields("oa05")
                  RsTemp.MoveNext
               Loop
            End If
            If strTemp <> "" Then
               strText = Mid(strText, 2)
               strTemp = Mid(strTemp, 2)
            End If
            strText = Replace(strText, ",", vbCrLf)

         ElseIf intA = 4 Then
            '登記證字號
            strText = Replace(strTemp, ",", vbCrLf)   '在抓受任人一併抓登記證
         ElseIf intA = 5 Then
            '名稱
            If Combo1.Text <> "" Then
               strText = Combo1.Text
            Else
               For intI = 5 To 7
                  If pa(intI) <> "" Then
                     strText = pa(intI)
                     Exit For
                  End If
               Next intI
            End If
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
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
            .Selection.Delete
            '性    質：其他 設定 底線
            .Selection.Font.Underline = False
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
         End If
      Next intA
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2023/09/22 取得申請人名稱/代表人/地址
Private Function GetCuXname(ByVal pKind As String, ByVal pCuList As String, Optional PreStr As String, Optional ByVal bolNum As Boolean = False) As String
'pkind: 1-申請人名稱,2-代表人名稱, 3-地址
'pCuList:X編號用,區隔
'preStr:每一項的預設開頭
'bolNum: 一筆以上加註數字1.
Dim intR As Integer, intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
Dim strMid As String
Dim tmpArr As Variant

   If pCuList = "" Then Exit Function
   
   tmpArr = Empty
   tmpArr = Split(pCuList, ",")
   For intR = 0 To UBound(tmpArr)
      If Trim(tmpArr(intR)) <> "" Then
         'Modified by Lydia 2024/02/27 改抓負責人
         'strQ1 = "select nvl(cu04,decode(cu05,null,cu06,CU05||' '||CU88||' '||CU89||' '||CU90)) cname,nvl(cu39,nvl(cu40,cu41)) aname,nvl(cu23,decode(cu24,null,cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102,cu29)) caddr " & _
                 "From CUSTOMER where cu01='" & Mid(ChangeCustomerL(tmpArr(intR)), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(tmpArr(intR)), 9, 1) & "' "
         strQ1 = "select nvl(cu04,decode(cu05,null,cu06,CU05||' '||CU88||' '||CU89||' '||CU90)) cname,decode(cu15,'0',null,nvl(cu07,nvl(cu39,nvl(cu40,cu41)))) aname,nvl(cu23,decode(cu24,null,cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102,cu29)) caddr " & _
                 "From CUSTOMER where cu01='" & Mid(ChangeCustomerL(tmpArr(intR)), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(tmpArr(intR)), 9, 1) & "' "
         intQ = 1
         Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
         If intQ = 1 Then
            If intR = 1 And strMid <> "" And bolNum = True Then
               strMid = "1." & strMid
            End If
            If "" & rsQD.Fields(Val(pKind) - 1) <> "" Then
               strMid = strMid & IIf(strMid <> "", "|" & PreStr & IIf(bolNum = True, intR + 1 & ".", ""), "") & rsQD.Fields(Val(pKind) - 1)
            End If
         End If
      End If
   Next intR
   GetCuXname = strMid
   
   Set rsQD = Nothing
End Function

'Added by Lydia 2023/09/22 取得讓與人/受讓人
Private Sub GetXcuList(ByRef pToA As String, ByRef pToB As String)
Dim intQ As Integer, strQuery As String
Dim rsQuery As New ADODB.Recordset
   
   pToA = "": pToB = ""
   'Modified by Lydia 2024/03/21 +收文號
   strQuery = "select cp05, cp55||','||cp93||','||cp94||','||cp95||','||cp96 as cust1, cp56||','||cp89||','||cp90||','||cp91||','||cp92 as cust2 " & _
               "from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
               "and cp09<'C' and cp159=0 and cp10 in ('701','708') and cp09='" & m_CP09 & "' order by cp05 desc "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      pToA = "" & rsQuery.Fields("cust1")
      pToB = "" & rsQuery.Fields("cust2")
   Else
      pToA = pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30)
      pToB = pToA
   End If

   Set rsQuery = Nothing
End Sub

'Added by Lydia 2023/09/22 台灣案讓與契約書：下載Word範本套印
Private Sub runWordProc4_1()
Dim strName As String
Dim strText
Dim intA As Integer
Dim strCuA As String '讓與人
Dim strCuB As String '受讓人(讓與申請人)
Dim intMax As Single
Dim tmpArr As Variant

On Error GoTo ErrHand
       
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000401-0-04 內專P台灣案讓與契約書_個案.docx", "M51", "000401", "0", "04", "4", "1")

   '下載範本檔
   'Modified by Lydia 2023/10/03 改副檔名為英文
   'm_FileName = "$$" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "讓與契約書.docx"
   m_FileName = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & ".ASSIGNMENT.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000401-0-04", , m_DefPath) = False Then
      Exit Sub
   End If

   Call GetXcuList(strCuA, strCuB)
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 10
         strName = "PS" & Format(intA, "00")
         strText = ""
         If intA = 1 Then
             '專利種類
             Call ClsPDGetPatentTrademarkKind("1", pa(8), strTemp, IIf(pa(9) <> "000", True, False), pa(9))
             strExc(1) = ""
             For intI = 1 To Len(strTemp)
               strExc(1) = strExc(1) & "|" & Mid(strTemp, intI, 1)
             Next intI
             strText = Replace(Mid(strExc(1), 2), "|", "　")
         ElseIf intA = 2 Then
            '專利證號/申請案號
            If pa(15) <> "" Then
               strText = pa(15)
            Else
               strText = pa(11)
            End If
         ElseIf intA = 3 Then
            '名稱
            If Combo1.Text <> "" Then
               strText = Combo1.Text
            Else
               For intI = 5 To 7
                  If pa(intI) <> "" Then
                     strText = pa(intI)
                     Exit For
                  End If
               Next intI
            End If
         ElseIf intA = 4 Then
            '受讓人
            strTemp = Replace(GetCuXname("1", strCuB), "|", "、")
            If GetTextLength(strTemp) < 30 Then
               strText = PUB_StrToStr(strTemp, 30, True)
            Else
               strText = strTemp
            End If
         ElseIf intA = 5 Then
            '讓與人：欄寬36字元 =>strCuB受讓人(讓與申請人)
            intMax = 36
            strTemp = GetCuXname("1", strCuA)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 6 Then
            '代表人
            strTemp = GetCuXname("2", strCuA)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 7 Then
            '地址
            strTemp = GetCuXname("3", strCuA)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 8 Then
            '受讓人(讓與申請人)
            strTemp = GetCuXname("1", strCuB)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 9 Then
            '代表人
            strTemp = GetCuXname("2", strCuB)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 10 Then
            '地址
            strTemp = GetCuXname("3", strCuB)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
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
            .Selection.Delete
            If intA >= 4 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = True
            End If

            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            If intA < 4 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = False
            End If
         End If
      Next intA
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2023/09/22 台灣案讓與契約書(多案)
Private Sub runWordProc4_2()
Dim intA As Integer
Dim strCuA As String '讓與人
Dim strCuB As String '受讓人(讓與申請人)
Dim strText As String
Dim intMax As Single
Dim tmpArr As Variant

On Error GoTo ErrHand

   Call GetXcuList(strCuA, strCuB)
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.add
   g_WordAp.Selection.PageSetup.LeftMargin = g_WordAp.CentimetersToPoints(3)
   g_WordAp.Selection.PageSetup.RightMargin = g_WordAp.CentimetersToPoints(3)
   g_WordAp.Selection.PageSetup.TopMargin = g_WordAp.CentimetersToPoints(3)
   g_WordAp.Selection.PageSetup.BottomMargin = g_WordAp.CentimetersToPoints(2.5)
   g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
   g_WordAp.Selection.ParagraphFormat.DisableLineHeightGrid = True
   
   g_WordAp.Selection.Font.Name = "標楷體"
   g_WordAp.Selection.PageSetup.Orientation = wdOrientPortrait
   g_WordAp.Selection.Orientation = wdTextOrientationHorizontal
   g_WordAp.Selection.Font.Size = 24

   '固定行高
   g_WordAp.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
   g_WordAp.Selection.ParagraphFormat.LineSpacing = 15
   g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
   g_WordAp.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceAtLeast '標題改最小行高,否則畫面顯示會不完整
   g_WordAp.Selection.TypeText "讓與契約書"
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
   g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '文字靠左
   g_WordAp.Selection.Font.Size = 14
   g_WordAp.Selection.Text = "　　茲將下列專利權讓與"
   strTemp = Replace(GetCuXname("1", strCuB), "|", "、")
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.Text = strTemp
   g_WordAp.Selection.Font.Underline = True
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.Text = "。"
   g_WordAp.Selection.Font.Underline = False
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   
   '抓所有案號
   strTemp = pa(1) & pa(2) & pa(3) & pa(4)
   For intA = 1 To List1.ListCount
      strTemp = strTemp & "," & List1.List(intA - 1)
   Next intA
   tmpArr = Empty
   tmpArr = Split(strTemp, ",")
   g_WordAp.Selection.ParagraphFormat.LineSpacing = 20
   '插入表格
   g_WordAp.Selection.Tables.add Range:=g_WordAp.Selection.Range, NumRows:=1, NumColumns:=4
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders.Shadow = False
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Cells.VerticalAlignment = wdAlignVerticalTop
   g_WordAp.Selection.Paragraphs.Alignment = wdAlignParagraphLeft '文字靠左
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(1), RulerStyle:=wdAdjustProportional
   g_WordAp.Selection.Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
   g_WordAp.Selection.Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Collapse Direction:=wdCollapseStart
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   g_WordAp.Selection.Text = "申請案號"
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   g_WordAp.Selection.Text = "證書號"
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   g_WordAp.Selection.Text = "專　利　名　稱"
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   For intA = 0 To UBound(tmpArr)
      strExc(0) = Trim(tmpArr(intA))
      Call ChgCaseNo(strExc(0), strExc)
      strSql = "select pa11,pa22,nvl(pa05,nvl(pa06,pa07)) pname from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         g_WordAp.Selection.Text = intA + 1 & "."
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
         g_WordAp.Selection.Text = "" & RsTemp.Fields("pa11") '申請案號
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
         g_WordAp.Selection.Text = "" & RsTemp.Fields("pa22") '證書號
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
         g_WordAp.Selection.Text = "" & RsTemp.Fields("pname") '專利名稱
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
      End If
   Next intA
   g_WordAp.Selection.Rows.Delete '刪除空白列
   g_WordAp.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
   '固定行高
   g_WordAp.Selection.ParagraphFormat.LineSpacing = 26
   g_WordAp.Selection.TypeParagraph
   
   intMax = 48
   '讓與人、地址
   tmpArr = Empty
   tmpArr = Split(strCuA, ",")
   For intA = 0 To UBound(tmpArr)
      If Trim(tmpArr(intA)) <> "" Then
         g_WordAp.Selection.Text = "　　讓與人" & intA + 1 & "："
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         strTemp = GetCuXname("1", Trim(tmpArr(intA)))
         If GetTextLength(strTemp) < intMax Then
            strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
         Else
            strText = Trim(strTemp)
         End If
         g_WordAp.Selection.Text = strText
         g_WordAp.Selection.Font.Underline = True
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         g_WordAp.Selection.Font.Underline = False
         g_WordAp.Selection.TypeParagraph
         'Added by Lydia 2024/02/26 增加代表人
         strTemp = GetCuXname("2", Trim(tmpArr(intA)))
         If strTemp <> "" Then
            g_WordAp.Selection.Text = "　　代表人" & intA + 1 & "："
            g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
            If GetTextLength(strTemp) < intMax Then
               strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
            Else
               strText = Trim(strTemp)
            End If
            g_WordAp.Selection.Text = strText
            g_WordAp.Selection.Font.Underline = True
            g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
            g_WordAp.Selection.Font.Underline = False
            g_WordAp.Selection.TypeParagraph
         End If
         'end 2024/02/26
         
         g_WordAp.Selection.Text = "　　地　址" & intA + 1 & "："
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         strTemp = GetCuXname("3", Trim(tmpArr(intA)))
         If GetTextLength(strTemp) < intMax Then
            strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
         Else
            strText = Trim(strTemp)
         End If
         g_WordAp.Selection.Text = strText
         g_WordAp.Selection.Font.Underline = True
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         g_WordAp.Selection.Font.Underline = False
         g_WordAp.Selection.TypeParagraph
      End If
   Next intA
   g_WordAp.Selection.TypeParagraph
   
   '受讓人、地址
   tmpArr = Empty
   tmpArr = Split(strCuB, ",")
   For intA = 0 To UBound(tmpArr)
      If Trim(tmpArr(intA)) <> "" Then
         g_WordAp.Selection.Text = "　　受讓人" & intA + 1 & "："
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         strTemp = GetCuXname("1", Trim(tmpArr(intA)))
         If GetTextLength(strTemp) < intMax Then
            strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
         Else
            strText = Trim(strTemp)
         End If
         g_WordAp.Selection.Text = strText
         g_WordAp.Selection.Font.Underline = True
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         g_WordAp.Selection.Font.Underline = False
         g_WordAp.Selection.TypeParagraph
         'Added by Lydia 2024/02/26 增加代表人
         strTemp = GetCuXname("2", Trim(tmpArr(intA)))
         If strTemp <> "" Then
            g_WordAp.Selection.Text = "　　代表人" & intA + 1 & "："
            g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
            If GetTextLength(strTemp) < intMax Then
               strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
            Else
               strText = Trim(strTemp)
            End If
            g_WordAp.Selection.Text = strText
            g_WordAp.Selection.Font.Underline = True
            g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
            g_WordAp.Selection.Font.Underline = False
            g_WordAp.Selection.TypeParagraph
         End If
         'end 2024/02/26
         
         g_WordAp.Selection.Text = "　　地　址" & intA + 1 & "："
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         strTemp = GetCuXname("3", Trim(tmpArr(intA)))
         If GetTextLength(strTemp) < intMax Then
            strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
         Else
            strText = Trim(strTemp)
         End If
         g_WordAp.Selection.Text = strText
         g_WordAp.Selection.Font.Underline = True
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         g_WordAp.Selection.Font.Underline = False
         g_WordAp.Selection.TypeParagraph
      End If
   Next intA
   
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.Text = "　　　　中　華　民　國　　　　年　　　　月　　　　日"
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
      
End Sub

'Added by Lydia 2023/09/22 大陸案讓與契約書：下載Word範本套印
Private Sub runWordProc4_3()
Dim strName As String
Dim strText
Dim intA As Integer
Dim strCuA As String '讓與人
Dim strCuB As String '受讓人(讓與申請人)
Dim intMax As Single
Dim tmpArr As Variant

On Error GoTo ErrHand
       
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000401-1-04 內專P大陸案讓與契約書.docx", "M51", "000401", "1", "04", "4", "1")

   '下載範本檔
   'Modified by Lydia 2023/10/03 改副檔名為英文
   'm_FileName = "$$" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "讓與契約書.docx"
   m_FileName = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & ".ASSIGNMENT.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000401-1-04", , m_DefPath) = False Then
      Exit Sub
   End If

   Call GetXcuList(strCuA, strCuB)
   
   intMax = 57
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 9
         strName = "PS" & Format(intA, "00")
         strText = ""
         If intA = 1 Then
            '申請號
            strText = IIf(Trim(pa(11)) = "", String(10, " "), pa(11))
         ElseIf intA = 2 Then
            '名稱
            If Combo1.Text <> "" Then
               strText = Combo1.Text
            Else
               For intI = 5 To 7
                  If pa(intI) <> "" Then
                     strText = pa(intI)
                     Exit For
                  End If
               Next intI
            End If
         ElseIf intA = 3 Then
            '受讓人
            strText = Replace(GetCuXname("1", strCuB), "|", "、")
         ElseIf intA = 4 Then
            '讓與人
            strTemp = Replace(GetCuXname("1", strCuA), "|", "、")
            strText = "|" & strText
            strExc(1) = strTemp: strExc(2) = ""
            Do While strExc(1) <> ""
               strExc(2) = PUB_StrToStr(strExc(1), 48, True)
               strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
               strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
            Loop
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 5 Then
            '代表人
            strTemp = Replace(GetCuXname("2", strCuA), "|", "、")
            If GetTextLength(strTemp) < 20 Then
               strText = PUB_StrToStr(IIf(strTemp = "", " ", strTemp), 20, True)
            Else
               strText = strTemp
            End If
         ElseIf intA = 6 Then
            '地址
            strTemp = GetCuXname("3", strCuA)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 7 Then
            '受讓人
            strTemp = Replace(GetCuXname("1", strCuB), "|", "、")
            strText = "|" & strText
            strExc(1) = strTemp: strExc(2) = ""
            Do While strExc(1) <> ""
               strExc(2) = PUB_StrToStr(strExc(1), 48, True)
               strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
               strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
            Loop
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 8 Then
            '代表人
            strTemp = Replace(GetCuXname("2", strCuB), "|", "、")
            If GetTextLength(strTemp) < 20 Then
               strText = PUB_StrToStr(IIf(strTemp = "", " ", strTemp), 20, True)
            Else
               strText = strTemp
            End If
         ElseIf intA = 9 Then
            '地址
            strTemp = GetCuXname("3", strCuB)
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
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
            .Selection.Delete
            
            .Selection.Font.Underline = True '性    質：其他 設定 底線
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            .Selection.Font.Underline = False '性    質：其他 設定 底線
         End If
      Next intA

   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2023/09/22 大陸案解聘書：下載Word範本套印
Private Sub runWordProc5_1()
Dim strName As String
Dim strText
Dim intA As Integer
Dim tmpArr As Variant
Dim strKind As String
Dim intMax As Single

On Error GoTo ErrHand
       
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000401-0-05 內專P台灣案簽章切結書_個案.docx", "M51", "000401", "0", "05", "4", "1")

   '下載範本檔
   'Modified by Lydia 2023/10/03 改副檔名為英文
   'm_FileName = "$$" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "簽章切結書.docx"
   m_FileName = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & ".ATT.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000401-0-05", , m_DefPath) = False Then
      Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      'Modified by Lydia 2024/03/25 +第2次出現PS00
      'For intA = 1 To 6
      For intA = 0 To 6
         strName = "PS" & Format(intA, "00")
         strText = ""
         'Modified by Lydia 2024/03/25 +第2次出現PS00
         If intA = 1 Or intA = 0 Then
            strSql = "select cu15 from customer where cu01='" & Mid(ChangeCustomerL(pa(26)), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(pa(26)), 9, 1) & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strKind = "" & RsTemp.Fields("cu15")
            End If
            If strKind = "1" Then
               strText = "本公司"
            ElseIf strKind = "2" Then
               strText = "本校"
            ElseIf strKind = "3" Then
               strText = "本機構"
            Else
               strText = "本人"
            End If
            'Added by Lydia 2024/04/11
            If intA = 0 Then
               strTemp = ""
            Else
            'end 2024/04/11
               '專利種類
               Call ClsPDGetPatentTrademarkKind("1", pa(8), strTemp, IIf(pa(9) <> "000", True, False), pa(9))
            End If 'Added by Lydia 2024/04/11
            strText = strText & strTemp
         ElseIf intA = 2 Then
            '申請案號
            strText = pa(11)
         ElseIf intA = 3 Then
            '名稱
            If Combo1.Text <> "" Then
               strText = Combo1.Text
            Else
               For intI = 5 To 7
                  If pa(intI) <> "" Then
                     strText = pa(intI)
                     Exit For
                  End If
               Next intI
            End If
         ElseIf intA = 4 Then
            '切結人
            strTemp = GetCuXname("1", pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
            intMax = 32
            tmpArr = Empty
            tmpArr = Split(strTemp, "|")
            strExc(1) = "": strExc(2) = ""
            strText = "|" & strText
            For intI = 0 To UBound(tmpArr)
               strExc(1) = Trim(tmpArr(intI))
               Do While strExc(1) <> ""
                  strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                  strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                  strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
               Loop
            Next intI
            strText = Mid(strText, 2)
            strText = Replace(strText, "|", vbCrLf)
         ElseIf intA = 5 Then
            '非個人才列出代表人
            If strKind <> "0" Then
               strText = "代表人："
            End If
         ElseIf intA = 6 Then
            '非個人才列出代表人
            If strKind <> "0" Then
               strTemp = GetCuXname("2", pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
               tmpArr = Empty
               tmpArr = Split(strTemp, "|")
               strExc(1) = "": strExc(2) = ""
               strText = "|" & strText
               For intI = 0 To UBound(tmpArr)
                  strExc(1) = Trim(tmpArr(intI))
                  Do While strExc(1) <> ""
                     strExc(2) = PUB_StrToStr(strExc(1), intMax, True)
                     strText = strText & IIf(strText <> "|", "|", "") & strExc(2)
                     strExc(1) = Trim(Replace(strExc(1), Trim(strExc(2)), ""))
                  Loop
               Next intI
               strText = Mid(strText, 2)
               strText = Replace(strText, "|", vbCrLf)
            End If
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
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
            .Selection.Delete
            If intA = 4 Or intA = 6 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = True
            End If

            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            If intA = 4 Or intA = 6 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = False
            End If
         End If
      Next intA
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2023/09/22 台灣案簽章切結書(多案)
Private Sub runWordProc5_2()
Dim intA As Integer
Dim strText As String
Dim tmpArr As Variant
Dim strKind As String
Dim intMax As Single

On Error GoTo ErrHand

   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.add
   g_WordAp.Selection.PageSetup.LeftMargin = g_WordAp.CentimetersToPoints(3.17)
   g_WordAp.Selection.PageSetup.RightMargin = g_WordAp.CentimetersToPoints(3.54)
   g_WordAp.Selection.PageSetup.TopMargin = g_WordAp.CentimetersToPoints(2.54)
   g_WordAp.Selection.PageSetup.BottomMargin = g_WordAp.CentimetersToPoints(2.54)
   g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
   g_WordAp.Selection.ParagraphFormat.DisableLineHeightGrid = True
   
   g_WordAp.Selection.Font.Name = "標楷體"
   g_WordAp.Selection.PageSetup.Orientation = wdOrientPortrait
   g_WordAp.Selection.Orientation = wdTextOrientationHorizontal
   g_WordAp.Selection.Font.Size = 16

   g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
   g_WordAp.Selection.TypeText "切　　結　　書"
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '文字靠左
   strSql = "select cu15 from customer where cu01='" & Mid(ChangeCustomerL(pa(26)), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(pa(26)), 9, 1) & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strKind = "" & RsTemp.Fields("cu15")
   End If
   If strKind = "1" Then
      strText = "本公司"
   ElseIf strKind = "2" Then
      strText = "本校"
   ElseIf strKind = "3" Then
      strText = "本機構"
   Else
      strText = "本人"
   End If
   g_WordAp.Selection.Text = "　　茲切結擔保" & strText & "於下列專利案件讓與契約書上之簽章確為" & strText & "之簽章，特此聲明。"
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   
   '抓所有案號
   strTemp = pa(1) & pa(2) & pa(3) & pa(4)
   For intA = 1 To List1.ListCount
      strTemp = strTemp & "," & List1.List(intA - 1)
   Next intA
   tmpArr = Empty
   tmpArr = Split(strTemp, ",")
   '插入表格
   g_WordAp.Selection.Tables.add Range:=g_WordAp.Selection.Range, NumRows:=1, NumColumns:=4
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
   g_WordAp.Selection.Borders.Shadow = False
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Cells.VerticalAlignment = wdAlignVerticalTop
   g_WordAp.Selection.Paragraphs.Alignment = wdAlignParagraphLeft '文字靠左
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(1), RulerStyle:=wdAdjustProportional
   g_WordAp.Selection.Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
   g_WordAp.Selection.Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
   g_WordAp.Selection.SelectRow
   g_WordAp.Selection.Font.Size = 14
   g_WordAp.Selection.Collapse Direction:=wdCollapseStart
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   g_WordAp.Selection.Text = "申請案號"
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   g_WordAp.Selection.Text = "證書號"
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   g_WordAp.Selection.Text = "專　利　名　稱"
   g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
   For intA = 0 To UBound(tmpArr)
      strExc(0) = Trim(tmpArr(intA))
      Call ChgCaseNo(strExc(0), strExc)
      strSql = "select pa11,pa22,nvl(pa05,nvl(pa06,pa07)) pname from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         g_WordAp.Selection.Text = intA + 1 & "."
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
         g_WordAp.Selection.Text = "" & RsTemp.Fields("pa11") '申請案號
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
         g_WordAp.Selection.Text = "" & RsTemp.Fields("pa22") '證書號
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
         g_WordAp.Selection.Text = "" & RsTemp.Fields("pname") '專利名稱
         g_WordAp.Selection.MoveRight Unit:=wdCell, Count:=1
      End If
   Next intA
   g_WordAp.Selection.Rows.Delete '刪除空白列
   g_WordAp.Selection.Font.Size = 16
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.Text = "　　此致"
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.Text = "經濟部智慧財產局"
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   
   intMax = 32
   tmpArr = Empty
   tmpArr = Split(pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), ",")
   For intA = 0 To UBound(tmpArr)
      If Trim(tmpArr(intA)) <> "" Then
         g_WordAp.Selection.Text = "　　　　　切結人" & IIf(pa(27) <> "", intA + 1, "") & "："
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         strTemp = GetCuXname("1", Trim(tmpArr(intA)))
         If GetTextLength(strTemp) < intMax Then
            strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
         Else
            strText = Trim(strTemp)
         End If
         g_WordAp.Selection.Text = strText
         g_WordAp.Selection.Font.Underline = True
         g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
         g_WordAp.Selection.Font.Underline = False
         g_WordAp.Selection.TypeParagraph
         
         strSql = "select cu15 from customer where cu01='" & Mid(ChangeCustomerL(Trim(tmpArr(intA))), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(Trim(tmpArr(intA))), 9, 1) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If "" & RsTemp.Fields("cu15") <> "0" Then
               g_WordAp.Selection.Text = "　　　　　代表人" & IIf(pa(27) <> "", intA + 1, "") & "："
               g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
               strTemp = GetCuXname("2", Trim(tmpArr(intA)))
               If GetTextLength(strTemp) < intMax Then
                  strText = PUB_StrToStr(IIf(Trim(strTemp) = "", " ", Trim(strTemp)), intMax, True)
               Else
                  strText = Trim(strTemp)
               End If
               g_WordAp.Selection.Text = strText
               g_WordAp.Selection.Font.Underline = True
               g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
               g_WordAp.Selection.Font.Underline = False
               g_WordAp.Selection.TypeParagraph
            End If
         End If
      End If
   Next intA
   
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Selection.Text = "中華民國　　　　年　　　　月　　　　日"
   g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
   g_WordAp.Selection.TypeParagraph
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
      
End Sub

'Added by Lydia 2024/03/26 PCT案委托書：下載Word範本套印
Private Sub runWordProc2_1()
Dim strName As String
Dim strText
Dim strTmpA As String
Dim intA As Integer
Dim tmpArr As Variant

On Error GoTo ErrHand
       
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000401-1-02 內專PCT案委托書.docx", "M51", "000401", "1", "02", "4", "1")

   '下載範本檔
   m_FileName = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & ".POA.docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000401-1-02", , m_DefPath) = False Then
      Exit Sub
   End If
   
   '讓與案(701,708)的委任書之委任人，非新案改抓進度的受讓人，新案才抓申請人=委任人
   Dim strCuA As String '讓與人
   Dim strCuB As String '受讓人(讓與申請人)
   Dim strCuNow As String '要帶入的委任人
   If InStr("701,708", m_CP10) > 0 And m_CP31 <> "Y" Then
       Call GetXcuList(strCuA, strCuB)
       strCuNow = strCuB
   Else
       strCuNow = pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30)
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 3
         strName = "PS" & Format(intA, "00")
         strText = ""
         If intA = 1 Then
            '名稱
            If Combo1.Text <> "" Then
               strText = Combo1.Text
            Else
               For intI = 5 To 7
                  If pa(intI) <> "" Then
                     strText = pa(intI)
                     Exit For
                  End If
               Next intI
            End If
            strText = String(2, " ") & strText & String(2, " ")
         ElseIf intA = 2 Then
            '委託人(申請人)
            strTmpA = Replace(GetCuXname("1", strCuNow), "|", vbCrLf)
            If Len(strTmpA) < 43 Then
               strTmpA = PUB_StrToStr(IIf(strTmpA = "", " ", strTmpA), 43, True)
            End If
            strText = strTmpA
         ElseIf intA = 3 Then
            '代表人
            strTmpA = Replace(GetCuXname("2", strCuNow), "|", vbCrLf)
            If Len(strTmpA) < 36 Then
               strTmpA = PUB_StrToStr(IIf(strTmpA = "", " ", strTmpA), 36, True)
            End If
            strText = strTmpA
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
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
            .Selection.Delete
            If intA <= 3 Then
               '設定底線
               .Selection.Font.Underline = True
            End If

            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            If intA <= 3 Then
               .Selection.Font.Underline = False
            End If
         End If
      Next intA
   End With
   
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   g_WordAp.Activate
   
   Set g_WordAp = Nothing
          
   Exit Sub
   
ErrHand:
   If Err.NUMBER = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.NUMBER <> 0 Then
      MsgBox Err.NUMBER & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub
