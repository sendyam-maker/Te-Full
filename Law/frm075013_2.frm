VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075013_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "出庭費查詢"
   ClientHeight    =   5484
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   9432
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "不領取確認"
      Height          =   400
      Index           =   3
      Left            =   6000
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   884
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認明細(&E)"
      Height          =   400
      Index           =   0
      Left            =   4853
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   884
      Width           =   1140
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Index           =   2
      Left            =   2472
      MaxLength       =   7
      TabIndex        =   14
      Top             =   884
      Width           =   1092
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Index           =   1
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   13
      Top             =   884
      Width           =   1092
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Index           =   0
      Left            =   1224
      MaxLength       =   1
      TabIndex        =   12
      Top             =   528
      Width           =   372
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔"
      Height          =   400
      Left            =   4853
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   72
      Width           =   1140
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8352
      TabIndex        =   7
      Top             =   72
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4032
      TabIndex        =   6
      Top             =   72
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   1
      Left            =   6014
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   72
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   2
      Left            =   7535
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   72
      Width           =   795
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   4092
      Left            =   72
      TabIndex        =   3
      Top             =   1320
      Width           =   9228
      _ExtentX        =   16277
      _ExtentY        =   7218
      _Version        =   393216
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtUsernum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   0
      Top             =   130
      Width           =   744
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y：已發放  N：未發放  K：確認不領取  空白：不限制)"
      Height          =   180
      Index           =   2
      Left            =   1728
      TabIndex        =   17
      Top             =   576
      Width           =   4368
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2250
      X2              =   2400
      Y1              =   1032
      Y2              =   1032
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "確認期間："
      Height          =   180
      Left            =   144
      TabIndex        =   11
      Top             =   936
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否已發放："
      Height          =   180
      Index           =   1
      Left            =   144
      TabIndex        =   10
      Top             =   576
      Width           =   1080
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   348
      Left            =   2112
      TabIndex        =   9
      Top             =   130
      Width           =   2124
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3746;614"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   144
      TabIndex        =   2
      Top             =   182
      Width           =   900
   End
   Begin MSForms.Label lblUserName 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   132
      Width           =   1476
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "2603;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm075013_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/09/30 (113/11/01上線)
Option Explicit
Dim intLastRow As Integer '記錄MGrid1勾選最後一筆
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Public cmdState As Integer
Dim mStatus As String  'R=律師個人, C=財務+電腦中心
Dim nFrm100101_6 As Form '顧問基本檔
Dim nFrm100101_5 As Form '法務基本檔
Dim nFrm100101_2 As Form '案件進度
Dim bolActFrm100 As Boolean '是否開啟共同查詢
Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
Const cntFixed As Integer = 5
Dim colCL01 As Integer, colCaseNo As Integer, colCL02 As Integer
Dim colCL06 As Integer 'Added by Lydia 2025/04/07

Private Sub cmdExit_Click()
   '關閉共同查詢的畫面
   If bolActFrm100 = True Then
      fnCloseAllFrm100
   End If
   
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   If cmdState = 0 Then
      If PUB_CheckFormExist("frm075013_1") Then
         MsgBox "請先關閉〔出庭費確認維護明細〕畫面！"
         Exit Sub
      End If
   End If
   
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
Dim intA As Integer, StrTag As String, intB As Integer
Dim Str01 As String, strKeyNo As String

On Error GoTo ErrorHandler
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For intA = 1 To MGrid1.Rows - 1
      MGrid1.col = 0
      MGrid1.row = intA
      If Trim(MGrid1.Text) = "V" Then
         bolRefresh = False
         MGrid1.col = 0
         MGrid1.Text = ""
         For intB = 0 To MGrid1.Cols - 1
            MGrid1.col = intB
            MGrid1.CellBackColor = &H80000005
         Next
                  
         StrTag = MGrid1.TextMatrix(intA, colCaseNo)
         '增加對執行”基本資料”和”案件進度查詢”的限閱案件控制
         If cmdState = 1 Or cmdState = 2 Then
            If PUB_ChkCufaByCaseNo(strUserNum, Me.Name, Replace(StrTag, "-", ""), "1") = False Then
               GoTo EXITSUB
            End If
            bolActFrm100 = True
         End If
         Str01 = SystemNumber(StrTag, 1)
         If cmdState = 1 Or cmdState = 2 Then
            If fnSaveParentForm(Me) = False Then
               GoTo EXITSUB
            End If
         End If
         strKeyNo = MGrid1.TextMatrix(intA, colCL01)
         
         Me.Show
         Select Case cmdState
            Case 0 '確認明細
               If Len(strKeyNo) = 9 Then
                  If Left(MGrid1.TextMatrix(intA, colCL02), 1) >= "6" And Left(MGrid1.TextMatrix(intA, colCL02), 1) < "F" Then
                     Call frm075013_1.SetParent(Me, strKeyNo, MGrid1.TextMatrix(intA, colCL02))
                  Else
                     MsgBox "查無資料！", vbExclamation
                     GoTo EXITSUB
                  End If
                  frm075013_1.Show
                  Me.Hide
               End If
            Case 1 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "LA" '顧問
                      Screen.MousePointer = vbHourglass
                      nFrm100101_6.Show
                      nFrm100101_6.Tag = StrTag
                      nFrm100101_6.StrMenu
                      Screen.MousePointer = vbDefault
                  Case Else  '法務
                      Screen.MousePointer = vbHourglass
                      nFrm100101_5.Show
                      nFrm100101_5.Tag = StrTag
                      nFrm100101_5.StrMenu
                      Screen.MousePointer = vbDefault
               End Select
               
            Case 2 '案件進度
               nFrm100101_2.Show
               nFrm100101_2.Tag = StrTag
               nFrm100101_2.StrMenu
            'Added by Lydia 2025/04/07
            Case 3 '不領取確認
               If "" & MGrid1.TextMatrix(intA, colCL01) <> "" And "" & MGrid1.TextMatrix(intA, colCL02) <> "" Then
                  If "" & MGrid1.TextMatrix(intA, colCL06) <> "" Then
                     Screen.MousePointer = vbDefault
                     'Memo by Lydia 2025/04/17 改用輸入txtDB(0)
                     MsgBox "" & MGrid1.TextMatrix(intA, colCaseNo) & "已有" & IIf(txtDB(0) = "K", "不領取確認", "發放日期"), vbExclamation
                     GoTo EXITSUB
                  Else
                     '判斷不領取且EMAIL通知次數已達4次的資料才能按
                     strExc(0) = "select decode(cl04,null,decode(cl05,null,null,'不領取'),'領取') as chktype,counting(cl07) cnt from caselawer where cl01='" & MGrid1.TextMatrix(intA, colCL01) & "' and cl02='" & MGrid1.TextMatrix(intA, colCL02) & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     'Modified by Lydia 2025/08/20 只要律師確認，就可以上財務確認
                     'If "" & RsTemp.Fields("chktype") = "不領取" And Val("" & RsTemp.Fields("cnt")) >= 4 Then
                     If "" & RsTemp.Fields("chktype") = "不領取" Then
                        strSql = "Update CaseLawer set cl09=to_char(sysdate,'yyyymmdd') where cl09 is null and cl01='" & MGrid1.TextMatrix(intA, colCL01) & "' and cl02='" & MGrid1.TextMatrix(intA, colCL02) & "'"
                        cnnConnection.Execute strSql
                        bolRefresh = True
                     Else
                        Screen.MousePointer = vbDefault
                        'Modified by Lydia 2025/08/20 只要律師確認，就可以上財務確認
                        'MsgBox "" & MGrid1.TextMatrix(intA, colCaseNo) & "確認結果：" & IIf("" & RsTemp.Fields("chktype") = "", "(未確認)", "" & RsTemp.Fields("chktype")) & "，Email通知次數：" & RsTemp.Fields("cnt") & vbCrLf & _
                                "必需為律師不領取並且EMAIL通知次數已達4次，才能執行不領取確認！", vbExclamation
                        MsgBox "" & MGrid1.TextMatrix(intA, colCaseNo) & "確認結果：" & IIf("" & RsTemp.Fields("chktype") = "", "(未確認)", "" & RsTemp.Fields("chktype")) & vbCrLf & _
                                "必需為律師確認不領取次，才能執行財務處不領取確認！", vbExclamation
                        GoTo EXITSUB
                     End If
                  End If
               End If
            'end 2025/04/07
         End Select
         'Modified by Lydia 2025/04/07
         'Exit For
         GoTo EXITSUB
      End If
   Next intA
   
   If bolRefresh = True Then
      cmdQuery_Click
   End If
   
   Exit Sub
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
EXITSUB:
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   'Added by Lydia 2025/04/07
   If bolRefresh = True Then
      cmdQuery_Click
   End If
   'end 2025/04/07
End Sub

Private Sub cmdQuery_Click()
   Call doQuery(True)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Form_Load()
Dim oObj As Object

   MoveFormToCenter Me
   Set nFrm100101_6 = Forms(0).GetForm("frm100101_6")
   If Not nFrm100101_6 Is Nothing Then
      cmdOK(1).Visible = True
   Else
      cmdOK(1).Visible = False
   End If
   Set nFrm100101_5 = Forms(0).GetForm("frm100101_5")
   
   Set nFrm100101_2 = Forms(0).GetForm("frm100101_2")
   If Not nFrm100101_2 Is Nothing Then
      cmdOK(2).Visible = True
   Else
      cmdOK(2).Visible = False
   End If
   
   For Each oObj In txtDB
      oObj.Text = ""
   Next
   Combo1.Clear

   mStatus = "R"
   
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
      mStatus = "C"
      txtUsernum.Visible = False: lblUserName.Visible = False
      Combo1.Visible = True
      Combo1.Left = txtUsernum.Left
      strQ1 = "select st01,st02 from staff where st03 in ('L01','L00') and st04='1' order by 1"
      intQ = 1
      Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         rsQD.MoveFirst
         Do While Not rsQD.EOF
            Combo1.AddItem rsQD.Fields("st01") & " " & rsQD.Fields("st02")
            rsQD.MoveNext
         Loop
      End If
      'P/T案
      Combo1.AddItem PUB_StrToStr("P", 4, True) & "內專P案"
      Combo1.AddItem PUB_StrToStr("T", 4, True) & "內商T案"
      Combo1.AddItem PUB_StrToStr("FCP", 4, True) & "外專FCP案"
      Combo1.AddItem PUB_StrToStr("FCT", 4, True) & "外商FCT案"
      Combo1.Text = ""
      cmdExcel.Visible = True
      cmdOK(3).Visible = True 'Added by Lydia 2025/04/07
      Label4(2).Caption = "(Y：已發放  N：未發放  K：確認不領取  空白：不限制)" 'Added by Lydia 2025/04/17 不用CheckBox，改用輸入選擇「(財務)確認不領取 」
   Else
      txtUsernum.Visible = True: lblUserName.Visible = True
      Combo1.Visible = False
      txtUsernum = strUserNum
      cmdExcel.Visible = False
      cmdOK(3).Visible = False  'Added by Lydia 2025/04/07
      Label4(2).Caption = "(Y：已發放  N：未發放  空白：不限制)" 'Added by Lydia 2025/04/17 不用CheckBox，改用輸入選擇「(財務)確認不領取 」
   End If
   Call SetDefDate(True)  '預設查詢：本月確認
   If nFrm100101_6 Is Nothing And mStatus = "C" Then  '財務系統
      cmdOK(0).Left = cmdOK(1).Left
      cmdOK(0).Top = cmdExcel.Top
   ElseIf Not nFrm100101_6 Is Nothing And mStatus = "C" Then   '非財務系統
      '不移動
   Else
      cmdOK(0).Top = cmdExcel.Top '法務系統:律師
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MenuEnabled
   Set rsQD = Nothing
   Set frm075013_2 = Nothing
End Sub

Private Sub MGrid1_Click()

   If "" & MGrid1.TextMatrix(MGrid1.row, 1) <> "" Then
      GridClick MGrid1, intLastRow, 0, 0, , "V"
   End If
   
End Sub

Private Sub MGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow MGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGrid1.col = nCol
   MGrid1.row = nRow
   If Me.MGrid1.row < 1 And Me.MGrid1.Text <> "V" Then
      If InStr("個人出庭費金額,出庭費總額", Me.MGrid1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 0 Then
      'Modified by Lydia 2025/04/17 +K
      If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
         'Added by Lydia 2025/04/17 K=確認不領取
         If KeyAscii = 75 And InStr(Label4(2), Chr(KeyAscii)) > 0 Then
         Else
         'end 2025/04/17
            KeyAscii = 0
            Beep
         End If
      End If
   End If
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0  '是否已發放

      Case 1, 2 '確認期間
         If txtDB(Index) <> "" Then
            If CheckIsTaiwanDate(txtDB(Index)) = True Then
               'If txtDB(0) = "" Then txtDB(0) = "Y"
            Else
               Cancel = True
            End If
         End If
   End Select
   
   If Cancel Then TextInverse txtDB(Index)
End Sub

Private Sub txtUsernum_GotFocus()
   TextInverse txtUsernum
End Sub

Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUsernum_Change()
   If Len(txtUsernum) >= 5 Then
      lblUserName = GetStaffName(txtUsernum, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub SetGrd1(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   'Modified by Lydia 2025/04/07 只顯示不領取確認 'Memo by Lydia 2025/04/17 改用輸入txtDB(0)
   'arrGridHeadText = Array("V", "律所案號", "智慧所案號", "承辦律師", "出庭費", "確認日期", "確認結果", "發放日期", "發文日", "收款日", "X01", "其他出庭律師", "出庭費總額", "Y01", "CL01")
   arrGridHeadText = Array("V", "律所案號", "智慧所案號", "承辦律師", "出庭費", "確認日期", "確認結果", IIf(txtDB(0) = "K", "財務確認不領取日期", "發放日期"), "發文日", "收款日", "X01", "其他出庭律師", "出庭費總額", "Y01", "CL01")
   If mStatus = "R" Then '律師個人隱藏:其他出庭律師,出庭費總額
      arrGridHeadWidth = Array(300, 1500, 1500, 1400, 1000, 1000, 1000, 1000, 1000, 1000, 0, 0, 0, 0, 0)
   Else
      'Modified by Lydia 2025/04/07 只顯示不領取確認 'Memo by Lydia 2025/04/17 改用輸入txtDB(0)
      'arrGridHeadWidth = Array(300, 1500, 1500, 1400, 1000, 1000, 1000, 1000, 1000, 1000, 0, 1300, 1200, 0, 0)
      arrGridHeadWidth = Array(300, 1500, 1500, 1400, 1000, 1000, IIf(txtDB(0) = "K", 1400, 1000), 1000, 1000, 1000, 0, 1300, 1200, 0, 0)
   End If
   
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      MGrid1.Clear
      MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
      MGrid1.row = 0
      MGrid1.col = iRow
      MGrid1.Text = arrGridHeadText(iRow)
      MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrid1.CellAlignment = flexAlignCenterCenter
   Next

   For intI = 1 To MGrid1.Rows - 1
      MGrid1.row = intI
      For iRow = 0 To MGrid1.Cols - 1
         MGrid1.col = iRow
         MGrid1.CellBackColor = &H80000005
         If InStr("04,12,", Format(iRow, "00")) > 0 Then '靠右
            MGrid1.CellAlignment = flexAlignRightCenter
         ElseIf InStr("06,", Format(iRow, "00")) > 0 Then  '置中
            MGrid1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   
   If colCL01 = 0 Then
      colCL01 = PUB_MGridGetId("CL01", MGrid1)
      colCaseNo = PUB_MGridGetId("律所案號", MGrid1)
      colCL02 = PUB_MGridGetId("Y01", MGrid1) '確認之出庭律師
      colCL06 = PUB_MGridGetId("發放日期", MGrid1) 'Added by Lydia 2025/04/07
   End If
   
   MGrid1.Visible = True
End Sub

Public Function doQuery(ByVal bolMsg As Boolean, Optional ByVal bolRefresh As Boolean = True, Optional ByVal bolRunExcel As Boolean = False) As Boolean
   
   doQuery = False
   
   Screen.MousePointer = vbHourglass
   ClearQueryLog (Me.Name)
   
   '員工編號
   If Combo1.Visible = True Or mStatus = "C" Then
      If Trim(Combo1.Text) = "" Then '空白=全部
         pub_QL05 = pub_QL05 & ";員工編號：ALL 全部"
      ElseIf InStr(",P,T,FCP,FCT,", "," & Left(Combo1, 1) & ",") = 0 Then
         pub_QL05 = pub_QL05 & ";員工編號：" & Trim(Left(Combo1, 6))
      ElseIf InStr(",P,T,FCP,FCT,", "," & Left(Combo1, 1) & ",") > 0 Then
         pub_QL05 = pub_QL05 & ";PT案：" & Trim(Left(Combo1, 4))
      End If
   Else
      pub_QL05 = pub_QL05 & ";員工編號：" & txtUsernum
   End If
   
   'Added by Lydia 2025/04/07 只顯示不領取確認 'Memo by Lydia 2025/04/17 改用輸入txtDB(0)
   If txtDB(0) = "K" Then
      pub_QL05 = pub_QL05 & ";只顯示不領取確認"
   Else
   'end 2025/04/07
      '是否已發放
      If Trim(txtDB(0)) = "Y" Then
         pub_QL05 = pub_QL05 & ";是否已發放：Y"
      ElseIf Trim(txtDB(0)) = "N" Then
         pub_QL05 = pub_QL05 & ";是否已發放：N"
      ElseIf Trim(txtDB(0)) = "" Then
         pub_QL05 = pub_QL05 & ";排除不領取確認"
      End If
   End If 'Added by Lydia 2025/04/07
   
   '確認期間
   If Trim(txtDB(1)) <> "" And Trim(txtDB(2)) <> "" And Trim(txtDB(1)) > Trim(txtDB(2)) Then
      MsgBox "確認期間起日不可大於迄日！", vbExclamation
      txtDB(1).SetFocus
      txtDB_GotFocus 1
      Exit Function
   End If
   If Trim(txtDB(1)) <> "" Or Trim(txtDB(2)) <> "" Then
      pub_QL05 = pub_QL05 & ";確認期間：" & txtDB(1) & "~" & txtDB(2)
   End If
   
   If bolRefresh = True Then
      Call SetGrd1(True) '清空
   End If
   If bolRunExcel = True Then
      pub_QL05 = pub_QL05 & ";執行產生Excel"
   End If
   
   'Modified by Lydia 2025/04/07 只顯示不領取確認 'Memo by Lydia 2025/04/17 改用輸入txtDB(0)
   strQ1 = PUB_GetFrm075013toSQL("1", IIf(Combo1.Visible = True Or mStatus = "C", Trim(Combo1.Text), txtUsernum), Trim(txtDB(0)), Trim(txtDB(1)), Trim(txtDB(2)))
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      InsertQueryLog (rsQD.RecordCount)
      If bolRefresh = True Then
         MGrid1.FixedCols = 0
         Set MGrid1.Recordset = rsQD
         Call SetGrd1
         MGrid1.FixedCols = cntFixed
      End If
      doQuery = True
   Else
      If bolMsg = True Then ShowNoData
      InsertQueryLog (0)
   End If
   
   Screen.MousePointer = vbDefault
End Function

'Mark by Lydia 2025/04/07 (保留)補CaseLawer
'Private Sub Command1_Click()
'Dim bolConn As Boolean
'
'   If MsgBox("是否執行補CaseLawer" & vbCrLf & "期間：2023/10/01~2024/04/30", vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
'      strExc(0) = "select cp09,cp14,sqldatet(cp158) as cp158t,cl01 From caseprogress, casepropertymap,caselawer " & _
'                  "where cp01=cpm01(+) and cp10=cpm02(+) and instr(',220113,',','||cpm12||',') > 0 and cp159=0 and cp158 >=20231001 and cp158<=20240430 " & _
'                  "and cp09=cl01(+) and cp14=cl02(+)"
'      intQ = 1
'      Set rsQD = ClsLawReadRstMsg(intQ, strExc(0))
'      If intQ = 1 Then
'On Error GoTo ErrHandle
'         cnnConnection.BeginTrans
'         bolConn = True
'         rsQD.MoveFirst
'         Do While Not rsQD.EOF
'            If "" & rsQD.Fields("cl01") = "" Then
'               strExc(1) = "select los15,los02 from caseprogress,lawofficesource where cp09='" & rsQD.Fields("cp09") & "' and cp162=los15(+) and los15 is not null "
'               strExc(2) = "15000" '現行未預設出庭費者，一律預設15000 ---frm071018
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'               If intI = 1 Then
'                  '比照PUB_UpdateTTFee
'                  If "" & RsTemp.Fields("los02") = "A4" Or "" & RsTemp.Fields("los02") = "B1" Then
'                         strExc(2) = 非B2律師費
'                  ElseIf "" & RsTemp.Fields("los02") = "B2" Then
'                      If PUB_IsB2NeedCourt("" & RsTemp.Fields("los15")) = True Then
'                         strExc(2) = B2律師費
'                      End If
'                  End If
'               Else '沒有案源>>純法律所
'               End If
'               If strExc(2) <> "" Then
'                  strSql = "Insert Into CaseLawer(CL01,CL02,CL03) Values ('" & rsQD.Fields("cp09") & "','" & rsQD.Fields("cp14") & "'," & Val(strExc(2)) & ") "
'                  cnnConnection.Execute strSql
'               End If
'            End If
'            'Email通知記錄(CL07)：記錄民國年月日xxx/xx/xx；用,區隔。
'            strSql = "Update CaseLawer set cl07='" & rsQD.Fields("cp158t") & "' where cl01='" & rsQD.Fields("cp09") & "' "
'            cnnConnection.Execute strSql
'            rsQD.MoveNext
'         Loop
'         cnnConnection.CommitTrans
'      End If
'      MsgBox "OK!"
'      bolConn = False
'   End If
'   Exit Sub
'
'ErrHandle:
'   If bolConn = True Then
'      cnnConnection.RollbackTrans
'   End If
'End Sub

Private Sub CmdExcel_Click()
Dim m_strFilePath As String
    
   m_strFilePath = strExcelPath & strSrvDate(2) & "_" & "出庭費清單" & MsgText(43)
   If Dir(m_strFilePath) <> "" Then
      If PUB_ChkFileOpening(m_strFilePath) = True Then
         Exit Sub
      End If
      Kill m_strFilePath
   End If
    
   If doQuery(True, True, True) = False Then
      '先更新畫面，並且記錄QueryLog
      Exit Sub
   End If
      
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   
   'Memo by Lydia 2025/04/17 改用輸入txtDB(0)
   Call PUB_GetFrm075013toXls(Me.Name, IIf(txtUsernum <> "", txtUsernum, Combo1.Text), txtDB(0), txtDB(1), txtDB(2), m_strFilePath, , True)
   
   If m_strFilePath = "" Then
      MsgBox "無資料可產生Excel ！", vbInformation
   Else
      MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN
   End If
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
End Sub

Private Sub SetDefDate(ByVal bolYes As Boolean)
   If bolYes = True Then
      '預設已確認期間
      txtDB(0) = "Y"
      If Val(Right(strSrvDate(1), 2)) < 16 Then
         txtDB(1) = TransDate(Left(CompDate(1, -2, strSrvDate(1)), 6) & "16", 1)
         txtDB(2) = TransDate(Left(CompDate(1, -1, strSrvDate(1)), 6) & "15", 1)
      Else
         txtDB(1) = TransDate(Left(CompDate(1, -1, strSrvDate(1)), 6) & "16", 1)
         txtDB(2) = TransDate(Left(strSrvDate(1), 6) & "15", 1)
      End If
   Else
      txtDB(0) = "": txtDB(1) = "": txtDB(2) = ""
   End If
End Sub

