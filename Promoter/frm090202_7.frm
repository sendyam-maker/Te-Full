VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090202_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "P非台灣案件新增歷程"
   ClientHeight    =   5748
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8964
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8964
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   0
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   3465
      MaxLength       =   2
      TabIndex        =   3
      Top             =   120
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   3075
      MaxLength       =   1
      TabIndex        =   2
      Top             =   120
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   1845
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   1212
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   4950
      TabIndex        =   4
      Top             =   30
      Width           =   855
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "新增歷程(&D)"
      Height          =   360
      Left            =   5880
      TabIndex        =   5
      Top             =   30
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7170
      TabIndex        =   6
      Top             =   30
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   4335
      Left            =   60
      TabIndex        =   7
      Top             =   1350
      Width           =   8835
      _ExtentX        =   15600
      _ExtentY        =   7641
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|收文日|總收文號|案件性質|相關收文號|承辦人|智權人員|本所期限|法定期限|進度備註"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      TabIndex        =   8
      Top             =   720
      Width           =   7665
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13520;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "分所號："
      Height          =   180
      Left            =   360
      TabIndex        =   19
      Top             =   1095
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   180
      TabIndex        =   18
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   180
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   2865
      TabIndex        =   16
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "審定證書號："
      Height          =   180
      Left            =   5760
      TabIndex        =   15
      Top             =   495
      Width           =   1080
   End
   Begin VB.Label Label3 
      Height          =   180
      Left            =   1110
      TabIndex        =   14
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label4 
      Height          =   180
      Left            =   3765
      TabIndex        =   13
      Top             =   480
      Width           =   1875
   End
   Begin VB.Label Label12 
      Height          =   180
      Left            =   6900
      TabIndex        =   12
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label Label15 
      Height          =   180
      Left            =   1170
      TabIndex        =   11
      Top             =   1095
      Width           =   2235
   End
   Begin MSForms.Label Label16 
      Height          =   280
      Left            =   4950
      TabIndex        =   10
      Top             =   1095
      Width           =   765
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "目前智權人員："
      Height          =   180
      Left            =   3630
      TabIndex        =   9
      Top             =   1095
      Width           =   1260
   End
End
Attribute VB_Name = "frm090202_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/23 Form2.0已修改
'Create by Sindy 2015/9/3
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim dblPrevRow As Double
Dim m_PrevForm As Form '前一畫面


'明細資料
Private Sub cmdDetail_Click()
Dim rsTmp As New ADODB.Recordset
Dim i As Integer, intMaxEEP02 As Integer
Dim m_EPMan As String
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
'         '檢查是否有聯絡以外的歷程，若有不可再此作業做新增判發歷程
'         strExc(0) = "Select eep01" & _
'                     " from EmpElectronProcess" & _
'                     " where eep01='" & GRD1.TextMatrix(i, 2) & "'" & _
'                     " and eep04 not in('" & EMP_聯絡 & "')"
'         intI = 1
'         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            MsgBox "此案已有其他歷程，不可在此作業直接新增判發歷程！", vbExclamation, "訊息"
'            txtCode(0).SetFocus
'            Exit Sub
'         End If
'         '檢查是否已經是判發送件
'         strExc(0) = "Select eep1.eep01" & _
'                     " from EmpElectronProcess eep1,CaseProgress" & _
'                     " where cp09='" & GRD1.TextMatrix(i, 2) & "' and eep1.eep01=cp09" & _
'                     " and eep1.eep04 in('" & EMP_判發 & "','" & EMP_退件重送 & "')" & _
'                     " And eep1.eep02=(select max(eep02) from EmpElectronProcess where eep01=eep1.eep01)"
'         intI = 1
'         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If MsgBox("此案已在判發送件階段，是否要進入歷程？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'               txtCode(0).SetFocus
'               Exit Sub
'            End If
'         Else
            '取得最大流水號
            strExc(0) = "Select nvl(max(eep02),0) from EmpElectronProcess where eep01='" & GRD1.TextMatrix(i, 2) & "'"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               intMaxEEP02 = rsTmp.Fields(0)
            End If
            intMaxEEP02 = intMaxEEP02 + 1
            '收受者
            'Modified by Morgan 2025/1/21 與Sindy確認收受者應可設為操作人。(P案程序人員工作將改依智權區域分配，不再固定工作)
            'm_EPMan = Pub_GetSpecMan("PS2") 'P非台灣案發文人員
            m_EPMan = strUserNum
            'end 2025/1/21
            strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07) values(" & _
                     CNULL(GRD1.TextMatrix(i, 2)) & "," & intMaxEEP02 & ",'" & strUserNum & "'," & _
                     CNULL(EMP_判發) & "," & CNULL(Left(Trim(m_EPMan), 5)) & "," & strSrvDate(1) & "," & _
                     Right("000000" & ServerTime, 6) & ")"
            cnnConnection.Execute strSql
'         End If
         
         'Add By Sindy 2017/1/9 為防止使用者明細畫面未關閉又從Menu再進入主作業畫面操作(因此下層明細先Unload再開啟)
         If TypeName(frm090202_4_1) <> "Nothing" Then
            Unload frm090202_4_1
         End If
         '2017/1/9 END
         frm090202_4_1.Hide
         frm090202_4_1.m_EEP01 = GRD1.TextMatrix(i, 2) '總收文號
         frm090202_4_1.m_AttEEP02 = intMaxEEP02 '序號 Add By Sindy 2017/8/14
         frm090202_4_1.SetParent Me
         'Modify By Sindy 2018/5/2
         frm090202_4_1.m_ProState = "P"
'         'Modify By Sindy 2018/4/27
'         If txtSystem = "CFP" Then
'            Set frm090202_4_1.m_SendRecvForm = frm050102_1
'         ElseIf txtSystem = "P" Then
'            Set frm090202_4_1.m_SendRecvForm = frm040104_1
'         End If
'         '2018/4/27 END
         '2018/5/2 END
         frm090202_4_1.m_NPManKind = "3"
         If frm090202_4_1.QueryData = True Then
            frm090202_4_1.Show
            Me.Hide
         End If
         Exit For
      End If
   Next i
   
   Set rsTmp = Nothing
End Sub

Private Sub cmdExit_Click()
   m_PrevForm.Hide
   'If UCase(m_PrevForm.Name) = UCase("frm090202_4") Then
      m_PrevForm.QueryData
   'End If
   m_PrevForm.Show
   Unload Me
End Sub

'查詢
Private Sub cmdQuery_Click()
   If QueryData = False Then ShowNoData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
   
   txtCode(0) = Right("000000" & Trim(txtCode(0)), 6)
   If txtCode(1) = "" Then txtCode(1) = "0"
   If txtCode(2) = "" Then txtCode(2) = "00"
   
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   SetGrd
   
   Screen.MousePointer = vbHourglass
   
   'P非台灣案
   'Modify By Sindy 2022/2/7 針對品薇的特殊操作方式是因為國外部均不會跑歷程,故才會開此特例,故請控制此模式僅限於FMP案件
   '==> and substr(cp12,1,1)='F'
   'Modify By Sindy 2022/2/11 要再開放承辦人掛陳品薇時,也能操作
   '==> and (substr(cp12,1,1)='F' or cp14='98012')
   'Modified by Morgan 2025/1/21 or cp14='98012' => or s2.st03='P12' P案程序人員工作改依智權區域分配承辦人不再限定98012
   strSql = "select ' ' as V,SqlDateT(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質" & _
            ",cp43 as 相關收文號,s2.st02 as 承辦人,s1.st02 as 智權人員,SqlDateT(CP06) as 本所期限,SqlDateT(CP07) as 法定期限" & _
            ",cp64 as 進度備註,SQLDatet2(CP05) sort1,cp66,CP67,CP09,pa11,pa22,pa05,pa06,pa07,pa47,pa09" & _
            " from caseprogress,casepropertymap,staff s1,staff s2,patent" & _
            " where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & txtCode(1) & "' and cp04='" & txtCode(2) & "'" & _
            " and cp27 is null and cp57 is null" & _
            " and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa09<>'000'" & _
            " and (substr(cp12,1,1)='F' or s2.st03='P12')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09 not in(Select eep01 from EmpElectronProcess where eep01=cp09 and eep04 not in('" & EMP_聯絡 & "'))"
   strSql = strSql & " union " & _
            "select ' ' as V,SqlDateT(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10) as 案件性質" & _
            ",cp43 as 相關收文號,s2.st02 as 承辦人,s1.st02 as 智權人員,SqlDateT(CP06) as 本所期限,SqlDateT(CP07) as 法定期限" & _
            ",cp64 as 進度備註,SQLDatet2(CP05) sort1,cp66,CP67,CP09,sp11,sp13,sp05,sp06,sp07,sp28,sp09" & _
            " from caseprogress,casepropertymap,staff s1,staff s2,ServicePractice" & _
            " where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & txtCode(1) & "' and cp04='" & txtCode(2) & "'" & _
            " and cp27 is null and cp57 is null" & _
            " and cp01=SP01 and cp02=SP02 and cp03=SP03 and cp04=SP04 and SP09<>'000'" & _
            " and (substr(cp12,1,1)='F' or s2.st03='P12')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09 not in(Select eep01 from EmpElectronProcess where eep01=cp09 and eep04 not in('" & EMP_聯絡 & "'))"
   strSql = strSql & " ORDER BY sort1 DESC, CP66 DESC, CP67 DESC, CP09 DESC"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      
      Label3.Caption = txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2)
      AddCboName Combo1, "" & rsTmp.Fields("pa05").Value, "" & rsTmp.Fields("pa06").Value, "" & rsTmp.Fields("pa07").Value
      If IsNull(rsTmp.Fields("pa11")) Then
          Label4.Caption = ""
      Else
          Label4.Caption = rsTmp.Fields("pa11")
      End If
      If IsNull(rsTmp.Fields("pa22")) Then
          Me.Label12.Caption = ""
      Else
          Me.Label12.Caption = rsTmp.Fields("pa22")
      End If
      If IsNull(rsTmp.Fields("pa47")) Then
          Me.Label15.Caption = ""
      Else
          Me.Label15.Caption = rsTmp.Fields("pa47")
      End If
      '顯示智權人員
      Me.Label16.Caption = ShowCurrCP13(txtSystem, txtCode(0), txtCode(1), txtCode(2), rsTmp.Fields("pa09"))
      Me.Label16.Caption = GetStaffName(Me.Label16.Caption)
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   
   dblPrevRow = 0
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   dblPrevRow = GRD1.row
   If GRD1.Rows - 1 = 1 And GRD1.TextMatrix(GRD1.row, 2) <> "" Then
      GRD1.Text = "V"
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   Label16.Caption = "" 'Add By Sindy 2021/12/23
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090202_7 = Nothing
End Sub

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "收文日", "總收文號", "案件性質", "相關收文號", _
                           "承辦人", "智權人員", "本所期限", "法定期限", "進度備註")
   arrGridHeadWidth = Array(200, 800, 800, 1200, 1000, _
                            800, 800, 800, 800, 3500)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
Dim i As Integer

GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 0
      GRD1.row = dblPrevRow
      GRD1.Text = ""
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
'   If grd1.Text = "V" Then
'      grd1.Text = ""
'      For i = 0 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = QBColor(15)
'      Next i
'   Else
      If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
'   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "目次" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub txtSystem_GotFocus()
   CloseIme
   txtSystem.SetFocus
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem)
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   If txtSystem <> "" Then
      If ChkSysName(txtSystem) = True Then
         If txtSystem <> "P" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtSystem
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   CloseIme
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
