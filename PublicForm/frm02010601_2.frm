VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010601_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   4410
   ClientLeft      =   -2220
   ClientTop       =   2520
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時收達(&R)"
      Height          =   405
      Index           =   3
      Left            =   4320
      TabIndex        =   38
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   3
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   3
      Top             =   4020
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   2
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3660
      Width           =   2892
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   1
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3300
      Width           =   1212
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   0
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2940
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6372
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5544
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7596
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   1380
      Width           =   7335
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12938;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPromoter 
      Height          =   255
      Left            =   1920
      TabIndex        =   37
      Top             =   2100
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   36
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   5340
      TabIndex        =   35
      Top             =   2100
      Width           =   2175
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   34
      Top             =   2460
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   33
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5340
      TabIndex        =   32
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   31
      Top             =   1740
      Width           =   375
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   30
      Top             =   1020
      Width           =   2535
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   29
      Top             =   660
      Width           =   2535
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   28
      Top             =   660
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "是否列印定稿：             （N：不印）"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4020
      Width           =   3015
   End
   Begin VB.Label Label30 
      Caption         =   "彼所案號："
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3660
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "代理人提申日："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   3285
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "代理人收達日："
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2940
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "代理人："
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2460
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label lblIssue 
      Caption         =   "發文日："
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   1740
      Width           =   975
   End
   Begin MSForms.Label lblSales 
      Height          =   255
      Left            =   6060
      TabIndex        =   17
      Top             =   1740
      Width           =   2055
      VariousPropertyBits=   27
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseProperty 
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   1740
      Width           =   2295
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   2460
      Width           =   6615
      VariousPropertyBits=   27
      Size            =   "11668;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1020
      Width           =   975
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   6000
      TabIndex        =   13
      Top             =   1020
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   11
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "審定號數："
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   9
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "frm02010601_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/9 改成Form2.0 (cboCaseName,lblSales...)
'Memo By Morgan 2012/12/17 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'Add By Cheng 2002/07/25
Dim m_blnSameReceive As Boolean
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_PrevForm As Form 'Add By Sindy 2016/10/11


'Add By Sindy 2016/10/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer

'Add By Cheng 2002/07/25
'預設非按下同時收達
m_blnSameReceive = False
Select Case Index
             Case 0 '確定
                        Screen.MousePointer = vbHourglass
                        For i = 0 To 3
                               If txtCaseField(i).Enabled And txtCaseField(i).Visible Then
                                  If CheckKeyIn(i) <> 1 Then
                                     txtCaseField(i).SetFocus
                                     txtCaseField_GotFocus (i)
                                     Exit For
                                  End If
                               End If
                        Next
                        If i = 4 Then
                           'Add By Cheng 2002/05/22
                           '重新檢查欄位有效性
                           If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
                           
                           'Add by Sindy 2017/10/23 已收達信件要歸卷
                           If m_strIR01 <> "" Then
'                              'Add By Sindy 2022/7/1
'                              If Left(Pub_StrUserSt03, 2) = "F2" Then
'                                 If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
'                                    Screen.MousePointer = vbDefault
'                                    Exit Sub
'                                 End If
'                              Else
'                              '2022/7/1 END
                                 If frm02010601_1.txtChoose.Text = "1" Then
                                    '下載信件檔,上傳卷宗區
                                    If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, cp(9), IIf(Pub_StrUserSt03 = "F22", "ALTR", "ACK")) = False Then
                                       Screen.MousePointer = vbDefault
                                       Exit Sub
                                    End If
                                 End If
'                              End If
                           End If
                           '2017/10/23 END
                           
                           If SaveDatabase Then
                              'Add By Sindy 2016/10/7
                              If Me.m_strIR01 <> "" Then
                                 bolLeave = True
                                 intLeaveKind = 1
                                 'Unload frm02010601_1
                                 'Add By Sindy 2016/10/11
                                 If Not m_PrevForm Is Nothing Then
                                    Call m_PrevForm.GoNext
                                 End If
                                 '2016/10/11 END
                                 m_blnSameReceive = True
                                 Unload Me
                              Else
                              '2016/10/7 END
                                 bolLeave = True
                                 intLeaveKind = 1
                                 'Add By Cheng 2002/07/25
                                 m_blnSameReceive = False
                                 Unload Me
                              End If
                            'Add By Cheng 2002/11/06
                            Else
                                MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
                           End If
                        End If
                        Screen.MousePointer = vbDefault
             Case 1, 2 '回前畫面, 結束
                        If Index = 2 Then
                           intLeaveKind = 0
                        Else
                           intLeaveKind = 1
                           m_blnSameReceive = True
                        End If
                        Unload Me
            'Add By Cheng 2002/07/25
            Case 3 '同時收達
                        Screen.MousePointer = vbHourglass
                        For i = 0 To 3
                               If txtCaseField(i).Enabled And txtCaseField(i).Visible Then
                                  If CheckKeyIn(i) <> 1 Then
                                     txtCaseField(i).SetFocus
                                     txtCaseField_GotFocus (i)
                                     Exit For
                                  End If
                               End If
                        Next
                        If i = 4 Then
                           'Add By Cheng 2002/05/22
                           '重新檢查欄位有效性
                           If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
                        
                           If SaveDatabase Then
                              bolLeave = True
                              intLeaveKind = 1
                              'Add By Cheng 2002/07/25
                              '預設按下同時收達
                              m_blnSameReceive = True
                              
                              Unload Me
                           End If
                        End If
                        Screen.MousePointer = vbDefault
            
End Select
End Sub

Private Function SaveDatabase() As Boolean
Dim strTxt(1 To 10) As String, iStep As Integer

'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
SaveDatabase = True
cnnConnection.BeginTrans
   
   If frm02010601_1.txtChoose = "1" Then
      cp(46) = txtCaseField(0)
      '91.12.28 ADD BY SONIA
   'Else
      If txtCaseField(1).Enabled = True And txtCaseField(1).Text <> "" Then
      '91.12.28 END
         cp(47) = txtCaseField(1)
      '91.12.28 ADD BY SONIA
      End If
      '91.12.28 END
   End If
   cp(45) = txtCaseField(2)
   
   strTxt(1) = GetCPSQL(cp())
   'Modify By Cheng 2002/11/06
   '   SaveDatabase = objLawDll.ExecSQL(1, strTxt())
   cnnConnection.Execute strTxt(1)
   'Add By Cheng 2002/11/08
   '更新基本檔的申請日
   '91.12.28 MODIFY BY SONIA
   'If Me.txtCaseField(4).Enabled = True And Me.txtCaseField(4).Text <> "" Then
   '    strTxt(2) = "Update Patent Set PA10 = " & DBDATE(Me.txtCaseField(4).Text) & " Where " & ChgPatent(cp(1) & cp(2) & cp(3) & cp(4))
   '    cnnConnection.Execute strTxt(2)
   'End If
   If txtCaseField(1).Enabled = True And txtCaseField(1).Text <> "" Then
      strTxt(2) = "Update Patent Set PA10 = " & DBDATE(txtCaseField(1).Text) & " Where " & ChgPatent(cp(1) & cp(2) & cp(3) & cp(4))
      cnnConnection.Execute strTxt(2)
   End If
   '91.12.28 END
   '2005/3/30 ADD BY SONIA
   strTxt(3) = "update nextprogress set np06='Y' where np01='" & cp(9) & "' and np02='" & cp(1) & _
     "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np07=" & 收達 & " "
   cnnConnection.Execute strTxt(3)
   '2005/3/30 END
    
   'Add by Morgan 2005/4/27 美國領證收達同時上公開費收達
   If field(9) = "101" And cp(10) = "601" Then
     strSql = "Update CaseProgress Set CP46=" & cp(46) & " Where CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='217' AND CP46 IS NULL"
     cnnConnection.Execute strSql
   End If
   '2005/12/19 ADD BY SONIA 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
   If txtCaseField(2) <> "" Then
     'Modified by Morgan 2012/2/15 取消 cp09<'C' 條件(C類也會有發文作業,有代理人就要更新彼號,資料才會一致)
     strTxt(4) = "update caseprogress set cp45=" & CNULL(ChgSQL(txtCaseField(2))) & " where cp09 in (select cp09 from caseprogress where cp45 is null and CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND cp44 in (select cp44 from caseprogress where cp09='" & cp(9) & "' ))"
     cnnConnection.Execute strTxt(4)
   End If
   '2005/12/19 END
   
   'Add by Morgan 2011/2/16
   '歐盟要同時更新集體設計子案的收達日(提申日)
   If field(9) = "239" Then
      strSql = "update caseprogress set cp46=nvl(cp46," & CNULL(cp(46), True) & "),cp47=nvl(cp47," & CNULL(cp(47), True) & ")" & _
      " where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp04='" & field(4) & "'" & _
      " and cp10='105' and cp27=" & DBDATE(cp(27))
      cnnConnection.Execute strSql, intI
      
      'Added by Morgan 2019/7/2 收達期限也要解除 Ex:CFP-031097-1-00--禧佩
      strSql = "update nextprogress set np06='Y'" & _
         " where (np01,np02,np03,np04,np05) in (select cp09,cp01,cp02,cp03,cp04" & _
         " from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp04='" & field(4) & "'" & _
         " and cp10='105' and cp27=" & DBDATE(cp(27)) & ") and np07=" & 收達 & " and np06 is null"
      cnnConnection.Execute strSql, intI
      'end 2019/7/2
   End If
   
   'Added by Lydia 2016/08/26 +438 再考量試行計畫(AFCP 2.0)要更新下一程序提申期限為收達日+7天
   If cp(10) = "438" Then
      strExc(1) = CompDate(2, 7, TransDate(txtCaseField(0), 2))
      strSql = "update nextprogress set np08=" & CNULL(strExc(1), True) & " ,np09=" & CNULL(strExc(1), True) & _
               " where np01='" & cp(9) & "' and np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'" & _
               " and nvl(np06,' ')=' ' and np07='" & 提申 & "' "
      cnnConnection.Execute strSql, intI
   End If
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/7/1 + , IIf(Pub_StrUserSt03 = "F22", cp(9), "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010601_1", IIf(Pub_StrUserSt03 = "F22", cp(9), "")
   End If
   '2016/10/7 END
   
   cnnConnection.CommitTrans
   Exit Function
ErrorHandler:
   cnnConnection.RollbackTrans
   SaveDatabase = False
End Function

Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer

On Error GoTo ErrHnd
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.ReadAllData(frm02010601_1.grdDataList.TextMatrix(frm02010601_1.grdDataList.Row, 0), cp(), field(), intCaseKind, intPWhere) Then
   ReDim cp(TF_CP) As String
   cp(9) = frm02010601_1.grdDataList.TextMatrix(frm02010601_1.grdDataList.row, 0)
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   If cp(1) = 馬德里案 Then
      lblCaseField(0) = cp(1) + " - " + Left(cp(2), 5) + _
         IIf(Right(cp(2), 1) = "0", "", " - " + Right(cp(2), 1)) + _
         IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
         IIf(cp(4) = "00", "", " - " + cp(4))
   Else
      lblCaseField(0) = cp(1) + " - " + cp(2) + _
         IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
         IIf(cp(4) = "00", "", " - " + cp(4))
   End If
   Select Case intCaseKind
                Case 專利
                           lblCaseField(1) = field(22)
                           lblCaseField(2) = field(11)
                           lblCaseField(8) = field(9)
                Case 商標
                           lblCaseField(1) = field(15)
                           lblCaseField(2) = field(12)
                           lblCaseField(8) = field(10)
                           
                Case Else
                           lblCaseField(1) = field(14)
                           lblCaseField(2) = field(11)
                           lblCaseField(8) = field(9)
   End Select
   lblCaseField(3) = cp(10)
   lblCaseField(4) = cp(13)
   lblCaseField(5) = cp(14)
   lblCaseField(6) = cp(44)
'91.12.28 MODIFY BY SONIA
'   If intPWhere <> 國外_CF Then
''      lblCaseField(7) = ChangeTStringToTDateString(cp(27))
'   Else
'      txtCaseField(1).Visible = False
'      Label6(1).Visible = False
''      lblCaseField(7) = ChangeWStringToWDateString(cp(27))
'   End If
   '若已收達的案件性質為新案申請(101 ~ 105), 則可輸入申請日
   'Modify by Morgan 2005/3/1 加CIP申請(112),CPA申請(113)
   'If Val(cp(10)) >= 101 And Val(cp(10)) <= 105 Then
   
   'Modified by Morgan 2016/9/29 提申一律從代理人案件提申作業--慧汶 (Ex.CFP-28836)
   'If (Val(cp(10)) >= 101 And Val(cp(10)) <= 105) Or cp(10) = "112" Or cp(10) = "113" Then
   '   txtCaseField(1).Visible = True
   '   txtCaseField(1).Enabled = True
   '   Label6(1).Visible = True
   'Else
      txtCaseField(1).Visible = False
      txtCaseField(1).Enabled = False
      Label6(1).Visible = False
   'End If
   'end 2016/9/29
   
'91.12.28 END
   lblCaseField(7) = ChangeTStringToTDateString(TransDate(cp(27), 1))
   SetNameToCombo cboCaseName, field(5), field(6), field(7)
   If cp(45) = "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseThatCode(cp()) = False Then GoTo Err1
      If ClsPDGetCaseThatCode(cp()) = False Then GoTo err1
   End If
   txtCaseField(2) = cp(45)
   '91.12.28 MODIFY BY SONIA
   'Add By Cheng 2002/11/08
   '若已收達的案件性質為新案申請(101 ~ 105), 則可輸入申請日
   'If Val(cp(10)) >= 101 And Val(cp(10)) <= 105 Then
   '     Me.Label6(2).Visible = True
   '     Me.txtCaseField(4).Visible = True
   '     Me.txtCaseField(4).Enabled = True
   'End If
   
Else
err1:
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
ErrHnd:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String, strCusTemp As String, bolIsChina As Boolean

Select Case Index
             Case 3
                       If lblCaseField(8) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                       'edit by nickc 2007/02/02 不用 dll 了
                       'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
                       If ClsPDGetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
                           lblCaseProperty.Caption = strTemp
                       End If
             Case 4
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
                        'Modified by Morgan 2022/5/6 離職還是要顯示且不必彈訊息
                        'If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
                        '   lblSales = strTemp
                        'Else
                        '   lblSales = ""
                        'End If
                        lblSales = GetStaffName(lblCaseField(Index), True)
                        'end 2022/5/6
             Case 5
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
                        'Modified by Morgan 2022/5/6 離職還是要顯示且不必彈訊息
                        'If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
                        '   lblPromoter = strTemp
                        'Else
                        '   lblPromoter = ""
                        'End If
                        lblPromoter = GetStaffName(lblCaseField(Index), True)
                        'end 2022/5/6
             Case 6
                        strCusTemp = lblCaseField(Index)
                        If lblCaseField(Index) <> "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           If ClsPDGetAgent(strCusTemp, strTemp) Then
                              lblCaseField(Index) = strCusTemp
                              lblAgent.Caption = strTemp
                           End If
                        End If
             Case 8
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
                        If ClsPDGetNation(lblCaseField(Index), strTemp) Then
                           lblNation.Caption = strTemp
                        End If
End Select
End Sub
Private Sub Form_Activate()
ReadAllData
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   'If frm02010601_1.txtChoose = "1" Then
   '   txtCaseField(1).Enabled = False
   'Else
   '   txtCaseField(0).Enabled = False
   'End If
   'Add By Cheng 2002/03/06
   '若CFP案為已收達, 則是否列印定稿預設為"N"
   If intPWhere = 國外_CF And frm02010601_1.txtChoose = "1" Then
      txtCaseField(3).Text = "N"
   End If
   
   Me.Caption = frm02010601_1.Caption
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   'If m_strIR01 <> "" Then intLeaveKind = 0 'Add By Sindy 2017/10/17
   'Add By Sindy 2016/10/13
'   If Me.m_strIR01 <> "" Then intLeaveKind = 0
   '2016/10/13 END
      If intLeaveKind = 1 Then
         'Modify By Cheng 2002/07/26
      '   frm02010601_1.Show
         frm02010601_1.Show: DoEvents
         'Modify By Cheng 2002/07/25
      '   frm02010601_1.Clear
         '若非按下同時收達
         If m_blnSameReceive = False Then
            frm02010601_1.Clear
         '若按下同時收達
         Else
            '還原為預設值
            m_blnSameReceive = False
            frm02010601_1.ClickCmdOk 2
         End If
      ElseIf intLeaveKind = 0 Then
        Unload frm02010601_1
      End If
'   End If
   
   'Add By Sindy 2016/10/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010601_2 = Nothing
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
             Case 3
                       KeyAscii = UpperCase(KeyAscii)
End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
End If
If Cancel Then txtCaseField_GotFocus (Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strCusTemp As String

CheckKeyIn = -1
Select Case intIndex
             Case 0
                     If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                        If intIndex = 0 Then
                           If Val(txtCaseField(intIndex)) > Val(strSrvDate(2)) Then
                              MsgBox "不可大於系統日 !", vbCritical
                           Else
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
                     End If
             '91.12.30 add by sonia
             Case 1
                  If txtCaseField(intIndex).Text <> "" Then
                     If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                        If intIndex = 0 Then
                           If Val(txtCaseField(intIndex)) > Val(strSrvDate(2)) Then
                              MsgBox "不可大於系統日 !", vbCritical
                           Else
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
                     End If
                  Else
                     CheckKeyIn = 1
                  End If
            '91.12.30 end
             Case 3
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
            '91.12.28 CANCEL BY SONIA
            'Add By Cheng 2002/11/08
            'Case 4 '申請日
            '    If Me.txtCaseField(intIndex).Text <> "" Then
            '         If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
            '            CheckKeyIn = 1
            '        End If
            '    Else
            '        CheckKeyIn = 1
            '    End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Me.txtCaseField
   'Modify By Cheng 2002/07/26
'   If objTxt.Enabled = True Then
   If objTxt.Enabled = True And objTxt.Visible = True Then
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added by Lydia 2016/08/26 AFCP要有收達日
If cp(10) = "438" And txtCaseField(0).Text = "" Then
   MsgBox "必要欄位不可空白！"
   txtCaseField(0).SetFocus
   Exit Function
End If

TxtValidate = True
End Function

