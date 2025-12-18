VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050312_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限通知管制表-EPC年費金額"
   ClientHeight    =   5730
   ClientLeft      =   -1290
   ClientTop       =   1290
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9300
   Begin VB.TextBox txt2 
      Height          =   264
      Left            =   5430
      MaxLength       =   2
      TabIndex        =   2
      Top             =   708
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改"
      Height          =   300
      Left            =   6480
      TabIndex        =   3
      Top             =   708
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   7980
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7200
      TabIndex        =   4
      Top             =   70
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Left            =   3630
      MaxLength       =   6
      TabIndex        =   1
      Top             =   708
      Width           =   1000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4692
      Left            =   0
      TabIndex        =   0
      Top             =   1044
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label2 
      Caption         =   "點數："
      Height          =   180
      Index           =   2
      Left            =   4785
      TabIndex        =   14
      Top             =   750
      Width           =   585
   End
   Begin VB.Label Label4 
      Caption         =   "點數："
      Height          =   180
      Left            =   1170
      TabIndex        =   13
      Top             =   540
      Width           =   540
   End
   Begin VB.Label LBL3 
      Height          =   180
      Left            =   1770
      TabIndex        =   12
      Top             =   540
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "金額："
      Height          =   180
      Index           =   1
      Left            =   2985
      TabIndex        =   11
      Top             =   750
      Width           =   585
   End
   Begin VB.Label LBL2 
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   750
      Width           =   1335
   End
   Begin VB.Label LBL2 
      Height          =   180
      Index           =   0
      Left            =   720
      TabIndex        =   9
      Top             =   744
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "國家："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   744
      Width           =   552
   End
   Begin VB.Label LBL1 
      Height          =   180
      Left            =   1770
      TabIndex        =   7
      Top             =   330
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "前畫面輸入之金額："
      Height          =   180
      Left            =   90
      TabIndex        =   6
      Top             =   330
      Width           =   1620
   End
End
Attribute VB_Name = "frm050312_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/13 智權人員欄已修改
'Memo By Sindy 2010/12/7 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim SeekMouseClick As Integer, i As Integer, j As Integer, s As Integer, IntTotle As Long
Dim strSql As String, strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String, strNP07 As String
Dim BolStart As Boolean, SeekRow As Integer, BolLostFocus As Boolean, BolGrdClick As Boolean
Dim StrExt1 As String '例外欄位---列印備註
Dim StrExt2 As String '例外欄位---費用
Dim StrExt3 As String '例外欄位---點數
Public m_NP01 As String, m_NP22 As String, m_st02 As String, m_NP09 As String, m_NP08 As String
Public m_PA25 As String 'Added by Morgan 2016/8/10
Dim iPoint As Integer
Dim intTotalPoint As Integer 'Added by Lydia 2016/09/29 加總輸入點數
'Add By Sindy 2017/12/29
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2017/12/29 END
Dim m_LD18 As String, m_PA09 As String, m_PA26 As String, m_PA75 As String 'Added by Morgan 2018/7/17

Private Sub Command1_Click()
   If IsNumeric(txt1) = False Then
      s = MsgBox("請輸入數字！！", , "User 輸入錯誤")
      txt1.SetFocus
      txt1_GotFocus
   End If
   'Added by Lydia 2016/09/29
   If IsNumeric(txt2) = False Then
      s = MsgBox("請輸入數字！！", , "User 輸入錯誤")
      txt2.SetFocus
      txt2_GotFocus
   End If
   '控制年費智權同仁可加的點數
   If Val(txt2) > CFP_dg605 Then
      If MsgBox("請注意！年費(605)超過" & CFP_dg605 & "點，是否確定？", vbYesNo + vbDefaultButton2, "控制點數") = vbNo Then
         txt2.SetFocus
         txt2_GotFocus
         Exit Sub
      End If
   End If
   'end 2016/09/29
   
   BolLostFocus = True
   Grd1_Click
   BolLostFocus = False
End Sub

Private Sub Form_Activate()
   If StrMenu = False Then
      frm050312.Show
      Unload Me
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
   BolStart = True
   SeekRow = 0
   MoveFormToCenter Me
   strNP02 = frm050312.txt2(0)
   strNP03 = frm050312.txt2(1)
   strNP04 = frm050312.txt2(2)
   strNP05 = frm050312.txt2(3)
   strNP07 = frm050312.txt2(4)
   SetDataListWidth
   
   'Add By Sindy 2017/12/28
   m_strIR01 = frm050312.m_strIR01
   m_strIR02 = frm050312.m_strIR02
   m_strIR03 = frm050312.m_strIR03
   m_strIR04 = frm050312.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/28 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050312_1 = Nothing
End Sub

Private Sub Grd1_Click()
With GRD1
    If BolStart = True Then
        SeekMouseClick = 1
    Else
        If BolLostFocus = False Then
            SeekMouseClick = .MouseRow
        Else
            SeekMouseClick = .row
            If SeekRow <> 0 Then
                .row = SeekRow
                .col = 2
                .Text = Val(txt1)
                
                'Added by Lydia 2016/09/29 +點數
                .row = SeekRow
                .col = 3
                .Text = Val(txt2)
            End If
            
        End If
    End If
    .Visible = False
    .col = 0
    If SeekMouseClick <> 0 Then
        If BolStart = False Then
            For i = 1 To .Rows - 1
                .row = i
                .col = 0
                If .CellBackColor = &HFFC0C0 Then
                    For j = 0 To .Cols - 1
                        .col = j
                        .CellBackColor = QBColor(15)
                    Next j
                    Exit For
                End If
            Next i
        Else
            BolStart = False
        End If
        .row = SeekMouseClick
        SeekRow = SeekMouseClick
        For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = &HFFC0C0
        Next i
        .col = 0
        lbl2(0).Caption = .Text
        .col = 1
        lbl2(1).Caption = .Text
        .col = 2
        txt1.Text = .Text
        'Added by Lydia 2016/09/29
        .col = 3
        txt2.Text = .Text
    End If
    .Visible = True
End With
End Sub

Private Sub txt1_GotFocus()
   TextInverse txt1
End Sub

Private Sub txt1_LostFocus()
   'Remove by Lydia 2016/09/29
   'If IsNumeric(txt1) = False Then
   '   s = MsgBox("請輸入數字！！", , "User 輸入錯誤")
   '   txt1.SetFocus
   '   txt1_GotFocus
   'End If
End Sub

'讀出資料
Function StrMenu() As Boolean
   '93.9.17 MODIFY BY SONIA 只帶出未閉卷之子案
   'strSQL = "select pa09,na03,'' from patent,nation where pa01='" & strNP02 & "' and pa02='" & strNP03 & "' " & IIf(Len(strNP04) = 0, " AND PA03='0' ", " and pa03='" & strNP04 & "' ") & " and pa04<>'00' and pa09=na01(+) ORDER BY PA09 "
   'Modify by Morgan 2008/1/24 EPC進各國要照各國年費規定，故需排除該年度不需繳費的國家
   'strSQL = "select pa09,na03,'' from patent,nation where pa01='" & strNP02 & "' and pa02='" & strNP03 & "' " & IIf(Len(strNP04) = 0, " AND PA03='0' ", " and pa03='" & strNP04 & "' ") & " and pa04<>'00' and pa09=na01(+) AND PA57 IS NULL ORDER BY PA09 "
   '93.9.17 END
   Dim iNextYear As Integer, strCon As String, ArrYear
   strSql = "select pa72,na21,pa09,pa26,pa75 from patent,nation where pa01='" & strNP02 & "' and pa02='" & strNP03 & "' " & IIf(Len(strNP04) = 0, " AND PA03='0' ", " and pa03='" & strNP04 & "' ") & " and pa04='00' and pa09=na01(+)"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      m_PA09 = "" & adoRecordset.Fields("pa09") 'Added by Morgan 2018/7/17
      m_PA26 = "" & adoRecordset.Fields("pa26") 'Added by Morgan 2018/7/17
      m_PA75 = "" & adoRecordset.Fields("pa75") 'Added by Morgan 2018/7/17
      If IsNull(adoRecordset.Fields("pa72")) Then
         If IsNull(adoRecordset.Fields("na21")) Then
            iNextYear = 3
         Else
            iNextYear = Val(adoRecordset.Fields("na21")) 'VB會回傳第一個數字
         End If
      Else
         ArrYear = Split(adoRecordset.Fields("pa72"), ",")
         iNextYear = Val(ArrYear(UBound(ArrYear))) + 1
      End If
   End If
   If iNextYear > 0 Then
      strCon = " and instr(','||na21||',','," & iNextYear & ",')>0"
   End If
   'Modified by Lydia 2016/09/29 pa09,na03,'' => pa09,na03,'',''
   strSql = "select pa09,na03,'','' from patent,nation where pa01='" & strNP02 & "' and pa02='" & strNP03 & "' " & IIf(Len(strNP04) = 0, " AND PA03='0' ", " and pa03='" & strNP04 & "' ") & " and pa04<>'00' and pa09=na01(+) AND PA57 IS NULL" & strCon & " ORDER BY PA09 "
   'end 2008/1/24
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Set GRD1.Recordset = adoRecordset
      SetDataListWidth
      'Modify by Morgan 2011/2/15
      'Grd1_Click
      'StrMenu = True
      'Remove by Lydia 2016/09/29 點數改成可分別輸入
      'If Val(frm050312_1.LBL3.Caption) Mod (grd1.Rows - 1) <> 0 Then
      '   s = MsgBox("點數分配錯誤，請重新輸入總點數！", , "User 輸入錯誤")
      '   StrMenu = False
      'Else
      '   iPoint = Val(frm050312_1.LBL3.Caption) / (grd1.Rows - 1)
         Grd1_Click
         StrMenu = True
      'End If
      'end 2016/09/29
      
   Else
      s = MsgBox("此本所案號" & strNP02 & "-" & strNP03 & "-" & strNP04 & "無子案！", , "User 輸入錯誤")
      StrMenu = False
   End If
End Function

Private Sub SetDataListWidth()
   With GRD1
      'Modified by Lydia 2016/09/29 +點數
      '.Cols = 3
      .Cols = 4
      .row = 0
      .col = 0
      .Text = "國家代號"
      .ColWidth(0) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 1
      .Text = "國家名稱"
      .ColWidth(1) = 2000
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "金額"
      .ColWidth(2) = 1500
      .CellAlignment = flexAlignCenterCenter
      'Added by Lydia 2016/09/29
      .col = 3
      .Text = "點數"
      .ColWidth(2) = 1500
      .CellAlignment = flexAlignCenterCenter
   End With
End Sub

Private Sub cmdOK_Click(Index As Integer)

   Select Case Index
      Case 0
         IntTotle = 0
         intTotalPoint = 0 'Added by Lydia 2016/09/29
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 2
            If Len(Trim(GRD1.Text)) = 0 Then
               GRD1.col = 1
               s = MsgBox("國家 " & GRD1.Text & " 之金額不可空白！", , "User 輸入錯誤")
               Exit Sub
            Else
               IntTotle = IntTotle + Val(GRD1.Text)
            End If
            
            'Added by Lydia 2016/09/29
            If Val("" & GRD1.TextMatrix(i, 3)) = 0 Then
               s = MsgBox("國家 " & GRD1.TextMatrix(i, 1) & " 之點數不可空白！", , "User 輸入錯誤")
               Exit Sub
            Else
               intTotalPoint = intTotalPoint + Val(GRD1.TextMatrix(i, 3))
            End If
            'end 2016/09/29
         Next i
         
         If IntTotle <> Val(lbl1.Caption) Then
            s = MsgBox("所輸入之總金額與前畫面不符！", , "User 輸入錯誤")
            Exit Sub
         End If
         'Added by Lydia 2016/09/29
         If intTotalPoint <> Val(lbl3.Caption) Then
            s = MsgBox("所輸入之總點數與前畫面不符！", , "User 輸入錯誤")
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         StrExt2 = lbl1.Caption
         StrExt3 = lbl3.Caption
         StrExt1 = ""
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            GRD1.col = 1
            StrExt1 = StrExt1 & "　　" & str(i) & "." & GRD1.Text & "：新台幣 "
            GRD1.col = 2
            StrExt1 = StrExt1 & GRD1.Text & " 元整。" & Chr$(13)
         Next i
      
         PrintLetter "10", m_NP01, "08"
         '92.7.7 ADD BY SONIA
         g_PrtForm001.PrintForm m_NP22, strNP02, strNP03, IIf(Len(Trim(strNP04)) = 0, "0", strNP04), IIf(Len(Trim(strNP05)) = 0, "00", strNP05)
         '92.7.7 END
         ShowPrintOk
         Screen.MousePointer = vbDefault
         
         'Add by Sindy 2017/12/29
         If m_strIR01 <> "" Then
            PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050312"
         End If
         '2017/12/29 END
         
         'Modified by Lydia 2016/09/29
         'For i = 0 To 7
         For i = 0 To 9
            If i <> 6 Then frm050312.txt2(i) = "" 'Modified by Morgan 2018/9/21 +i<>6
         Next i
         
         'Add By Sindy 2017/12/29
         If m_strIR01 <> "" Then
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
            Unload frm050312
            Unload Me
         Else
         '2017/12/29 END
            frm050312.m_InputEPC = True 'Added by Lydia 2016/09/29
            frm050312.lbl1 = ""
            frm050312.Show
            Unload Me
         End If
      Case 1
         frm050312.Show
         Unload Me
      Case Else
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter(ByVal strET01 As String, ByVal strCP09 As String, ByVal strET03 As String)
   Dim arrCP(1 To 4) As String, dblYear As Double
   
   'Added by Morgan 2013/9/27
   'Modified by Morgan 2016/1/13 改轉定稿時上發文日且要在定稿產生前新增,否則報價即時轉定稿的狀況會沒有進度可上發文日
   If PUB_AddCP1913(strNP02, strNP03, strNP04, strNP05, m_NP08, m_NP09, m_NP01, m_NP22, m_PA09, m_PA26, m_LD18, m_PA75, True, , IIf(frm050312.Check1.Value = vbChecked, False, True)) = False Then
      MsgBox "新增進度檔【通知期限】失敗！作業中斷！", vbCritical
      Exit Sub
   End If
   'end 2013/9/27
   
   'Remove by Morgan 2008/8/20
   'InsExpField strET01, strCP09, strET03
   'NowPrint strCP09, strET01, strET03, IIf(frm050312.txt2(6) = "Y", True, False), strUserNum, 0
   
   'Add by Morgan 2008/5/1 新增年費報價通知
   'Modify by Morgan 2010/1/11 +下次繳費年
   'PUB_AddLetterCache m_NP01, m_NP22, m_NP01, strET01, strET03
   arrCP(1) = strNP02: arrCP(2) = strNP03: arrCP(3) = strNP04: arrCP(4) = strNP05
   dblYear = PUB_GetNextYear(arrCP)
   PUB_AddLetterCache m_NP01, m_NP22, m_NP01, strET01, strET03, dblYear, m_LD18
   InsExpField1
   'end 2008/5/1
         
   'Add by Morgan 2008/10/21
   '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
   strExc(0) = CompWorkDay(5, strSrvDate(1))
   strExc(1) = DBDATE(m_NP08)
   'Modify by Morgan 2009/1/6 開放維護功能(因常有需要重新產生定稿)--慧汶
   'Memo 2018/03/20 維護功能已取消--Morgan
   If Val(strExc(1)) <= Val(strExc(0)) Then
      PUB_Cache2Letter m_NP01, m_NP22, False
   End If
End Sub

Private Sub InsExpField(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
   Dim strTxt(1 To 99) As String, iStep As Integer
   
   iStep = 1

   EndLetter ET01, ET02, ET03, strUserNum
   
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','下一程序業務員','" & m_st02 & "')"
   iStep = iStep + 1
   
   
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','下一程序','" & strNP07 & "')"
   iStep = iStep + 1
   
   
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','下一程序名稱','" & GetPrjState6HM("CFP", strNP07) & "')"
   iStep = iStep + 1
   
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','法定期限','" & m_NP09 & "')"
   iStep = iStep + 1
   
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','本所期限','" & m_NP08 & "')"
   iStep = iStep + 1
   
   'modify by Morgan 2008/5/12 變數"費用"改為"費用合計",預定5/19 定稿修改後移除
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','費用合計','" & StrExt2 & "')"
   iStep = iStep + 1
   
   'modify by Morgan 2008/5/12 變數名稱"點數"改為"點數合計"
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','點數合計','" & StrExt3 & "')"
   iStep = iStep + 1
   
   'modify by Morgan 2008/5/12 變數名稱"列印備註"改為"EPC指定國家年費費用"
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','EPC指定國家年費費用','" & StrExt1 & "')"
   iStep = iStep + 1
   
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add by Morgan 2008/5/16
Private Sub InsExpField1()
   Dim strTxt(1 To 99) As String, iStep As Integer
   
   iStep = 1
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
      "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'下一程序業務員','" & m_st02 & "')"
   iStep = iStep + 1
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
      "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'下一程序','" & strNP07 & "')"
   iStep = iStep + 1
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
      "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'下一程序名稱','" & GetPrjState6HM("CFP", strNP07) & "')"
   iStep = iStep + 1
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
      "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'法定期限','" & m_NP09 & "')"
   iStep = iStep + 1
   
   'Added by Morgan 2016/8/9
   '若法定期限與專用期止日相差不足半年時定稿帶出將屆滿的句子
   If Val(m_PA25) > 0 And Val(m_NP09) > 0 Then
      If m_PA25 < CompDate(1, 6, m_NP09) Then
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'即將屆滿','♀')"
         iStep = iStep + 1
      End If
   End If
   'end 2016/8/9
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
      "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'本所期限','" & m_NP08 & "')"
   iStep = iStep + 1
   
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
      "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'費用合計','" & StrExt2 & "','')"
   iStep = iStep + 1
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
      "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'點數合計','" & StrExt3 & "','')"
   iStep = iStep + 1
   '指定國年費、點數
   For i = 1 To GRD1.Rows - 1
      GRD1.row = i
      GRD1.col = 1
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'" & GRD1.TextMatrix(i, 1) & "年費','" & GRD1.TextMatrix(i, 2) & "','Y')"
      iStep = iStep + 1
      
      'Modified by Lydia 2016/09/29 iPoint => grd1.TextMatrix(i, 3)
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & m_NP01 & "'," & m_NP22 & ",'" & GRD1.TextMatrix(i, 1) & "年費點數','" & GRD1.TextMatrix(i, 3) & "')"
      iStep = iStep + 1
   Next i
   
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Added by Lydia 2016/09/29
Private Sub txt2_GotFocus()
   TextInverse txt2
End Sub

