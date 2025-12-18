VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_2 
   Appearance      =   0  '平面
   BackColor       =   &H80000004&
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（延期）"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8544
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8544
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4140
      MaxLength       =   4
      TabIndex        =   10
      Top             =   3951
      Width           =   540
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   17
      Left            =   1365
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3630
      Width           =   375
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   16
      Left            =   4170
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "Y"
      Top             =   3630
      Width           =   375
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   19
      Left            =   7035
      MaxLength       =   1
      TabIndex        =   9
      Top             =   3630
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1575
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2640
      Width           =   300
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   4
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3300
      Width           =   372
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   492
      Index           =   2
      Left            =   7080
      TabIndex        =   15
      Top             =   60
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   492
      Index           =   0
      Left            =   4200
      TabIndex        =   12
      Top             =   60
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&Q)"
      CausesValidation=   0   'False
      Height          =   492
      Index           =   1
      Left            =   5640
      TabIndex        =   14
      Top             =   60
      Width           =   1332
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   3
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3300
      Width           =   372
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   2
      Left            =   5655
      MaxLength       =   8
      TabIndex        =   4
      Top             =   2970
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   1
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2970
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   0
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2310
      Width           =   972
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   1392
      Left            =   120
      TabIndex        =   11
      Top             =   4260
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   2455
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   2
   End
   Begin MSForms.ComboBox cboPatentName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1140
      TabIndex        =   13
      Top             =   1020
      Width           =   7275
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12832;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   3240
      TabIndex        =   44
      Top             =   3996
      Width           =   765
   End
   Begin VB.Label Label4 
      Caption         =   "是否印指示信：        （N:不印）"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   43
      Top             =   3645
      Width           =   2505
   End
   Begin VB.Label Label21 
      Caption         =   "是否修改指示信：        （Y:WORD）"
      Height          =   255
      Left            =   2715
      TabIndex        =   42
      Top             =   3645
      Width           =   2850
   End
   Begin VB.Label Label4 
      Caption         =   "是否印傳真封面：        （N:不印）"
      Height          =   255
      Index           =   4
      Left            =   5595
      TabIndex        =   41
      Top             =   3645
      Width           =   2685
   End
   Begin VB.Label Label12 
      Caption         =   "延期月數："
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label11 
      Caption         =   "是否修改通知函內容：           (Y:Word)"
      Height          =   255
      Left            =   4170
      TabIndex        =   39
      Top             =   3300
      Width           =   3315
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   6180
      TabIndex        =   38
      Top             =   2310
      Width           =   2235
      VariousPropertyBits=   27
      Size            =   "3942;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCasePropertyName 
      Height          =   255
      Left            =   1800
      TabIndex        =   37
      Top             =   1680
      Width           =   2175
      VariousPropertyBits=   27
      Size            =   "3836;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   5340
      TabIndex        =   36
      Top             =   1980
      Width           =   2475
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   35
      Top             =   1980
      Width           =   2775
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   34
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5340
      TabIndex        =   33
      Top             =   1380
      Width           =   2415
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   32
      Top             =   1380
      Width           =   2775
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   1
      Left            =   5160
      TabIndex        =   31
      Top             =   720
      Width           =   2652
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   30
      Top             =   720
      Width           =   2652
   End
   Begin VB.Label Label2 
      Caption         =   "代理人："
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Label40 
      Caption         =   "欲延期期限："
      Height          =   252
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "法定期限："
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   25
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "是否列印通知函：           (N:不印)"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3300
      Width           =   2775
   End
   Begin VB.Label Label9 
      Caption         =   "延期後本所期限："
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2970
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "延期日："
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "延期後法定期限："
      Height          =   255
      Index           =   1
      Left            =   4215
      TabIndex        =   21
      Top             =   2970
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "本所期限："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "案件性質："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "收文日："
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "收文號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號："
      Height          =   252
      Left            =   4200
      TabIndex        =   16
      Top             =   720
      Width           =   972
   End
End
Attribute VB_Name = "frm050102_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/4 改成Form2.0 (grdDataList,cboPatentName,lblCasePropertyName...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'intWhereComeFrom按延期鈕或延期性質 1:延期鈕     2:延期性質
Public intWhereComeFrom As Integer
'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'Add By Cheng 2002/06/20
Public m_str_DL05 As String '延期記錄檔的資料來源
'Add By Cheng 2003/09/16
Dim strCountry As String '存放EPC指定國家
Dim m_CP09 As String 'Added by Morgan 2013/9/24
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22
Dim m_bolEngLetter As Boolean, m_strSubject As String 'Added by Morgan 2018/9/6

Private Sub cmdOK_Click(Index As Integer)
   Dim varSaveCursor  As Variant, i As Integer
   Dim stLetter As String
   
   Select Case Index
      Case 0
         varSaveCursor = Screen.MousePointer
         Screen.MousePointer = vbHourglass
         For i = 0 To 4
            If CheckKeyIn(i) = False Then
               If i = 4 Then
                  Combo1.SetFocus
               Else
                  txtCaseField(i).SetFocus
                  txtCaseField_GotFocus (i)
               End If
               Exit For
            End If
         Next
         If i = 5 Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            
            If FormSave Then
               'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
               PUB_CheckEMail cp(44), cp(116)
               PUB_CheckEMail field(75), field(144)
               If field(145) <> "" Then
                  PUB_CheckEMail field(75), field(145)
               End If
               'end 2008/2/20
              
               'Added by Morgan 2013/9/24
               If txtCaseField(17) <> "N" Then
                  '傳真封面
                  If txtCaseField(19) <> "N" Then
                     StartLetter "01", "99"
                     If txtCaseField(16).Text = "Y" Then
                        NowPrint m_CP09, "01", "99", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
                     Else
                        NowPrint m_CP09, "01", "99", False, strUserNum, , , , , , , , , , , , , m_strAF01
                     End If
                     If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
                  End If
                  '指示信
                  StartLetter "01", "01"
                  NowPrint m_CP09, "01", "01", IIf(Me.txtCaseField(16).Text = "Y", True, False), strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
                  
                  'Added by Morgan 2018/8/22 CFP電子化
                  If txtCaseField(16).Text = "Y" And m_strAF01 <> "" Then
                     frm1105_1.m_RecNo = m_strAF01
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & 延期 & ".DATA.PDF"
                     frm1105_1.Show
                     If txtCaseField(4).Text = "Y" Then
                        MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                        txtCaseField(4).Text = ""
                     End If
                  End If
                  'end 2018/8/22
                  
               'Added by Morgan 2018/9/6
               ElseIf m_bolEngLetter Then
                  PUB_SendOrderLetterP m_strAF01, m_strSubject
               'end 2018/9/6
               
               End If
               'end 2013/9/24
               
               If txtCaseField(3) <> "N" Then
                  NowPrint cp(9), "01", "00", IIf(Me.txtCaseField(4).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                  
                  'Added by Morgan 2018/8/22 CFP電子化
                  If txtCaseField(4).Text = "Y" And m_strLD18 <> "" Then
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".CUS.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/8/22
               End If
               
               bolLeave = True
               intLeaveKind = 1
               'Add By Cheng 2002/04/30
               '若有未發文資料顯示警告
               If Me.lblCaseField(4).Caption = 延期 Then PUB_GetCPunIssueDatas "" & Me.lblCaseField(0).Caption
               
               Unload Me
            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
         End If
         Screen.MousePointer = varSaveCursor
      Case 1, 2
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            If Index = 2 Then
               intLeaveKind = 0
            Else
               intLeaveKind = 1
            End If
         End If
         Unload Me
   End Select
   ' 發文回前畫面時
   Select Case Index
      Case 0:
         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            frm050102_1.Clear
         End If
   End Select
End Sub

Private Function FormSave() As Boolean
 Dim i As Integer, j As Integer
'intWhereComeFrom按延期鈕或延期性質 1:延期鈕     2:延期性質
 Dim strTxt(1 To 50) As String, iStep As Integer
 'edit by nickc 2007/02/02
 'ReDim strDataTemp(1 To T_CP) As String
 ReDim strDataTemp(1 To TF_CP) As String
 
 Dim stNP01 As String 'Add by Morgan 2004/9/15
 Dim stNP22 As String 'Add by Morgan 2011/4/22
 Dim strCF03 As String 'Added by Morgan 2012/3/21
 'Added by Morgan 2015/8/7
 Dim strTemp As String, strCP09 As String, strCP10 As String
 Dim strLetterJudge As String '指示信判發人/主旨 Added by Morgan 2018/8/22
 
'911106 nick transation
FormSave = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   'FormSave = False
   iStep = 1
   
   'Modify by Morgan 2008/2/20
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/20
   cp(44) = ChangeCustomerL(cp(44))
      
   If intWhereComeFrom = 1 Then
      cnnConnection.Execute "delete datelimit where dl01='" & cp(9) & "' and dl02=" & DBDATE(txtCaseField(0)) 'Added by Morgan 2016/8/17
      
      strTxt(iStep) = "insert into datelimit (DL01,DL02,DL03,DL04,DL05) values (" + CNULL(cp(9)) + "," + CNULL(TransDate(txtCaseField(0), 2)) + "," + _
         CNULL(cp(6)) + "," + CNULL(cp(7)) + ",'1')"
      '911106 nick transation
      cnnConnection.Execute strTxt(iStep)
      
      iStep = iStep + 1
      
      strTxt(iStep) = "UPDATE CASEPROGRESS SET CP06=" & CNULL(TransDate(txtCaseField(1), 2)) & _
         ",CP07=" & CNULL(TransDate(txtCaseField(2), 2)) & " WHERE CP09 = " & CNULL(cp(9))
      
      '911106 nick transation
      cnnConnection.Execute strTxt(iStep)
      
      iStep = iStep + 1
      '92.6.30 ADD BY SONIA
      'cancel by sonia 2015/9/7 延期後不需工程師再輸收卷註記 P-104887
      'strTxt(iStep) = "UPDATE ENGINEERPROGRESS SET EP27=NULL,EP31=NULL WHERE EP02 = " & CNULL(cp(9))
      'cnnConnection.Execute strTxt(iStep)
      'iStep = iStep + 1
      '92.6.30 END
      strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
         " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
         " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      strDataTemp(45) = ""
      If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then strDataTemp(45) = RsTemp.Fields("CP45")
      
      strDataTemp(1) = cp(1)
      strDataTemp(2) = cp(2)
      strDataTemp(3) = cp(3)
      strDataTemp(4) = cp(4)
      strDataTemp(5) = strSrvDate(1)
'      strDataTemp(6) = cp(6)
      strDataTemp(6) = PUB_GetWorkDay1(cp(6), True)
      strDataTemp(7) = cp(7)
      strDataTemp(9) = 內部收文
      strDataTemp(10) = 延期
      strDataTemp(12) = cp(12)
      strDataTemp(13) = cp(13)
      strDataTemp(14) = strUserNum
      strDataTemp(20) = "N"
      strDataTemp(26) = "N"
      strDataTemp(27) = txtCaseField(0)
      strDataTemp(32) = "N"
      strDataTemp(43) = cp(9)
      strDataTemp(44) = cp(44)
      strDataTemp(116) = cp(116)
      strDataTemp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數: 新增B類收文
      
      strTxt(iStep) = GetCPSQL(strDataTemp(), False)
      
      '911106 nick transation
      cnnConnection.Execute strTxt(iStep)
      
      iStep = iStep + 1
        
        'Add By Cheng 2003/09/16
        '若有ECP指定國家, 則新增案件進度檔資料
        If field(9) = EPC指定國家 And strCountry <> "" Then
            'Modify by Morgan 2006/12/25
            'If Not objPublicData.SaveCountry(1, intCaseKind, strDataTemp(1) & strDataTemp(2) & strDataTemp(3) & strDataTemp(4), strDataTemp(9), strCountry) Then
            If Not PUB_SaveCountry(1, intCaseKind, strDataTemp(1) & strDataTemp(2) & strDataTemp(3) & strDataTemp(4), strDataTemp(9), strCountry) Then
                GoTo CheckingErr
            End If
        End If
   
   Else
      If cp(10) = 延期 Then
         Dim bolDo As Boolean
         bolDo = True
         For i = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(i, 0) = "ˇ" Then
               strCF03 = grdDataList.TextMatrix(i, 11) 'Added by Morgan 2012/3/21
               stNP01 = grdDataList.TextMatrix(i, 7)
               'Add by Morgan 2011/4/22
               '要和 stNP01(cp43) 同步
               If grdDataList.TextMatrix(i, 10) = "0" Then
                  stNP22 = ""
               Else
                  stNP22 = grdDataList.TextMatrix(i, 10)
               End If
               
                'Modify By Cheng 2003/12/08
                '若本所期限非工作天則抓最近的工作天
'               strTxt(iStep) = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(TransDate(txtCaseField(1), 2)) & _
'                  ",NP09=" & CNULL(TransDate(txtCaseField(2), 2)) & " WHERE NP22 = " & grdDataList.TextMatrix(i, 10)

               'Add by Morgan 2009/12/24 +cp
               If Val(grdDataList.TextMatrix(i, 10)) = 0 Then
                  strTxt(iStep) = "UPDATE CASEPROGRESS SET CP06=" & CNULL(PUB_GetWorkDay1(TransDate(txtCaseField(1), 2), True)) & ", CP07=" & CNULL(TransDate(txtCaseField(2), 2), True) & " WHERE CP09='" & stNP01 & "'"
                  
               Else
               'end 2009/12/24
                  'Modify by Morgan 2006/1/24 加NP01
                  strTxt(iStep) = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(PUB_GetWorkDay1(TransDate(txtCaseField(1), 2), True)) & _
                     ",NP09=" & CNULL(TransDate(txtCaseField(2), 2), True) & " WHERE NP22 = " & grdDataList.TextMatrix(i, 10) & " and np01='" & grdDataList.TextMatrix(i, 7) & "'"
               End If
                '911106 nick transation
                
                cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
                cnnConnection.Execute strTxt(iStep), intI
                cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
                
               iStep = iStep + 1
               
               If bolDo Then
                  cnnConnection.Execute "delete datelimit where dl01='" & stNP01 & "' and dl02=" & DBDATE(txtCaseField(0))  'Added by Morgan 2016/8/17
                  
                  strExc(3) = TransDate(Replace(grdDataList.TextMatrix(i, 2), "/", ""), 2)
                  strExc(4) = TransDate(Replace(grdDataList.TextMatrix(i, 3), "/", ""), 2)
                  'Modify by Morgan 2004/9/15 DL01改放NP01
'                  strTxt(iStep) = "insert into datelimit (DL01,DL02,DL03,DL04,DL05,DL06) values (" + CNULL(CP(9)) + "," + CNULL(TransDate(txtCaseField(0), 2)) + "," + _
'                     CNULL(strExc(3)) + "," + CNULL(strExc(4)) + ",'2'," & grdDataList.TextMatrix(i, 10) & ")"
                  strTxt(iStep) = "insert into datelimit (DL01,DL02,DL03,DL04,DL05,DL06) values (" + CNULL(stNP01) + "," + CNULL(TransDate(txtCaseField(0), 2)) + "," + _
                     CNULL(strExc(3)) + "," + CNULL(strExc(4)) + ",'2'," & grdDataList.TextMatrix(i, 10) & ")"
                '911106 nick transation
                cnnConnection.Execute strTxt(iStep)
                     
                  iStep = iStep + 1
                  
                  bolDo = False
               End If
            End If
         Next
         
      End If
      
      'Modified by Morgan 2012/2/15 改呼叫共用函數
      'strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
      '   " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
      '   " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      'cp(45) = ""
      'If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then cp(45) = RsTemp.Fields("CP45")
      If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
      'end 2012/2/15
      
      cp(27) = txtCaseField(0)
      cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數: 先收文
      
      strTxt(iStep) = GetCPSQL(cp)
      
      '911106 nick transation
      cnnConnection.Execute strTxt(iStep)
      
      iStep = iStep + 1
      
      'Add by Morgan 2011/4/22
      If cp(10) = 延期 Then
         strExc(1) = ",CP30='" & stNP22 & "'"
      End If
      
      'Add by Morgan 2004/9/15 更新相關總收文號
      strSql = "Update CaseProgress Set CP43='" & stNP01 & "'" & strExc(1) & " Where CP09='" & cp(9) & "'"
      cnnConnection.Execute strSql, intI
      '2004/9/15 end
      
        'Add By Cheng 2003/09/16
        '若有ECP指定國家, 則新增案件進度檔資料
        If field(9) = EPC指定國家 And strCountry <> "" Then
            'Modify by Morgan 2006/12/25
            'If Not objPublicData.SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
            If Not PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
                GoTo CheckingErr
            End If
        End If
   
   End If
   
   '911106 nick transation
   'FormSave = objLawDll.ExecSQL(iStep - 1, strTxt())

   ' 90.12.05 modify by louis 若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   If FormSave = True Then
      'Modified by Morgan 2012/3/21 若案件性質為延期則要用被延期的案件性質判斷,發文日要用延期日,收文號要用延期的收文號
      'Dim strCF23 As String
      'Dim strNPDate As String
      'Dim strNPSerial As String
      'If IsExistCasefee(cp(1), field(9), cp(10), strCF23) Then
      '   strNPDate = DBDATE(DateAdd("d", Val(strCF23), ChangeWStringToWDateString(DBDATE(cp(27)))))
      '   strNPSerial = InsertNextProgress_997(cp(9), cp(1), cp(2), cp(3), cp(4), strNPDate)
      'End If
      If cp(10) = 延期 Then
         PUB_SetArriveDate cp(9), strCF03
         m_CP09 = cp(9) 'Added by Morgan 2013/9/24
      Else
         PUB_SetArriveDate strDataTemp(9), cp(10)
         m_CP09 = strDataTemp(9) 'Added by Morgan 2013/9/24
      End If
      'end 2012/3/21
      
      'Added by Morgan 2015/8/7
      '提申管制
      If cp(10) = 延期 Then
         strCP09 = cp(9)
         strCP10 = strCF03
      Else
         strCP09 = strDataTemp(9)
         strCP10 = cp(10)
      End If
      PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), cp(7), strCP09, strCP10, txtCaseField(0), field(9)
      'end 2015/8/7
   End If
   '911106 nick transaction
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/8/22 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      'Modified by Morgan 2018/9/6 +有工程師的指示信
      If txtCaseField(17) <> "N" Or m_bolEngLetter = True Then
         If m_bolEngLetter Then
            strLetterJudge = strUserNum
         Else
            strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), 延期, field(9), strCP10)
         End If
         m_strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), cp(4), 延期, field(11), IIf(cp(10) = 延期, cp(45), strDataTemp(45)), field(9))
         PUB_AddAppForm strCP09, True, strLetterJudge, m_strSubject
         m_strAF01 = strCP09
      End If
   End If
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(3) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("1", cp(1), 延期, field(9), strCP10)
         PUB_AddLetterProgress strCP09, 0, True, strLetterJudge, False, field(26), 延期, field(75)
         m_strLD18 = strCP09
      End If
   End If
   'end 2018/8/22
   
   cnnConnection.CommitTrans
Exit Function
CheckingErr:
    FormSave = False
     cnnConnection.RollbackTrans

End Function

Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String
Dim adoRecord As Object, strSameName As String

On Error GoTo ErrHnd
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.ReadAllData(frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.Row, 5), cp(), field(), intCaseKind, intPWhere) Then
ReDim cp(TF_CP) As String
cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)
If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   
   lblCaseField(0) = cp(1) + " - " + cp(2) + _
      IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
      IIf(cp(4) = "00", "", " - " + cp(4))
   lblCaseField(2) = cp(9)
   lblCaseField(3) = TransDate(cp(5), 1)
   lblCaseField(4) = cp(10)
   lblCaseField(5) = TransDate(cp(6), 1)
   lblCaseField(6) = TransDate(cp(7), 1)
   SetNameToCombo cboPatentName, field(5), field(6), field(7)
   lblCaseField(1) = field(11)
   If cp(10) = 延期 And intWhereComeFrom = 2 Then
      'Modify by Morgan 2009/12/24
      'GetCaseDeadLineData grdDataList, intLastRow, cp(1), cp(2), cp(3), cp(4), , True
      GetGrid grdDataList, intLastRow, cp(1), cp(2), cp(3), cp(4)
   Else
      'Modify by Morgan 2009/12/24
      'GetCaseDeadLineData grdDataList, intLastRow, cp(1), cp(2), cp(3), "##", , True
      GetGrid grdDataList, intLastRow, cp(1), cp(2), cp(3), "##"
   End If
   
   '2007/4/23 MODIFY BY SONIA 改與其他發文畫面相同之語法
'   strExc(0) = "SELECT DISTINCT CP44 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4))
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      Do While Not RsTemp.EOF
'         If Not IsNull(RsTemp.Fields(0).Value) Then Combo1.AddItem RsTemp.Fields(0).Value
'         RsTemp.MoveNext
'      Loop
'      Combo1 = Combo1.List(0)
'      CheckKeyIn 4
'   End If
'
   'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
   If cp(31) = "Y" Then
      AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
      If Combo1 <> "" Then CheckKeyIn 4
      
   Else '非新案照原本
      If ClsPDSelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc", adoRecord) Then
         Do While adoRecord.EOF = False
            If IsNull(adoRecord.Fields(0).Value) = False Then
               If strSameName <> adoRecord.Fields(0).Value Then
                  Combo1.AddItem adoRecord.Fields(0).Value
                  strSameName = adoRecord.Fields(0).Value
               End If
            End If
            adoRecord.MoveNext
         Loop
         Combo1 = Combo1.List(0)
      End If
      '2007/4/23 END
      'Added by Morgan 2023/10/30 已有設定時不必再重新設定(IDS分案會先設,且抓預設代理人時也會剔除)
      If cp(44) <> "" Then
         Combo1 = cp(44) & IIf(cp(116) <> "", "-" & cp(116), "")
         CheckKeyIn 4
      Else
      'end 2023/10/30
      
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCasePreAgent(cp(), strTemp) Then
         If ClsPDGetCasePreAgent(cp(), strTemp) Then
            Combo1 = strTemp
            CheckKeyIn 4
         End If
         
      End If 'Added by Morgan 2023/10/30
   End If
   'end 2016/10/27
   
'   If objPublicData.GetCasePreAgent(cp(), strTemp) Then
'       Combo1 = strTemp
'       CheckKeyIn 4
'    End If
    
'   If objPublicData.GetCaseDelayDay(cp(1), field(9), cp(10), strTemp) Then
'      txtCaseField(2) = TransDate(CompDate(2, Val(strTemp), strSrvDate(2)), 1)
      
      'Add By Cheng 2002/02/18
      '若是從延期按鈕進入, 則延期後法定期限=原法定期限+延期天數或延期月數, 延期後本所期限=延期後法定期限-14天
'      If Me.intWhereComeFrom = 1 Then
'         strExc(0) = "SELECT CF22,CF25 FROM CaseFee WHERE CF01='" & cp(1) & "' AND CF02='" & field(9) & "' AND CF03='" & cp(10) & "'"
'         intI = 1
'         Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If Len("" & rsTemp.Fields(0).Value) > 0 Then
'               If Len("" & lblCaseField(6).Caption) > 0 Then
                  '延期後法定期限
'                  txtCaseField(2).Text = TransDate(DateSerial(Year(lblCaseField(6).Caption), Month(lblCaseField(6).Caption), Day(lblCaseField(6).Caption) + rsTemp.Fields(0).Value), 1)
                  '延期後本所期限
'                  txtCaseField(1).Text = TransDate(DateSerial(Year(lblCaseField(6).Caption), Month(lblCaseField(6).Caption), Day(lblCaseField(6).Caption) + rsTemp.Fields(0).Value - 14), 1)
'               End If
'            ElseIf Len("" & rsTemp.Fields(1).Value) > 0 Then
'               If Len("" & lblCaseField(6).Caption) > 0 Then
                  '延期後法定期限
'                  txtCaseField(2).Text = TransDate(ChangeWDateStringToWString(DateSerial(Year(lblCaseField(6).Caption), Month(lblCaseField(6).Caption) + rsTemp.Fields(1).Value, Day(lblCaseField(6).Caption))), 1)
                  '延期後本所期限
'                  txtCaseField(1).Text = TransDate(ChangeWDateStringToWString(DateSerial(Year(lblCaseField(6).Caption), Month(lblCaseField(6).Caption) + rsTemp.Fields(1).Value, Day(lblCaseField(6).Caption) - 14)), 1)
'               End If
'            End If
'         End If
'      End If
      
'   End If
    'Add By Cheng 2003/09/16
    '讀取ECP指定國家
    'edit by nickc 2007/02/02 不用 dll 了
    'If field(9) = EPC指定國家 Then objPublicData.ReadCountry intCaseKind, cp(), strCountry, True, False
    If field(9) = EPC指定國家 Then ClsPDReadCountry intCaseKind, cp(), strCountry, True, False
End If

'Added by Lydia 2021/05/25
txtCP113 = ""
If cp(113) <> "" Then txtCP113 = cp(113)
'end 2021/05/25
      
Screen.MousePointer = varSaveCursor
'txtCaseField(4) = "Y"
Exit Sub
ErrHnd:
ErrorMsg
Screen.MousePointer = varSaveCursor
Resume
End Sub

Private Sub Combo1_Change()
'Remove by Morgan 2010/8/19
' Dim strAgentName As String
'   lblAgent.Caption = ""
'   If Combo1.Text <> "" Then
'      'edit by nickc 2007/02/02 不用 dll 了
'      'If objPublicData.GetAgent(Combo1, strAgentName) = True Then
'      If ClsPDGetAgent(Combo1, strAgentName) = True Then
'         lblAgent.Caption = strAgentName
'      End If
'   End If
End Sub

Private Sub Combo1_Click()
'Remove by Morgan 2010/8/19
'   Combo1_Change
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub


Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo1.Text <> "" Then
      If CheckKeyIn(4) = -1 Then
         Cancel = True
      End If
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         'Add by Morgan 2008/2/18 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/18
         
         If PUB_CheckStatus(strNo) = False Then Cancel = True
      End If
      
      If Cancel Then Combo1.SetFocus
   End If
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
   
   ReadAllData
   blnOKtoShow = True
   txtCaseField(0) = strSrvDate(2)
   'Remove by Morgan 2009/12/24 改輸月份計算
   'If IsEmptyText(txtCaseField(0)) = False Then
   '   CaculateNP08NP09
   'End If
   txtCaseField(19) = "N" 'Added by Morgan 2018/10/22 預設不印傳真封面--慧汶
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If intLeaveKind = 1 Then
      frm050102_1.Show
   ElseIf intLeaveKind = 0 Then
      Unload frm050102_1
   End If
   ShowEditForm 'Added by Morgan 2018/8/22
   
   Set frm050102_2 = Nothing
End Sub

Private Sub lblCaseField_Change(Index As Integer)
 Dim strTemp As String
   Select Case Index
      Case 4
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp) Then
         If ClsPDGetCaseProperty(cp(1), lblCaseField(Index), strTemp) Then
            lblCasePropertyName = strTemp
         End If
   End Select
End Sub

Private Sub grdDataList_GotFocus()
   GridGotFocus grdDataList
End Sub

Private Sub grdDataList_LostFocus()
   GridLostFocus grdDataList
End Sub

Private Sub grdDataList_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then GrdDataList_Click
End Sub

Private Sub GrdDataList_Click()
   intLastRow = grdDataList.row
   If grdDataList.TextMatrix(intLastRow, 0) = "ˇ" Then
      grdDataList.TextMatrix(intLastRow, 0) = ""
   Else
      grdDataList.TextMatrix(intLastRow, 0) = "ˇ"
   End If
   
   'Add by Morgan 2009/12/24
   If grdDataList.TextMatrix(intLastRow, 0) = "ˇ" Then
      strExc(2) = Replace(grdDataList.TextMatrix(intLastRow, 3), "/", "")
      If DBDATE(strExc(2)) <> cp(7) Then
         MsgBox "所點選案件性質的法定期限與延期程序不同，不可點選！"
         grdDataList.TextMatrix(intLastRow, 0) = ""
      End If
   End If
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 8
      blnOKtoShow = True
   End If
End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   CloseIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Val(Text9) > 0 Then
      strExc(1) = field(1)
      strExc(2) = field(9)
      strExc(3) = CompDate("1", Val(Text9), cp(7))
      GetCtrlDT strExc
      txtCaseField(2) = TransDate(strExc(3), 1)
      txtCaseField(1) = TransDate(PUB_GetWorkDay1(strExc(0), True), 1)
      'Modified by Lydia 2015/12/01
      'US107Check 'Added by Morgan 2012/3/21
      If US107Check = False Then
         Text9_GotFocus
         Text9.SetFocus
         Cancel = True
      End If
      'end 2015/12/01
   End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      'N
      Case 3, 17, 19
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
         End If
      'Y
      Case 4, 16
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
         End If
   End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = False Then
   Cancel = True
End If
If Cancel Then txtCaseField_GotFocus (Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, bolIsChina As Boolean, strCusTemp As String

Select Case intIndex
             Case 0
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           'Remove by Morgan 2009/12/24 改輸月份計算
                           'If txtCaseField(1) = "" And txtCaseField(2) = "" Then
                           '   CaculateNP08NP09
                           'End If
                           CheckKeyIn = 1
                        End If
             Case 1 '延期後本所期限
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           'If txtCaseField(1) <= txtCaseField(2) Or txtCaseField(2) = "" Then
                           '   If CheckReKey(txtCaseField(intIndex)) Then
                                 CheckKeyIn = 1
                           '   End If
                           'Else
                           '   ShowMsg MsgText(9210)
                           '   CheckKeyIn = 0
                           'End If
                            'Add By Cheng 2003/12/08
                            '若本所期限非工作天則直接調整至最近的工作天
                            Me.txtCaseField(intIndex).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(intIndex).Text, True), 1)
                        End If
             Case 2
                        If txtCaseField(intIndex) <> "" Then
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              'Modify by Morgan 2010/8/11 百年蟲
                              'If txtCaseField(1) <= txtCaseField(2) Then
                              If Val(txtCaseField(1)) <= Val(txtCaseField(2)) Then
                                 If CheckReKey(txtCaseField(intIndex)) Then
                                    CheckKeyIn = 1
                                 End If
                              Else
                                 ShowMsg MsgText(9210)
                                 CheckKeyIn = 0
                              End If
                           End If
                        ElseIf txtCaseField(1) <> "" Then
                           ShowMsg MsgText(1033)
                        Else
                           CheckKeyIn = 1
                        End If
             Case 3
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 4 '代理人
                        strCusTemp = Combo1
                        CheckKeyIn = 0
                        lblAgent = ""
                        If strCusTemp = "" Then
                           MsgBox "代理人欄不可空白!!!", vbExclamation
                        
                        'Add by Morgan 2008/2/20 加判斷是否為聯絡人
                        ElseIf InStr(strCusTemp, "-") > 0 Then
                           If ClsPDGetContact(strCusTemp, strTemp) Then
                              Combo1 = strCusTemp
                              lblAgent.Caption = strTemp
                              CheckKeyIn = 1
                           End If
                              
                        ElseIf ClsPDGetAgent(strCusTemp, strTemp) Then
                           Combo1 = strCusTemp
                           lblAgent.Caption = strTemp
                           CheckKeyIn = 1
                           
                        End If
                        
             Case Else
                        CheckKeyIn = 1
End Select
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   TextInverse txtCaseField(Index)
   '儲存未修改前之值至Tag中,供再確認時使用
   txtCaseField(Index).Tag = txtCaseField(Index)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/12/02
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim bFind As Boolean
Dim nIndex  As Integer

TxtValidate = False

'Added by Morgan 2012/3/21
If US107Check = False Then
   Text9.SetFocus
   Text9_GotFocus
   Exit Function
End If
'end 2012/3/21

   'add by nickc 2008/05/01
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
    'Add By Cheng 2002/12/02
   ' 當案件性質為延期時, 未收文期限至少要選取一筆
   If cp(10) = "404" Then
      If Me.grdDataList.Rows <= 1 Then
         strTit = "檢核資料"
         strMsg = "未收文期限無資料, 無法執行延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       Exit Function
      End If
      
      bFind = False
      For nIndex = 1 To Me.grdDataList.Rows - 1
         'Modify by Morgan 2004/9/1 bug
         'If Me.grdDataList.TextMatrix(nIndex, 0) = "V" Then
         If Me.grdDataList.TextMatrix(nIndex, 0) = "ˇ" Then
            bFind = True
            Exit For
         End If
      Next nIndex
      If bFind = False Then
         strTit = "檢核資料"
         strMsg = "請先選取未收文期限的資料來做延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Exit Function
      End If
   End If

For Each objTxt In Me.txtCaseField
   If objTxt.Enabled = True Then
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2004/9/14
If Combo1.Enabled = True Then
   Cancel = False
   Combo1_Validate Cancel
   If Cancel = True Then
      Combo1.SetFocus
      Exit Function
   End If
End If

'Added by Morgan 2018/9/6
'若系統不出指示信時判斷是否有工程師的指示信要寄送
m_bolEngLetter = False
If txtCaseField(17) = "N" And cp(10) = "404" Then
   If PUB_EngLtrChk(cp(9), txtCaseField(0).Text, m_bolEngLetter) = False Then
      Exit Function
   End If
End If
'end 2018/9/6

'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
   If cp(10) = "404" And cp(9) < "B" And Left(cp(12), 1) <> "F" Then
      If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
         Exit Function
      End If
   End If
End If
'end 2018/9/12

'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
      txtCP113.SetFocus
      txtCP113_GotFocus
      Exit Function
End If
'end 2021/05/25
   
TxtValidate = True
End Function

' 計算本所期限及法定期限
Private Sub CaculateNP08NP09()
   If IsEmptyText(txtCaseField(0)) = False Then
      strExc(0) = TransDate(txtCaseField(0).Text, 2)
      'edit by nickc 2007/02/05 不用 dll 了
      'If objLawDll.GetCaseFeeDelay(field(1), field(9), cp(10), strExc) Then
      If ClsLawGetCaseFeeDelay(field(1), field(9), cp(10), strExc) Then
         txtCaseField(2) = TransDate(strExc(1), 1)
         txtCaseField(1) = TransDate(strExc(2), 1)
        'Add By Cheng 2003/12/08
        '本所期限若非工作天則抓最近工作天
        Me.txtCaseField(1).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(1).Text, True), 1)
      End If
   End If
End Sub

'Add by Morgan 2009/12/24
Private Sub GetGrid(ByRef grdTemp As MSHFlexGrid, ByRef intLastRow As Integer, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String, Optional strCode5 As String)
Dim varSaveCursor, varGridWidth() As Variant
Dim strSql As String

   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   varGridWidth = Array(300, 1350, 900, 900, 1200, 1200, 1200, 1000, 0, 0, 0, 0, 0)
   SetGridDataListWidth grdTemp, varGridWidth()
   
   '+AB類未發文未取消收文的程序,且下一程序要排除程序管制的案件性質
   If strCode1 = "CFP" Then
      strSql = "select '' ˇ,decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,"
   Else
      strSql = "select '' ˇ,decode(sp09," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,"
   End If
   'Modified by Morgan 2012/3/21 +NP01(CP43)
   strSql = strSql + SQLDate("np08") & " 本所期限," & SQLDate("np09") & " 法定期限,np13 機關文號,np14 相關人," & SQLDate("np11") & " 解除期限日期, np01 總收文號, dbms_rowid.rowid_to_restricted(nextprogress.rowid,0), NP15 備註, NP22 序號, NP07,NP08,NP01 "
   If strCode1 = "CFP" Then
      strSql = strSql + "from nextprogress,patent,casepropertymap where pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
   Else
      strSql = strSql + "from nextprogress,servicepractice,casepropertymap where sp01(+)=np02 and sp02(+)=np03 and sp03(+)=np04 and sp04(+)=np05"
   End If
   strSql = strSql + " and np02=cpm01(+) and np07=cpm02(+) and np06 is null"
   'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
   strSql = strSql + " and np02=" + CNULL(strCode1) + " and np03=" + CNULL(strCode2) + " and np04=" + CNULL(strCode3) + " and np05=" + CNULL(strCode4) & strNpSqlOfNoSalesDuty
   
   If cp(10) = "404" Then
      strSql = strSql & " union SELECT '',DECODE(PA09,'" & 台灣國家代號 & "',CPM03,CPM04)" & _
         ",SQLDateT(CP06),SQLDateT(CP07),CP08,NVL(CP40,NVL(CP41,CP42)),''" & _
         ",CP09,DBMS_ROWID.ROWID_TO_RESTRICTED(CASEPROGRESS.RowID,0),CP64,0,CP10,CP06,CP43 FROM CASEPROGRESS,CASEPROPERTYMAP,PATENT" & _
         " WHERE " & ChgCaseprogress(strCode1 & strCode2 & strCode3 & strCode4) & " AND CP09<'C' and cp10<>'404' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND pa01(+)=CP01 and pa02(+)=CP02 and pa03(+)=CP03 and pa04(+)=CP04"
   End If
   
   Set grdTemp.Recordset = ClsPDReadRst(strSql)
   
   SetDataListVision grdTemp, True
   intLastRow = 0
   If grdTemp.Rows > 1 Then
      ShowBar grdTemp, intLastRow, grdTemp.Cols - 1
   End If
   Screen.MousePointer = varSaveCursor
End Sub
'Added by Morgan 2012/3/21
Private Function US107Check() As Boolean
   Dim stCP09 As String, stCP10 As String, ii As Integer
   Dim inD As Integer  'Added by Lydia 2015/12/01
   US107Check = True
   If txtCaseField(2) <> "" And field(9) = "101" Then
      'Modified by Lydia 2015/12/01 美國答辯(107),RCE(424)之延期發文時請檢查延期後之法定期限不可超過原法定期限+3個月
      '                             美國選取(208)之延期發文時請檢查延期後之法定期限不可超過原法定期限+4個月
      'If cp(10) = "107" And cp(43) <> "" Then
      'Modified by Morgan 2016/3/3 +126 期末拋棄
      'Modified by Lydia 2016/08/29 +438 再考量試行計畫(AFCP2.0)
      If InStr("107,424,208,126,438", cp(10)) > 0 And cp(43) <> "" Then
         stCP09 = cp(43)
         stCP10 = cp(10) 'Added by Lydia 2015/12/01
      'Modified by Lydia 2015/12/01
      ElseIf cp(10) = 延期 Then
         For ii = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(ii, 0) = "ˇ" Then
               stCP10 = grdDataList.TextMatrix(ii, 11)
               'Modified by Lydia 2015/12/01
               'If stCP10 = "107" Then
               'Modified by Morgan 2016/3/3 +126 期末拋棄
               'Modified by Lydia 2016/08/29 +438 再考量試行計畫(AFCP2.0)
               If InStr("107,424,208,126,438", stCP10) > 0 Then
                  stCP09 = grdDataList.TextMatrix(ii, 13)
                  Exit For
               End If
            End If
         Next
      End If

      If stCP09 <> "" Then
         'Added by Lydia 2015/12/01
         'Modified by Morgan 2016/3/3 +126 期末拋棄
         'Modified by Lydia 2016/08/29 +438 再考量試行計畫(AFCP2.0)
         If InStr("107,424,126,438", stCP10) > 0 Then
            inD = 3
         ElseIf stCP10 = "208" Then
            inD = 4
         End If
         'end 2015/12/01
         strExc(0) = "select cp07 from caseprogress where cp09='" & stCP09 & "' and cp07>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Lydia 2015/12/01
            'strExc(1) = CompDate(1, 3, RsTemp(0))
            strExc(1) = CompDate(1, inD, RsTemp(0))
            If DBDATE(txtCaseField(2)) > strExc(1) Then
               'Modified by Lydia 2015/12/01
               'MsgBox "美國答辯延期不可超過來函法限+3個月!!"
               MsgBox "美國" & lblCasePropertyName & "延期不可超過來函法限+" & inD & "個月!!"
               US107Check = False
            End If
         End If
      End If
   End If
End Function

'Added by Morgan 2013/9/24
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 3) As String, ii As Integer
   EndLetter ET01, m_CP09, ET03, strUserNum
   
   ii = 0
   
   If ET03 = "99" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
         "','傳真頁數','2')"
   Else
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
         "','法定期限','" & cp(7) & "')"
      
      If Text9 <> "" Then 'Added by Morgan 2023/8/18
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
            "','延期月數','" & Num2Eng(Text9) & "')"
      End If
   End If
      
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Function Num2Eng(iNum As Integer) As String
   Select Case iNum
   Case 1
      Num2Eng = "one"
   Case 2
      Num2Eng = "two"
   Case 3
      Num2Eng = "three"
   Case 4
      Num2Eng = "four"
   Case 5
      Num2Eng = "five"
   Case 6
      Num2Eng = "six"
   Case 7
      Num2Eng = "seven"
   Case 8
      Num2Eng = "eight"
   Case 9
      Num2Eng = "nine"
   Case 10
      Num2Eng = "ten"
   Case 11
      Num2Eng = "eleven"
   Case 12
      Num2Eng = "twelve"
   Case Else
      Num2Eng = iNum
   End Select
   
   If iNum = 1 Then
      Num2Eng = Num2Eng & " month"
   Else
      Num2Eng = Num2Eng & " months"
   End If
End Function

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
