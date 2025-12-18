VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010306_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-退費"
   ClientHeight    =   5270
   ClientLeft      =   180
   ClientTop       =   1290
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5270
   ScaleWidth      =   8040
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   6900
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   49
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   50
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "繳費記錄(&R)"
      Height          =   400
      Index           =   3
      Left            =   3420
      TabIndex        =   47
      Top             =   90
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5556
      TabIndex        =   46
      Top             =   90
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4643
      TabIndex        =   45
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6780
      TabIndex        =   44
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   4590
      MaxLength       =   7
      TabIndex        =   5
      Top             =   4650
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   4590
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   1290
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2655
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4170
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2655
      Width           =   300
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1152
      MaxLength       =   3
      TabIndex        =   14
      Top             =   600
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1704
      MaxLength       =   6
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2556
      MaxLength       =   1
      TabIndex        =   12
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2796
      MaxLength       =   2
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   4590
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2985
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1290
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2985
      Width           =   375
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "(1~3選項請至下一程序做相對應之修改)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   8
      Left            =   1890
      TabIndex        =   51
      Top             =   3750
      Width           =   3070
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   15
      Top             =   1230
      Width           =   5715
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "10081;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   2655
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "7. 其他退費"
      Height          =   180
      Index           =   7
      Left            =   1880
      TabIndex        =   43
      Top             =   4650
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "6. 申請書及摘要均附英文資料，減免規費800元整"
      Height          =   180
      Index           =   6
      Left            =   1880
      TabIndex        =   42
      Top             =   4410
      Width           =   3870
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額:"
      Height          =   180
      Left            =   3780
      TabIndex        =   41
      Top             =   4680
      Width           =   770
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "5. 實審、再審、續行母案再審退費"
      Height          =   180
      Index           =   5
      Left            =   1880
      TabIndex        =   40
      Top             =   4190
      Width           =   3330
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "4. 面詢退費"
      Height          =   180
      Index           =   4
      Left            =   1880
      TabIndex        =   39
      Top             =   3980
      Width           =   900
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5520
      TabIndex        =   38
      Top             =   2700
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   60
      X2              =   8000
      Y1              =   2500
      Y2              =   2500
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   8000
      Y1              =   2530
      Y2              =   2530
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "退費金額:"
      Height          =   180
      Index           =   1
      Left            =   3780
      TabIndex        =   37
      Top             =   4970
      Width           =   770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收據編號:"
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   36
      Top             =   4970
      Width           =   770
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "3. 移做次年 ( 至繳費紀錄加一年 )"
      Height          =   180
      Index           =   3
      Left            =   1875
      TabIndex        =   35
      Top             =   3525
      Width           =   2595
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "2. 取消繳納 ( 至繳費紀錄減一年 )"
      Height          =   180
      Index           =   2
      Left            =   1875
      TabIndex        =   34
      Top             =   3300
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   33
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   2490
      TabIndex        =   32
      Top             =   2700
      Width           =   2880
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3570
      TabIndex        =   31
      Top             =   630
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3570
      TabIndex        =   30
      Top             =   2130
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   150
      TabIndex        =   29
      Top             =   2130
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4380
      TabIndex        =   28
      Top             =   630
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3540
      TabIndex        =   27
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   330
      TabIndex        =   26
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   330
      TabIndex        =   25
      Top             =   630
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   330
      TabIndex        =   24
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3570
      TabIndex        =   23
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   330
      TabIndex        =   22
      Top             =   1290
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1140
      TabIndex        =   21
      Top             =   930
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4380
      TabIndex        =   20
      Top             =   930
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1140
      TabIndex        =   19
      Top             =   1800
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4380
      TabIndex        =   18
      Top             =   1800
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1140
      TabIndex        =   17
      Top             =   2130
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4380
      TabIndex        =   16
      Top             =   2130
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label22 
      Caption         =   "下次繳費日:"
      Height          =   255
      Left            =   3570
      TabIndex        =   10
      Top             =   3015
      Width           =   975
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請書類別:"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   9
      Top             =   3015
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "1. 重複繳納退費"
      Height          =   180
      Index           =   0
      Left            =   1875
      TabIndex        =   8
      Top             =   3060
      Width           =   1260
   End
End
Attribute VB_Name = "frm06010306_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/5 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, cp() As String, m_CP110 As String, m_AgentName As String

Dim intWhere As Integer
Dim strYear As String, m_CP43 As String
'Add by Morgan 2006/1/12
Dim m_CP20 As String, m_CP60 As String, m_bUpdCP20 As Boolean
Dim m_404CP27 As String '延期發文日
Dim stRefCP10 As String 'Add by Morgan 2009/9/11
Public m_CP118isY As Boolean 'Add By Sindy 2019/1/2 是否為電子送件申請書:True.是
Dim m_CaseNo As String 'Add By Sindy 2019/1/2
Dim m_SendDate As String, m_SendWord As String, m_SendNumber As String 'Add By Sindy 2019/7/23
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/27


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(5) As String, ii As Integer
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','年費收據編號'," & CNULL(Text8(0).Text) & ")"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','退費金額'," & CNULL(Text8(1).Text) & ")"
   If m_404CP27 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期發文日','" & m_404CP27 & "')"
   End If
   
   'Added by Morgan 2013/6/6 未收文延期退費要改抓
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " select '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','退費案件性質',cpm03" & _
      " from casepropertymap where cpm01='" & pa(1) & "' and cpm02='" & stRefCP10 & "'"
   'end 2013/6/6
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add By Sindy 2019/1/3
'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim varTmp As Variant
'Dim bol107 As Boolean 'Add By Sindy 2019/12/9
Dim strNP07 As String 'Add By Sindy 2022/11/24
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())

   '出名代理人
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
   '辦理依據
   If m_SendDate <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(m_SendDate) & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & m_SendWord & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & m_SendNumber & "')"
   End If
   
   'Add By Sindy 2019/7/23
   '依申請書類別產生不同的申請內容
   varTmp = Split(pa(72), ",")
   Select Case Text6
      Case "1" '重複繳納退費
         'Add By Sindy 2020/3/6 敏莉反應FCP-054180要改抓進度的繳費年度起迄
         strExc(0) = "select cp53,cp54 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
                     " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='" & 年費 & "'" & _
                     " and cp158>0 and cp159=0 order by cp27 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val("" & RsTemp.Fields("cp53")) = Val("" & RsTemp.Fields("cp54")) Then
               strTmp = Val("" & RsTemp.Fields("cp53"))
            ElseIf Val("" & RsTemp.Fields("cp54")) > Val("" & RsTemp.Fields("cp53")) Then
               strTmp = CStr("" & RsTemp.Fields("cp53")) & "-" & CStr("" & RsTemp.Fields("cp54"))
            Else
               strTmp = "  "
            End If
            strTmp = "1.退費事由：「因本案第" & strTmp & "年年費已繳納在案，附上收據，懇請　鈞局准予退還重複繳納之費用」。" & vbCrLf
         Else
            strTmp = "1.退費事由：「因本案第  年年費已繳納在案，附上收據，懇請　鈞局准予退還重複繳納之費用」。" & vbCrLf
         End If
'         If UBound(varTmp) >= 0 Then
'            strTmp = "1.退費事由：「因本案第" & varTmp(UBound(varTmp)) & "年年費已繳納在案，附上收據，懇請　鈞局准予退還重複繳納之費用」。" & vbCrLf
'         Else
'            strTmp = "1.退費事由：「因本案第  年年費已繳納在案，附上收據，懇請　鈞局准予退還重複繳納之費用」。" & vbCrLf
'         End If
         '2020/3/6 END
      Case "2" '取消繳納 (至繳費紀錄減一年)
         strExc(0) = "select cp27,cp10 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
                     " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in('" & 年費 & "')" & _
                     " and cp158>0 and cp159=0 order by cp27 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTmp = "1.本案於" & Val(Left(RsTemp.Fields("cp27"), 4)) - 1911 & "年" & Mid(RsTemp.Fields("cp27"), 5, 2) & "月" & Right(RsTemp.Fields("cp27"), 2) & "日"
         Else
            strTmp = "1.本案於年月日"
         End If
         If UBound(varTmp) >= 0 Then
            strTmp = strTmp & "繳納第" & varTmp(UBound(varTmp)) + 1 & "年年費，因申請人不擬續辦原因，故懇請　鈞局准予退費。" & vbCrLf
         Else
            strTmp = strTmp & "繳納第  年年費，因申請人不擬續辦原因，故懇請　鈞局准予退費。" & vbCrLf
         End If
      Case "3" '移做次年 (至繳費紀錄加一年)
         If UBound(varTmp) >= 0 Then
            'Modify By Sindy 2020/3/12 + ，且請一併將電子收據繳納年度改成第" & varTmp(UBound(varTmp)) & "年。
            strTmp = "因本案第" & varTmp(UBound(varTmp)) - 1 & "年年費已繳納在案之原因，懇請  鈞局准予將重複繳納之費用，移做繳交第" & varTmp(UBound(varTmp)) & "年年費，" & _
                     "且請一併將電子收據繳納年度改成第" & varTmp(UBound(varTmp)) & "年。"
         Else
            'Modify By Sindy 2020/3/12 + ，且請一併將電子收據繳納年度改成第  年。
            strTmp = "因本案第  年年費已繳納在案之原因，懇請  鈞局准予將重複繳納之費用，移做繳交第  年年費，" & _
                     "且請一併將電子收據繳納年度改成第  年。"
         End If
      Case "4" '面詢退費
         strExc(0) = "select cp27,cp10 from caseprogress" & _
                     " where cp09 in(select cp43 from caseprogress where cp09='" & cp(9) & "' and cp43 is not null)" & _
                     " and cp158>0 and cp159=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTmp = "1.退費事由：「本案於" & Val(Left(RsTemp.Fields("cp27"), 4)) - 1911 & "年" & Mid(RsTemp.Fields("cp27"), 5, 2) & "月" & Right(RsTemp.Fields("cp27"), 2) & "日申請面詢並繳納規費，但今接獲專利核准審定書，故檢附收據，辦理退費。」。" & vbCrLf
         Else
            strTmp = "1.退費事由：「本案於年月日申請面詢並繳納規費，但今接獲專利核准審定書，故檢附收據，辦理退費。」。" & vbCrLf
         End If
      Case "5" '實審、再審退費
         'Add By Sindy 2019/12/9
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','文件描述','辦理退費之電子收據')"
         strTmp = "1.請准予撤回本申請案。" & vbCrLf
         '2019/12/9 END
         '107.再審申請
         'Modify By Sindy 2022/11/24 + 續行母案再審435
         strExc(0) = "select cp27,cp10 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
                     " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in('" & 實體審查 & "','107','435')" & _
                     " and cp158>0 and cp159=0 order by cp27 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         bol107 = False
'         If intI = 1 Then
'            strTmp = "1.退費事由：「本案於" & Val(Left(RsTemp.Fields("cp27"), 4)) - 1911 & "年" & Mid(RsTemp.Fields("cp27"), 5, 2) & "月" & Right(RsTemp.Fields("cp27"), 2) & "日申請" & IIf(RsTemp.Fields("cp10") = 實體審查, "實體審查", "再審查") & "並繳納規費，但今客戶欲終止辦理並自請撤回此專利申請案，故依規費法第十九條檢附收據，辦理退費」。" & vbCrLf
'         Else
'            strTmp = "1.退費事由：「本案於年月日申請實體審查/再審查並繳納規費，但今客戶欲終止辦理並自請撤回此專利申請案，故依規費法第十九條檢附收據，辦理退費」。" & vbCrLf
'         End If
         If intI = 1 Then
            'Modify By Sindy 2022/11/24
            strNP07 = "" & RsTemp.Fields("cp10")
'            If RsTemp.Fields("cp10") = "107" Then
'               bol107 = True
'            End If
            '2022/11/24
         End If
      Case "6" '申請書及摘要均附英文資料，減免規費800元整
         strTmp = "1.退費事由：「本案所檢附之申請書中發明名稱、申請人姓名或名稱、發明人姓名及說明書摘要同時附有英文翻譯，申請費得減收新台幣800元，故檢還收據辦理退費。」。" & vbCrLf
      Case "7" '其他退費
         strTmp = "1.退費事由：「」。" & vbCrLf
   End Select
   If Text6 <> "3" Then
      If Text6 = "5" Then
         'Modify By Sindy 2020/3/16 Ex:FCP-055314:代辦退費-5. 實審、再審退費(專利申請案撤回申請書)
         'Modified by Lydia 2020/03/31 改模組A0802Query => CompNameQuery
         'Modify By Sindy 2022/11/24 Mark:IIf(bol107 = True, "再審查", "實體審查")
         strExc(10) = ""
         If pa(9) = "000" Then
            Call ClsPDGetCasePropertyL(1, pa(1), strNP07, strExc(10))
         Else
            Call ClsPDGetCasePropertyL(2, pa(1), strNP07, strExc(10))
         End If
         strTmp = strTmp & _
                  "2.退還本案已繳" & strExc(10) & "規費「" & Val(Text8(1)) & "」元。" & vbCrLf & _
                  "3.檢還之國庫支票抬頭請開立：「" & CompNameQuery("2") & "」。" & vbCrLf & _
                  "4.本案收據正本已遺失，檢附[收據無法檢還原因切結書]1份。"
      Else
         'Modify By Sindy 2020/3/6 Ex:FCP-054180:代辦退費-1. 重複繳納退費
         'Modified by Lydia 2020/03/31 改模組A0802Query => CompNameQuery
         strTmp = strTmp & _
                  "2.退還本案已繳規費新台幣「" & Val(Text8(1)) & "」元。（收據號碼：" & Text8(0) & "）。" & vbCrLf & _
                  "3.檢還之國庫支票抬頭請開立：「" & CompNameQuery("2") & "」。"
      End If
   End If
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','" & strTmp & "')"
   '2019/7/23 END
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-電子收據','" & m_CaseNo & ".ATT.pdf')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

'Add by Morgan 2005/8/8
Private Function FormSave() As Boolean
Dim stCP64 As String
Dim strCon As String
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   'Add by Morgan 2009/10/15
   '儲存退費金額
   If Val(Text8(1)) > 0 Then
'      strSql = "update caseprogress set cp19=" & Val(Text8(1)) & " where cp09='" & strReceiveNo & "'"
'      cnnConnection.Execute strSql, intI
      strCon = strCon & ",cp19=" & Val(Text8(1))
   End If
   '2009/10/15 END
   
   If m_CP118isY = True Then
      If m_CP118isY = True Then
         cp(118) = "A"
      Else
         cp(118) = ""
      End If
      strCon = strCon & ",cp118=" & CNULL(cp(118))
   End If
   strCon = strCon & ",cp84=" & Val(txtCP84) '發文規費
   
'   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(m_CP110) & strCon & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql
'   End If
   
   'Modify by Morgan 2007/2/7 領證繳年費退費不更新NP
   'If strYear <> Text9 Then
   If Text9 <> "" And strYear <> Text9 Then
   'End 2007/2/7
      strExc(2) = TransDate(Text9.Text, 2)
      
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/29
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   strExc(3) = PUB_GetOurDeadline(strExc(2))
      'Else
      ''end 2014/10/29
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
         strExc(3) = PUB_GetFCPOurDeadline(strExc(2), 2, , m_pAgreeOnDate)
      Else
      'end 2019/7/11
         strExc(3) = CompDate(2, -2, strExc(2))
      End If 'Added by Morgan 2019/7/11
      'End If
      'end 2014/11/20
      'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
      strSql = "UPDATE NEXTPROGRESS SET NP08=" & strExc(3) & "," & _
         "NP09=" & strExc(2) & ",NP23=" & CNULL(DBDATE(m_pAgreeOnDate)) & " WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & 年費 & " And NP09 = ( Select Max(NP09) From NextProgress Where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " And NP07 =" & 年費 & " )"
      cnnConnection.Execute strSql
   End If
   
   '申請書類別
   Select Case Text6.Text
      Case "1"
         stCP64 = ",重複繳納退費"
      Case "2"
         stCP64 = ",取消繳納"
      Case "3"
         stCP64 = ",移做次年"
      Case "4"
         stCP64 = ",面詢退費"
      Case "5"
         stCP64 = ",實審、再審退費"
      'Add By Sindy 2019/7/23
      Case "6"
         stCP64 = ",申請書及摘要均附英文資料，減免規費800元整"
      '2019/7/23 END
   End Select
   If stCP64 <> "" Then
      strSql = "Update caseprogress set cp64=cp64||'" & ChgSQL(stCP64) & "' where cp09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   '若未請款則自動上不請款
   If m_bUpdCP20 = True Then
      strSql = "Update CaseProgress Set CP20='N' Where CP09='" & m_CP43 & "' and cp20 is null and cp60 is null"
      cnnConnection.Execute strSql
   End If
         
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim strFolder As String, strFileName As String 'Add By Sindy 2019/1/3
   
   Select Case Index
      Case 0
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         If Text6 <> "3" And Text8(0) = "" Then
            MsgBox "請輸入收據編號 !", vbCritical
            Text8(0).SetFocus
            Exit Sub
         End If
         If Text8(1) = "" Then
            MsgBox "請輸入退費金額 !", vbCritical
            Text8(1).SetFocus
            Exit Sub
         End If
         
         '相關收文若未請款則自動上不請款
         m_bUpdCP20 = False
         If (Text6 = "1" Or Text6 = "2") And m_CP20 = Empty And m_CP60 = Empty Then
            m_bUpdCP20 = True
            MsgBox "相關收文號【" & m_CP43 & "】將被更新為不請款！", vbExclamation
         End If
         'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
         If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
             MsgBox MsgText(1111), vbInformation
             If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                 Exit Sub
             End If
         End If
         'end 2020/02/17
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Add by Sindy 2019/1/2 電子送件申請書
         If m_CP118isY = True Then
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            'If Pub_StrUserSt03 = "M51" Then
            If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
               strFolder = PUB_Getdesktop
            Else
               strFolder = FCP電子送件檔案存放路徑
            End If
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            '1.基本資料
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
            '2.申請書
            'Add By Sindy 2019/7/23
            '依申請書類別產生不同的申請書
            Select Case Text6
               '2.取消繳納 (至繳費紀錄減一年)
               '3.移做次年 (至繳費紀錄加一年)
               Case "2", "3"
                  If StartLetter2("01", "01") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "一般事項申復申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
               'Add By Sindy 2019/12/9
               Case "5" '5.實審、再審退費
                  If StartLetter2("01", "10") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "10", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利申請案撤回申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
               Case Else
            '2019/7/23 END
                  If StartLetter2("01", "00") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "00", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "專利規費退費申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
            End Select
            
         Else
         '2019/1/2 END
            If Text7 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            
            Select Case Text6
               Case "1" '重覆繳納退費 1
                  If Not IsNull(m_CP43) And m_CP43 > "C" Then '有來函
                     strTmp = "01"
                  Else
                     strTmp = "04"
                  End If
               Case "2" '取消繳納 2
                  'Modify by Morgan 2007/2/7 加判斷領證繳年費退費
                  'strTmp = "02"
                  If Text9 = "" Then
                     strTmp = "05"
                  Else
                     strTmp = "02"
                  End If
                  'End 2007/2/7
               Case "3" '取消繳納移做次年 3
                  If Not IsNull(m_CP43) And m_CP43 > "C" Then '有來函
                     strTmp = "03"
                  Else
                     strTmp = "06"
                  End If
               Case "4"
                  strTmp = "07"
               Case "5"
                  strTmp = "08"
            End Select
            
            StartLetter "01", strTmp
            strLetterDate = Text5.Text
            NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum
         End If
         
         frm060103_1.Show
         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
         frm060103_1.ClearForm
         Unload Me
      Case 1
         frm060103_1.Show
         Unload Me
      Case 2
         Unload frm060103_1
         Unload Me
      Case 3
         Set frm060104_b.oParent = Me 'Add by Morgan 2011/10/5
         frm060104_b.LoadMe pa(1), pa(2), pa(3), pa(4), 2
         Me.Hide
   End Select
End Sub

Private Sub Form_Activate()
'   ReadPatent
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Add By Sindy 2019/1/2
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010306_1 = Nothing
End Sub

Private Sub ReadPatent()
Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
Dim strCP27 As String 'Add By Sindy 2025/9/26
   
   For Each Lbl In Label12
      Lbl = ""
   Next
   
   'Add By Sindy 2019/1/2
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If Val(cp(84)) > 0 Then
         txtCP84 = cp(84) '發文規費
      ElseIf Val(cp(17)) > 0 Then
         txtCP84 = cp(17) '規費
      Else
         txtCP84 = 0
      End If
   End If
   '2019/1/2 END
   
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      AddCboName Combo1, pa(5), pa(6), pa(7)
   End If
   
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,CP110 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      Label12(0) = "" & .Fields(0)
      Label12(4) = "" & .Fields(1)
      Label12(5) = "" & .Fields(2)
      m_CP43 = "" & .Fields("CP43")
      m_CP110 = "" & .Fields("cp110")
      '抓相關總收文號內容
      If m_CP43 <> Empty Then
         'Modify By Sindy 2021/10/6 + ,CP64
         strExc(0) = "SELECT CP05,CP08,CP20,CP60,CP10,NVL(CP84,CP17) as CP84,CP43,CP64 FROM CASEPROGRESS WHERE CP09='" & m_CP43 & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With rsTemp1
            If m_CP43 > "C" Then
               Label12(6) = TransDate("" & .Fields("CP05"), 1)
               Label12(7) = "" & .Fields("CP08")
            End If
            '退費的相關收文若未請款則更新為不請款
            m_CP20 = "" & .Fields("CP20")
            m_CP60 = "" & .Fields("CP60")
            
            'Add by Morgan 2009/9/11 +判斷是否為實審、再審退費
            stRefCP10 = .Fields("CP10")
            
            'Added by Morgan 2013/6/6 延期退費改抓下一程序案件性質 Ex:FCP-32877
            If stRefCP10 = "404" Then
               If .Fields("CP43") > "C" Then
                  strExc(0) = "select np07 from nextprogress where np01='" & .Fields("CP43") & "'"
               Else
                  strExc(0) = "select cp10 from caseprogress where cp09='" & .Fields("CP43") & "'"
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  stRefCP10 = RsTemp(0)
               End If
            End If
            'end 2013/6/6
            
            If stRefCP10 = "416" Or stRefCP10 = "107" Then
               Text6 = "5"
               Text6.Locked = True
               '預設發文規費
               Text8(1) = Val("" & .Fields("CP84"))
            'Add By Sindy 2021/10/6
            ElseIf cp(10) = "908" Then '代辦退費
               '預設發文規費
               Text8(1) = Val("" & .Fields("CP84"))
            End If
            
            End With
         End If
      End If
   End If
   End With
   
   'Add By Sindy 2020/3/2
   '來函文號:相關總收文號有1901通知退費
   strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
               " AND CP43 in(SELECT c2.CP09" & _
                  " FROM caseprogress c1,caseprogress c2" & _
                  " WHERE c1.CP01='" & pa(1) & "' AND c1.CP02='" & pa(2) & "' AND c1.CP03='" & pa(3) & "' AND c1.CP04='" & pa(4) & "'" & _
                  " AND c1.CP09='" & cp(9) & "' AND c1.cp43 is not null AND c1.cp43=c2.CP09(+))" & _
               " AND cp10='1901' AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) DESC"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("ED08")) Then
         m_SendDate = RsTemp("ED08") - 19110000
         If Not IsNull(RsTemp("cp08")) Then
            strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
            m_SendWord = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
            strExc(0) = Replace(strExc(0), m_SendWord & "字第", "")
            m_SendNumber = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
         End If
      End If
   Else
   '2020/3/2 END
      'Add By Sindy 2019/7/23
      '來函文號:
      strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & cp(9) & "' AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp("ED08")) Then
            m_SendDate = RsTemp("ED08") - 19110000
            If Not IsNull(RsTemp("cp08")) Then
               strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
               m_SendWord = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
               strExc(0) = Replace(strExc(0), m_SendWord & "字第", "")
               m_SendNumber = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
            End If
         End If
      End If
   End If
   
   'Modify by Morgan 2009/9/11 若為再審退費時，若有延期時退費金額抓延期發文規費
   If Text6 = "5" Then
      If stRefCP10 = "107" Then
         strExc(0) = "SELECT CP27,CP84 FROM CASEPROGRESS WHERE CP43='" & m_CP43 & "' AND CP10='404' AND CP27>0" & _
            "union all SELECT CP27,CP84 FROM CASEPROGRESS A WHERE CP43=(SELECT B.CP43 FROM CASEPROGRESS B WHERE B.CP09='" & m_CP43 & "') AND CP10='404' AND CP27>0 ORDER BY CP27 ASC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Text8(1) = Val("" & RsTemp.Fields("CP84"))
            m_404CP27 = RsTemp.Fields("cp27")
         End If
      End If
   Else
      strExc(0) = "SELECT Max(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & 年費
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then If Not IsNull(RsTemp.Fields(0)) Then Text9 = TransDate(RsTemp.Fields(0), 1):
      strYear = Text9
   End If
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
   Call Text6_Validate(False) 'Add By Sindy 2020/9/21
   
   'Add By Sindy 2025/9/26
   '針對輸入各式申請書中之"代辦退費-實體審查"時，
   '判斷其進度檔有超頁費或超項費且「實體審查之發文日」與「超頁費或超項費之發文日」不同日，
   '若不同,則彈跳提醒"本案除實審規費外另有超頁費或超項費需一併退費，請注意退費金額及收據有二張"
   If Text6 = "5" Then
      strExc(0) = "select * from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 in('938','939')" & _
                  " AND exists (select c.cp09 from caseprogress c where c.CP01='" & pa(1) & "' AND c.CP02='" & pa(2) & "' AND c.CP03='" & pa(3) & "' AND c.CP04='" & pa(4) & "' and c.cp10='416')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strCP27 = "" & RsTemp.Fields("cp27")
         strExc(0) = "select * from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10='416'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val(strCP27) > 0 And (Val("" & RsTemp.Fields("cp27")) <> Val(strCP27)) Then
               MsgBox "本案除實審規費外另有 超頁費 或 超項費 需一併退費" & vbCrLf & "，請注意退費金額及收據有二張 !", vbInformation
            End If
         End If
      End If
   End If
   '2025/9/26 END
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 49 Or KeyAscii > 55) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   'Add by Morgan 2007/1/12 選4面詢退費時預設要修改申請書
   ElseIf KeyAscii = 52 Then
      Text7 = "Y"
   End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
Dim strTmpCP64 As String
   
   'Add By Sindy 2020/9/21 抓取收據號碼
   If Text6 = "1" Or Text6 = "2" Or Text6 = "5" Or Text6 = "7" Then
      '要去抓代辦退費掛的相關總收文號進度備註裡的收據號碼
      strExc(0) = "select cp09||cp64 from caseprogress where CP09 in(select c.cp43 from caseprogress c where c.CP09='" & cp(9) & "' AND c.CP43 is not null)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTmpCP64 = RsTemp.Fields(0)
      End If
      
   ElseIf Text6 = "6" Then
      '要去抓發明申請進度備註裡的收據號碼
      strExc(0) = "select cp09||cp64 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 in(" & NewCasePtyList & ") order by cp05 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTmpCP64 = RsTemp.Fields(0)
      End If
   End If
   If strTmpCP64 = "" And Text6 = "5" Then
      '實審、再審退費
      '要去抓代辦退費掛的相關總收文號進度備註裡的收據號碼
      '若實審的代辦退費抓不到收據號碼則抓(911)補收款(掛實體審查的相關總收文號)
      strExc(0) = "select cp09||cp64,cp43 from caseprogress" & _
                  " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                  " AND CP10='911' AND cp43 is not null" & _
                  " AND exists (select c.cp09 from caseprogress c where c.cp09=cp43 and c.cp10='416')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTmpCP64 = RsTemp.Fields(0)
      End If
   End If
   'Add By Sindy 2025/9/26
   If Trim(Text8(0)) = "" Then
   '2025/9/26 END
      If InStr(strTmpCP64, "收據號碼:") > 0 Then
         Text8(0) = Mid(strTmpCP64, InStr(strTmpCP64, "收據號碼:") + 5, 11)
      End If
   End If
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text8_GotFocus(Index As Integer)
  TextInverse Text8(Index)
End Sub

Private Sub Text8_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If Text6 <> "3" And Text8(Index) = "" Then
            MsgBox "請輸入年費收據編號 !", vbCritical
            Cancel = True
         End If
      Case 1
         If Text8(Index) = "" Then
            MsgBox "請輸入退費金額 !", vbCritical
            Cancel = True
         End If
   End Select
   If Cancel = True Then TextInverse Text8(Index)
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
 Static strDateOld As String
   If Text9 <> "" Then
      If Not ChkDate(Text9) Then
         Cancel = True
         TextInverse Text9
      Else
         If Text9 <> strYear Then
            If MsgBox("是否確定修改 ?", vbQuestion + vbYesNo) = vbNo Then
               Cancel = True
               TextInverse Text9
            End If
         End If
      End If
   End If
   'If Cancel = False Then strDateOld = Text9
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   If Me.Text5.Enabled = True Then
      Cancel = False
      Text5_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   Call Text6_Validate(False) 'Add By Sindy 2020/9/21
   
   For Each objTxt In Text8
      If objTxt.Enabled = True Then
         Cancel = False
         Text8_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   If Me.Text9.Enabled = True Then
      Cancel = False
      Text9_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   'Add by Morgan 2005/8/8
'   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
'   End If
   
   TxtValidate = True
End Function

'Add by Morgan 2005/8/8
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
