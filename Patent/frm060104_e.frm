VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_e 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-機關來函"
   ClientHeight    =   3948
   ClientLeft      =   276
   ClientTop       =   960
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3948
   ScaleWidth      =   8760
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   14
      Top             =   540
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   13
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   12
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   11
      Top             =   540
      Width           =   375
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3285
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2670
      Width           =   600
   End
   Begin VB.TextBox txtCP14 
      Height          =   270
      Left            =   4995
      MaxLength       =   9
      TabIndex        =   2
      Top             =   2670
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7404
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6576
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtCP27 
      Height          =   270
      Left            =   1035
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2670
      Width           =   1095
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   10
      Top             =   840
      Width           =   7365
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "12991;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   825
      Left            =   1035
      TabIndex        =   3
      Top             =   3030
      Width           =   6660
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11747;1455"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   6
      Left            =   6105
      TabIndex        =   33
      Top             =   2670
      Width           =   1680
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2963;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   32
      Top             =   1830
      Width           =   3780
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6667;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   4
      Left            =   1140
      TabIndex        =   31
      Top             =   2160
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   30
      Top             =   1170
      Width           =   2310
      VariousPropertyBits=   27
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   29
      Top             =   2160
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   9
      Left            =   3645
      TabIndex        =   28
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   8
      Left            =   180
      TabIndex        =   27
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Index           =   7
      Left            =   3645
      TabIndex        =   26
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Index           =   3
      Left            =   3660
      TabIndex        =   25
      Top             =   1170
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   24
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   21
      Top             =   1830
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   1140
      TabIndex        =   20
      Top             =   1830
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   19
      Top             =   1170
      Width           =   2310
      VariousPropertyBits=   27
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   1500
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日:"
      Height          =   180
      Index           =   5
      Left            =   3660
      TabIndex        =   17
      Top             =   1500
      Width           =   585
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   16
      Top             =   1500
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   1500
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   12
      Left            =   2430
      TabIndex        =   9
      Top             =   2715
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Index           =   11
      Left            =   4275
      TabIndex        =   8
      Top             =   2715
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   8520
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   8520
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   13
      Left            =   180
      TabIndex        =   7
      Top             =   3030
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Index           =   10
      Left            =   180
      TabIndex        =   6
      Top             =   2715
      Width           =   585
   End
End
Attribute VB_Name = "frm060104_e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/17 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Add by Morgan 2007/7/19
Option Explicit

Dim pa() As String
Dim intWhere As Integer
Dim m_CP10 As String, m_CP09 As String
Dim m_CP60 As String 'Added by Lydia 2015/02/26
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      m_CP09 = .Tag
      Label3(0) = m_CP09
   End With
   ReDim pa(TF_PA)
   ReadPatent
   Combo1.ListIndex = 0
   txtCP27 = strSrvDate(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Call PUB_SendMailCache 'Added by Lydia 2025/04/25
   Set frm060104_e = Nothing
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If TxtValidate = False Then Exit Sub
         'Add by Morgan 2008/3/19
         '若未辦或不辦重新委任時提醒
         If m_CP10 = "1802" Then
            If PUB_Check928NotOk(pa) = True Then
               If MsgBox("本案下一程序有重新委任之補文件未辦理，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
         End If
         
         'Add by Sindy 2021/11/17 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         If FormSave = False Then
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
            Exit Sub
         Else
         
            'Add by Morgan 2008/2/20 檢查代理人Email
            If pa(1) = "FCP" Then
               PUB_CheckEMail pa(75), pa(144)
               If pa(145) <> "" Then
                  PUB_CheckEMail pa(75), pa(145)
               End If
            Else
               PUB_CheckEMail pa(26), pa(76)
               If pa(77) <> "" Then
                  PUB_CheckEMail pa(26), pa(77)
               End If
            End If
            'end 2008/2/20
            
            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2023/11/9 End
            '若有未發文資料顯示警告
            'Modify By Sindy 2023/11/9
            If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
               frm060104_1.Show
               frm060104_1.ReQuery
            Else
               'Add By Sindy 2023/11/9
               If frm060104_1.bolIsEMPFlow = True Then
                  Unload frm060104_1
               Else
               '2023/11/9 End
                  frm060104_1.Show
                  frm060104_1.Clear
               End If
            End If
         End If
      Case 1
         frm060104_1.Show
   End Select
   Unload Me
End Sub

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   txtCP27_Validate bCancel
   If bCancel = True Then Exit Function
   txtCP14_Validate bCancel
   If bCancel = True Then Exit Function
   txtCP113_Validate bCancel
   If bCancel = True Then Exit Function
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   Dim stCust As String, stFAgent As String
   
On Error GoTo CheckingErr
   cnnConnection.BeginTrans
   strSql = "Update Caseprogress set cp27=" & DBDATE(txtCP27) & ",cp14='" & txtCP14 & "',cp113=" & IIf(txtCP113 = "", "NULL", txtCP113) & ",cp64='" & ChgSQL(txtCP64) & "' where cp09='" & m_CP09 & "'"
   cnnConnection.Execute strSql, intI
   
 'Added by Lydia 2015/02/26 若已開請款單則換承辦人或核稿人時發Mail通知靜芳
   If m_CP60 > "X" Then
      'Modified by Lydia 2019/10/17 本所案號+"-"
      'PUB_PointReAssignInform Text1 & Text2 & Text3 & Text4, m_CP60, txtCP14.Tag, txtCP14.Text
      PUB_PointReAssignInform pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)), m_CP60, txtCP14.Tag, txtCP14.Text
   End If
   
   'Added by Lydia 2016/07/11 代理人Y20656+申請人X7072201 & X70286 之案件，機關來函除核准函之外之有期限來函在OA發文後自動建立行事曆。
   'Modified by Morgan 2018/8/9 +FG案判斷(欄位不同)
   'strExc(0) = ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30))
   If pa(1) = "FG" Then
      stCust = ChangeCustomerL(pa(8)) & "," & ChangeCustomerL(pa(58)) & "," & ChangeCustomerL(pa(59)) & "," & ChangeCustomerL(pa(65)) & "," & ChangeCustomerL(pa(66))
      stFAgent = ChangeCustomerL(pa(26))
   Else
      stCust = ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30))
      stFAgent = ChangeCustomerL(pa(75))
   End If
   'end 2018/8/9
   
   'Modified by Lydia 2017/03/08 +Y53942案件在OA發文後自動建立行事曆
   'If m_CP10 <> 核准 And Label3(5).Caption <> "" And InStr(ChangeCustomerL(pa(75)), "Y20656") > 0 And pa(75) <> "" And strExc(0) <> "" And (InStr(strExc(0), "X7072201") > 0 Or InStr(strExc(0), "X70286") > 0) Then
   'Modified by Lydia 2017/09/22 Y53942 已無此要求
   'If m_CP10 <> 核准 And Label3(5).Caption <> "" And ((InStr(ChangeCustomerL(pa(75)), "Y20656") > 0 And pa(75) <> "" And strExc(0) <> "" And (InStr(strExc(0), "X7072201") > 0 Or InStr(strExc(0), "X70286") > 0)) _
                                                      Or (InStr(ChangeCustomerL(pa(75)), "Y53942") > 0 And pa(75) <> "")) Then
   'Modified by Morgan 2018/8/9
   'If m_CP10 <> 核准 And Label3(5).Caption <> "" And ((InStr(ChangeCustomerL(pa(75)), "Y20656") > 0 And pa(75) <> "" And strExc(0) <> "" And (InStr(strExc(0), "X7072201") > 0 Or InStr(strExc(0), "X70286") > 0))) Then
   If m_CP10 <> 核准 And Label3(5).Caption <> "" And ((InStr(stFAgent, "Y20656") > 0 And pa(75) <> "" And stCust <> "" And (InStr(stCust, "X7072201") > 0 Or InStr(stCust, "X70286") > 0))) Then
   'end 2018/8/9
      strExc(1) = PUB_GetWorkDay1(CompDate(2, 14, DBDATE(txtCP27)), True) 'OA發文後兩周(以日曆天計算，結果若為假日則提前至前一工作天)
      
      '提醒人員
      strExc(2) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序管制人
      strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) '承辦智權人員
      'Modified by Lydia 2017/03/08 Y53942案件
      'Remove by Lydia 2017/09/22 Y53942 已無此要求
      'If InStr(ChangeCustomerL(pa(75)), "Y53942") > 0 Then
      '   'Modified by Lydia 2017/03/14 Y53942 更名 Tessera->Xperi
      '   strExc(4) = "如未獲指示(請向工程師確認)則交承辦向Xperi發Reminder。" '備註
      'Else
         strExc(4) = "如未獲指示(請向工程師確認)則交承辦向Tessera發Reminder，副本Lerner。" '備註
      'End If
      'end 2017/03/08
      'end 2017/09/22
      
      '可解除人員預設為程序管制人
      If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(2) & IIf(strExc(3) <> "", "," & strExc(3), ""), strExc(4), strExc(2), "1", pa(1), pa(2), pa(3), pa(4)) Then
      End If
   End If
   'end 2016/07/11
   
   'Add By Sindy 2017/1/11 核准函之C類來函發文時（即工程師已將分割交程序後），計算核准函發文日＋8個日曆天(當日不算)去更新該案"通知告准" D類進度的承辦期限
   If m_CP10 = 核准 Then
      strSql = "Update Caseprogress set cp48=" & CompDate(2, 8, DBDATE(txtCP27)) & _
               " where cp43='" & m_CP09 & "' and cp10='1917' and cp48 is null and cp27||cp57 is null"
      cnnConnection.Execute strSql, intI
   End If
   '2017/1/11 END
   
   'Added by Morgan 2017/8/17
   '非屬相同創作發文新增行事曆
   If m_CP10 = "1919" Then
         strExc(1) = CompDate(1, 1, strSrvDate(1))
         strExc(2) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序管制人
         strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) '承辦智權人員
         strExc(4) = "追蹤客戶是否解除一案二請"
         PUB_AddFCPStaffCalendar strExc(1), "1", strExc(2) & IIf(strExc(3) <> "", "," & strExc(3), ""), strExc(4), strExc(2), "1", pa(1), pa(2), pa(3), pa(4)
   End If
   'end 2017/8/17
   
   'Added by Lydia 2018/05/02 :Y20600000 BREVALEX 請管制以下C類來函性質發文後 , 自動產生行事曆
   'Modified by Morgan 2018/8/9
   'If InStr(ChangeCustomerL(pa(75)), "Y20600") > 0 And InStr("1002,1201,1202,1203,1227,1232,1401", m_CP10) > 0 Then
   If InStr(stFAgent, "Y20600") > 0 And InStr("1002,1201,1202,1203,1227,1232,1401", m_CP10) > 0 Then
   'end 2018/8/9
         strExc(1) = CompWorkDay(3, strSrvDate(1))
         strExc(2) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序管制人
         strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) '承辦智權人員
         strExc(5) = strExc(2) & IIf(strExc(3) <> "", "," & strExc(3), "") '提醒和解除
         strExc(4) = "追蹤客戶有無確收"
         PUB_AddFCPStaffCalendar strExc(1), "1", strExc(5), strExc(4), strExc(5), "1", pa(1), pa(2), pa(3), pa(4)
   End If
   'end 2018/05/02
   
   'Added by Lydia 2018/10/30 Y55129 ADELI LLP + X79754 Xcelsis Corporation 之案件機關來函發文後自動建立行事曆
                            '來函案件性質: 審查意見通知函1202、最後通知1227、核駁函1002
   If InStr(stFAgent, "Y55129") > 0 And InStr(stCust, "X79754") > 0 And InStr("1002,1202,1227", m_CP10) > 0 Then
         '期限請抓OA發文後兩周,非工作日往前推
         strExc(1) = CompWorkDay(1, CompDate(2, 14, strSrvDate(1)), 1)
         strExc(2) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序管制人
         strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) '承辦智權人員
         strExc(5) = strExc(2) & IIf(strExc(3) <> "", "," & strExc(3), "") '提醒
         strExc(4) = "如未獲指示則交承辦發Reminder"
         '提醒人員請掛管制人及承辦組人員，解除人員請掛管制人及其案件職代(模組有預設抓案件職代)
         PUB_AddFCPStaffCalendar strExc(1), "1", strExc(5), strExc(4), strExc(2), "1", pa(1), pa(2), pa(3), pa(4)
   End If
   'end 2018/10/30
   
   'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：收到學名藥廠P4通知1922: (1)程序輸入C類來函「P4通知1922」(承辦期限3個工作天, 法定期限45天)
                              '(2)「P4通知」發文時,自動收文告代(承辦期限10個工作天)
   If pa(1) = "FCP" And m_PA177 = "Y" And m_CP10 = "1922" Then
      strExc(1) = AutoNo("B", 6)
      strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
      strExc(2) = GetSalesArea(strExc(3))
      strExc(4) = CompWorkDay(11, DBDATE(txtCP27))
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP48,CP43) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
                     strSrvDate(1) & ",'" & strExc(1) & "','901','" & strExc(2) & "','" & strExc(3) & "','" & txtCP14 & "'," & strExc(4) & ",'" & m_CP09 & "')"
      cnnConnection.Execute strSql
   End If
   'end 2023/07/28
   
   'Added by Lydia 2025/04/25 特定客戶優先二核期限控管; 日代<Y4520400> SOEI、<Y5518900>TOKOSHIE 各項指示、二核相關備註
   If m_CP10 = 專利證書 And InStr("Y45204000,Y55189000", Mid(stFAgent, 1, 8)) > 0 Then  'Memo by Lydia 2025/08/22 若有異動請一併修改整批發文frm060118
      strExc(0) = "select cp14,st04,cp09 from caseprogress,staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp158=0 and cp10='926' and cp14=st01(+) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '掛上【核對已准專利】期限：承辦期限=本所期限=指定日期之前'證書發文日起算14個日曆天(遇假日往前至前一工作天)
         strExc(1) = PUB_GetWorkDay1(CompDate(2, 14, DBDATE(txtCP27)), True)
         strSql = "Update CaseProgress Set cp48=" & strExc(1) & ", cp06=" & strExc(1) & " where cp09='" & RsTemp.Fields("cp09") & "' "
         cnnConnection.Execute strSql
         '自動發一封Email給承辦工程師及程序人員(特殊設定)  , 內容如下
         strExc(2) = "" & RsTemp.Fields("cp14")
         If "" & RsTemp.Fields("st04") <> "1" Then
            Call PUB_GetFCPCP14_F21(pa, strExc(2))
         End If
         If strExc(2) <> "" Then
            '收件人員: 承辦工程師、江如玉(固定核對公報程序人員)  副本:工程師主管
            strExc(3) = PUB_GetFCPEngSup(strExc(2))
            strExc(4) = Pub_GetSpecMan("外專程序-匯入公告本收件者")
            strExc(5) = "【優先二次核對】" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "證書已寄出、請優先二次核對"
            strExc(6) = "TO:" & GetStaffName(strExc(2), True) & vbCrLf & _
                        String(4, " ") & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "客戶要求於寄證書後2週內二核請款報告，期限: " & ChangeWStringToTDateString(strExc(1)) & "，請優先處理二次核對報告。"
            strExc(6) = strExc(6) & vbCrLf & vbCrLf & "TO:" & PUB_ReadUserData(strExc(4)) & vbCrLf & _
                         String(4, " ") & "請優先核對公報速退工程師進行二核。"
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " values( '" & strUserNum & "','" & strExc(2) & IIf(strExc(4) <> "", ";" & strExc(4), "") & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(5)) & "' ,'" & ChgSQL(strExc(6)) & "' ,'" & strExc(3) & "' )"
            cnnConnection.Execute strSql
         End If
      End If
   End If
   'end 2025/04/25
   
   cnnConnection.CommitTrans
   
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub ReadPatent()
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   m_PA177 = "" 'Added by Lydia 2023/07/28
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
            m_PA177 = pa(177)  'Added by Lydia 2023/07/28 FCP專利連結通知
         End If
      Case "FG"
         If PUB_ReadServicePracticeDatabase(pa(), intWhere) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
   End Select
   'Added by Lydia 2015/02/26 +CP60
   strExc(0) = "select cp05,cp10,cpm03,cp08,cp06,cp07,cp14,cp113,cp64,st02,CP60" & _
      " from caseprogress,casepropertymap,staff where cp09='" & Label3(0) & "'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   With RsTemp
      Label3(1) = .Fields("cp05") - 19110000
      Label3(2) = .Fields("cp10") & " " & .Fields("cpm03")
      m_CP10 = .Fields("cp10")
      'Modify By Sindy 2012/10/17
      'Label3(3) = .Fields("cp08")
      Label3(3) = "" & .Fields("cp08")
      '2012/10/17 End
      If Not IsNull(.Fields("cp06")) Then
         Label3(4) = .Fields("cp06") - 19110000
      'Added by Lydia 2016/08/01
      Else
         Label3(4) = ""
      End If
      If Not IsNull(.Fields("cp07")) Then
         Label3(5) = .Fields("cp07") - 19110000
      'Added by Lydia 2016/08/01
      Else
         Label3(5) = ""
      End If
      txtCP14 = "" & .Fields("cp14")
      'modify by sonia 2015/9/21
      'Label3(6) = "" & .Fields("st02")
      txtCP14_Validate False
      'end 2015/9/21
      'Added by Lydia 2015/02/26
      txtCP14.Tag = txtCP14.Text
      If Not IsNull(.Fields("CP60")) Then
         m_CP60 = .Fields("CP60")
      Else
         m_CP60 = ""
      End If
      'end 2015/02/26
      
      txtCP113 = "" & .Fields("cp113")
      txtCP64 = "" & .Fields("cp64")
   End With
   End If
End Sub

Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

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
   Cancel = Not PUB_CheckCP113(txtCP113, pa(1), m_CP10, txtCP14)
End Sub

Private Sub txtCP14_Change()
   Label3(6) = ""
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP14_Validate(Cancel As Boolean)
   If txtCP14 = "" Then
      MsgBox "承辦人不可空白 !", vbCritical
      Cancel = True
   Else
      'ADD BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
      txtCP14 = GetFCPUser(txtCP14)
      'END 2015/9/21
      Label3(6) = GetStaffName(txtCP14, True)
   End If
End Sub

Private Sub txtCP27_GotFocus()
   TextInverse txtCP27
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If txtCP27 = "" Then
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   ElseIf Not ChkDate(txtCP27) Then
      txtCP27_GotFocus
      Cancel = True
   End If
End Sub
