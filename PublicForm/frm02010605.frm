VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010605 
   BorderStyle     =   1  '單線固定
   Caption         =   "來函期限2次確認"
   ClientHeight    =   2964
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8472
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2964
   ScaleWidth      =   8472
   Begin VB.TextBox txtCaseField 
      Height          =   285
      Index           =   4
      Left            =   1170
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1770
      Width           =   915
   End
   Begin VB.TextBox txtCaseField 
      Height          =   285
      Index           =   3
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2550
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Height          =   405
      Left            =   1125
      TabIndex        =   32
      Top             =   2070
      Width           =   4305
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1950
         MaxLength       =   2
         TabIndex        =   4
         Top             =   105
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   840
         MaxLength       =   2
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   105
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   2970
         MaxLength       =   8
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   105
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      日"
         Height          =   225
         Index           =   2
         Left            =   2700
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   128
         Width           =   1515
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        月"
         Height          =   180
         Index           =   1
         Left            =   1710
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   150
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到           天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   150
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   350
      Index           =   0
      Left            =   6390
      TabIndex        =   8
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7215
      TabIndex        =   9
      Top             =   60
      Width           =   1170
   End
   Begin MSForms.Label lblSales 
      Height          =   255
      Left            =   6000
      TabIndex        =   35
      Top             =   1488
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4233;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseProperty 
      Height          =   255
      Left            =   6000
      TabIndex        =   37
      Top             =   1161
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4233;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgent 
      Height          =   252
      Left            =   1776
      TabIndex        =   36
      Top             =   840
      Width           =   2448
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4318;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   330
      Left            =   990
      TabIndex        =   10
      Top             =   420
      Width           =   7350
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12965;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   90
      TabIndex        =   34
      Top             =   2595
      Width           =   900
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      Caption         =   "來函期限："
      Height          =   180
      Left            =   90
      TabIndex        =   33
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label Label16 
      Caption         =   "官方發文日："
      Height          =   180
      Left            =   90
      TabIndex        =   31
      Top             =   1815
      Width           =   1080
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   3615
      TabIndex        =   30
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請號："
      Height          =   255
      Left            =   2865
      TabIndex        =   29
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   195
      Index           =   0
      Left            =   4275
      TabIndex        =   28
      Top             =   1161
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日："
      Height          =   195
      Left            =   90
      TabIndex        =   27
      Top             =   1488
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   195
      Left            =   90
      TabIndex        =   26
      Top             =   1161
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本所號："
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   25
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblIssue 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   195
      Left            =   2385
      TabIndex        =   24
      Top             =   1161
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   4275
      TabIndex        =   23
      Top             =   1488
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   5265
      TabIndex        =   22
      Top             =   1488
      Width           =   645
   End
   Begin VB.Label lblCaseField 
      Height          =   260
      Index           =   5
      Left            =   3200
      TabIndex        =   21
      Top             =   1160
      Width           =   950
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   900
      TabIndex        =   20
      Top             =   1161
      Width           =   1365
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   870
      TabIndex        =   19
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   1260
      TabIndex        =   18
      Top             =   1488
      Width           =   1005
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   2
      Left            =   900
      TabIndex        =   17
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   834
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   507
      Width           =   915
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   5265
      TabIndex        =   14
      Top             =   1161
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   195
      Left            =   4275
      TabIndex        =   13
      Top             =   834
      Width           =   915
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   9
      Left            =   5265
      TabIndex        =   12
      Top             =   834
      Width           =   645
   End
   Begin VB.Label lblNation 
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   834
      Width           =   2400
   End
End
Attribute VB_Name = "frm02010605"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/22 改成Form2.0 ; cboCaseName、lblSales、lblCaseProperty、lblAgent
'Created by Morgan 2020/7/20
Option Explicit

Public m_CP09 As String
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String

Dim strCP07 As String, strCP65 As String
Dim bolChkCP142 As Boolean 'Added by Morgan 2023/5/22
Dim bolCN445 As Boolean 'Added by Morgan 2024/9/25

Private Sub cmdok_Click(Index As Integer)
   Dim strLBL As String
   
   If Index = 0 Then
      strLBL = Replace(Label32, ":", "")
      If txtCaseField(3) = "" Then
         MsgBox "請輸入" & strLBL & "！", vbExclamation
         txtCaseField(3).SetFocus
         Exit Sub
      End If
      
      If DBDATE(txtCaseField(3)) <> strCP07 Then
         If MsgBox("本次輸入" & strLBL & "與原程序輸入的不同，是否確定要將來信退回？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
            Exit Sub
         Else
            If SaveDatabase(2) = False Then
               Exit Sub
            Else
               'Modify By Sindy 2022/5/20
               'frm04010519.GoNext
               Forms(0).Tmpfrm04010519.GoNext
               Set Forms(0).Tmpfrm04010519 = Nothing
               '2022/5/20 END
            End If
         End If
      Else
         If SaveDatabase(1) = False Then
            Exit Sub
         Else
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         End If
      End If
   End If
   Unload Me
End Sub

Private Sub Form_Activate()
   Static bolActivated As Boolean
   If bolActivated = False Then
      If ReadAllData = True Then
         bolActivated = True
      Else
         Unload Me
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache
   Set frm02010605 = Nothing
End Sub

Private Function ReadAllData() As Boolean
   'Removed by Morgan 2023/4/12 改呼叫時設定
   'm_strIR01 = frm04010519.m_strIR01
   'm_strIR02 = frm04010519.m_strIR02
   'm_strIR03 = frm04010519.m_strIR03
   'm_strIR04 = frm04010519.m_strIR04
   'end 2023/4/12
   
   'Modify By Sindy 2023/5/23 + trademark
   'Modify By Sindy 2024/4/2 + ServicePractice
   strExc(0) = "select a.*,b.cp10 cp10r,pa05,pa06,pa07,pa09,pa11,pa26 from caseprogress a,patent,caseprogress b where a.cp09='" & m_CP09 & "'" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and b.cp09(+)=a.cp43 and pa01 is not null" & _
               " union all select a.*,b.cp10 cp10r,tm05,tm06,tm07,tm10,tm12,tm23 from caseprogress a,trademark,caseprogress b where a.cp09='" & m_CP09 & "'" & _
      " and tm01(+)=a.cp01 and tm02(+)=a.cp02 and tm03(+)=a.cp03 and tm04(+)=a.cp04 and b.cp09(+)=a.cp43 and tm01 is not null" & _
               " union all select a.*,b.cp10 cp10r,sp05,sp06,sp07,sp09,sp11,sp08 from caseprogress a,ServicePractice,caseprogress b where a.cp09='" & m_CP09 & "'" & _
      " and sp01(+)=a.cp01 and sp02(+)=a.cp02 and sp03(+)=a.cp03 and sp04(+)=a.cp04 and b.cp09(+)=a.cp43 and sp01 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      strCP07 = "" & .Fields("cp07")
      
      'Added by Morgan 2023/4/19
      'Modified by Morgan 2023/5/22 FMP案會有C類指定送件日,增加判斷P的延期受理(檢查含在途法限cp142)
      If .Fields("cp01") = "P" Then
         If .Fields("cp10") = "1004" And Not IsNull(.Fields("cp142")) Then
            strCP07 = "" & .Fields("cp142")
            bolChkCP142 = True
         End If
      'Added by Morgan 2021/2/26
      ElseIf strCP07 = "" Then
         strCP07 = "" & .Fields("cp142")
         bolChkCP142 = True 'Added by Morgan 2023/5/22
      End If
      'end 2023/4/19
      
      'Added by Morgan 2024/9/26
      If .Fields("cp01") = "P" And .Fields("pa09") = "020" And .Fields("cp10") = "1001" And .Fields("cp10r") = "445" And Not IsNull(.Fields("cp142")) Then
         strCP07 = "" & .Fields("cp142")
         bolChkCP142 = True
         bolCN445 = True
         Label32 = "專利權期滿終止日:"
         txtCaseField(3).Left = Label32.Left + Label32.Width + 150
      End If
      'end 2024/9/26
         
      strCP65 = .Fields("cp65")
      lblCaseField(0) = MergeString(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
      lblCaseField(1) = "" & .Fields("pa11")
      SetNameToCombo cboCaseName, "" & .Fields("pa05"), "" & .Fields("pa06"), "" & .Fields("pa07")
      lblCaseField(2) = "" & .Fields("pa26")
      If ClsPDGetCustomer(lblCaseField(2), strExc(1)) Then
         lblAgent.Caption = strExc(1)
      End If
      lblCaseField(9) = "" & .Fields("pa09")
      If ClsPDGetNation(lblCaseField(9), strExc(1)) Then
         lblNation.Caption = strExc(1)
      End If
      lblCaseField(4) = m_CP09
      lblCaseField(5) = TransDate(.Fields("cp05"), 1)
      lblCaseField(6) = .Fields("cp10")
      If ClsPDGetCaseProperty(.Fields("cp01"), lblCaseField(6), strExc(1)) Then
         'Modified by Morgan 2024/9/26
         'lblCaseProperty = strExc(1)
         lblCaseProperty = strExc(1) & PUB_GetRelateCasePropertyName(.Fields("cp09"), "1")
         'end 2024/9/26
      End If
      lblCaseField(7) = .Fields("cp13")
      lblSales = GetStaffName(lblCaseField(7), True)
      'Modify By Sindy 2023/5/23
      If Val("" & .Fields("cp119")) = 0 Then
         Label3.Visible = False
         lblCaseField(8).Visible = False
      Else
      '2023/5/23 END
         lblCaseField(8) = TransDate("" & .Fields("cp119"), 1)
      End If
      
      'Added by Morgan 2023/4/18 參考P案核准
      '大陸核准領證期限=核准日+2個月
      If .Fields("pa09") = "020" And .Fields("cp10") = "1001" And Len("" & .Fields("cp10r")) = 3 Then
         If InStr("101,102,103,104,105,307", .Fields("cp10r")) > 0 Then
            Text11 = "2"
         End If
      End If
      'end 2023/4/18
      'Added by Lydia 2023/09/25 行政訴訟期限2次稽核作業: 改變顯示
      If .Fields("cp01") = "FCP" And Len(.Fields("cp10")) = 4 Then
         If (.Fields("cp10") = "1002" And .Fields("cp10r") = "503") Or InStr("1813,1502,1807", .Fields("cp10")) > 0 Then
            '行政訴訟核駁、1813行政裁定、1502撤銷原處分、1807對方補充說明=>再次輸入送達日期
            Label16.Caption = "送達日期："
         ElseIf InStr("1210,1211,1212", .Fields("cp10")) > 0 Then
            Label16.Caption = "開庭日期："
         End If
      End If
      'end 2023/09/25
      End With
      
      ReadAllData = True
      
      'Modified by Morgn 2024/9/25 + Or bolCN445
      If Left(Pub_StrUserSt03, 2) = "F1" Or bolCN445 Then txtCaseField(3).SetFocus  'Add By Sindy 2023/6/30
   Else
      MsgBox "無法讀取案件資料，請確認收文號是否正確！", vbExclamation
   End If
End Function

Private Sub Option4_Click(Index As Integer)
   If Option4(Index).Value = True Then
      If Index = 0 Then
         If Text10.Enabled Then Text10.SetFocus
      ElseIf Index = 1 Then
         If Text11.Enabled Then Text11.SetFocus
      ElseIf Index = 2 Then
         If Text12.Enabled Then Text12.SetFocus
      End If
   End If
End Sub

Private Sub Option4_GotFocus(Index As Integer)
   Option4_Click Index
End Sub

Private Sub Text10_Change()
   If Text10 <> "" Then
      Option4(0).Value = True
      Text11 = ""
      Text12 = ""
   End If
End Sub

Private Sub Text10_GotFocus()
    TextInverse Text10
    CloseIme
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_Change()
   If Text11 <> "" Then
      Option4(1).Value = True
      Text10 = ""
      Text12 = ""
   End If
End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
  CloseIme
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_Change()
   If Text12 <> "" Then
      Option4(2).Value = True
      Text10 = ""
      Text11 = ""
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   CloseIme
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Or Text12 = "" Then Exit Sub
   If ChkDate(Text12.Text) Then
        If Val(DBDATE(Text12)) < Val(strSrvDate(1)) Then
            MsgBox "來函期限不可小於系統日 !", vbCritical
            Cancel = True
        Else
            '轉民國年
            txtCaseField(3) = TransDate(Text12, 1)
        End If
    Else
        Cancel = True
    End If
End Sub

'Added by Morgan 2020/7/16 計算期限
Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '期限起算日
   Dim iDays1 As Integer, iDays2 As Integer, iDays3 As Integer, iDays4 As Integer 'Add by Morgan 2009/12/1
   
   If txtCaseField(4) <> "" Then
      '起算日=准駁通知日
      strFromDate = txtCaseField(4)
         
      '文到天數
      If Option4(0).Value = True And Text10 <> "" Then
         txtCaseField(3) = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         
      '文到月數
      ElseIf Option4(1).Value = True And Text11 <> "" Then
         txtCaseField(3) = TransDate(AddMonth(strFromDate, Val(Text11)), 1)
         
      End If
   End If
End Sub

Private Sub txtCaseField_GotFocus(Index As Integer)
   TextInverse txtCaseField(Index)
   CloseIme
End Sub

Private Sub txtCaseField_LostFocus(Index As Integer)
   'Added by Morgan 2023/4/17
   If Index = 4 Then
      'Modified by Morgan 2024/12/18
      'If Me.ActiveControl = txtCaseField(3) Then Exit Sub 'Added by Morgan 2024/9/26
      If TypeName(ActiveControl) <> "OptionButton" Then Exit Sub
      If Pub_StrUserSt03 = "F22" Then
         If lblCaseField(6) = "1004" Then
            Option4(2).Value = True
            If Text12.Enabled Then Text12.SetFocus
         'Else
         '   Option4(1).Value = True
         '   If Text11.Enabled Then Text11.SetFocus
         End If
      'ElseIf Option4(1).Value = True Then
      '   If Text11.Enabled Then Text11.SetFocus
      End If
      'end 2024/12/18
   End If
   'end 2023/4/17
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   If Index = 4 And txtCaseField(Index) <> "" Then
      If ChkDate(txtCaseField(Index).Text) Then
         If DBDATE(txtCaseField(Index)) > Val(strSrvDate(1)) Then
            'Modified by Lydia 2023/09/25
            'MsgBox "官方發文日不可大於系統日！", vbCritical
            If InStr(Label16, "發文") > 0 Then 'Added by Lydia 2025/04/29 開庭通知可能是預先
               MsgBox Mid(Label16, 1, Len(Label16) - 1) & "不可大於系統日！", vbCritical
               Cancel = True
            End If
         Else
            If txtCaseField(Index).Tag <> txtCaseField(Index) Then GetTime
            txtCaseField(Index).Tag = txtCaseField(Index)
         End If
      Else
         txtCaseField_GotFocus Index
         Cancel = True
      End If
      
   End If
End Sub

Private Function SaveDatabase(pResult As Integer) As Boolean
   Dim strUpdTime As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrorHandler

   strUpdTime = Right("000000" & ServerTime, 6)
   
   If pResult = 1 Then
      'Added by Morgan 2023/4/12
      '外專程序2次確認後上7.已確認(主管核准)
      'If Pub_StrUserSt03 = "F22" Then
      'Modify By Sindy 2023/5/23 + 外商 => Or Left(Pub_StrUserSt03, 2) = "F1"
      If Pub_StrUserSt03 = "F22" Or Left(Pub_StrUserSt03, 2) = "F1" Then
         strSql = "update InputRecord set ir16='7',ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                  " where ir01=" & m_strIR01 & " and ir02=" & m_strIR02 & " and ir03='" & m_strIR03 & "' and ir04='" & m_strIR04 & "' and ir08=0"
         cnnConnection.Execute strSql, intI
         
         'Add By Sindy 2025/4/29
         '未處理的都自沖(上確認日期時間人員)
         '同部門判斷 and exists(select st01 from staff where st01=ir04 and st03='" & PUB_GetST03(Trim(Left(Combo1, 6))) & "')
         strSql = "update InputRecord set " & _
                  "ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                  " where ir01=" & m_strIR01 & _
                    " and ir02=" & m_strIR02 & _
                    " and ir03='" & m_strIR03 & "'" & _
                    " and exists(select st01 from staff where st01=ir04 and st03='" & PUB_GetST03(m_strIR04) & "')" & _
                    " and ir08=0"
         cnnConnection.Execute strSql, intI
         '2025/4/29 END
         
         'Added by Morgan 2023/4/18
         strSql = "update IPDeptInput set ii16=" & strSrvDate(1) & _
                  " where Ii01=" & m_strIR01 & " and Ii02=" & m_strIR02 & _
                    " and Ii03='" & m_strIR03 & "' and ii16=0"
         cnnConnection.Execute strSql, intI
         'end 2023/4/18
         
         'Added by Morgan 2023/4/19 清除暫存的官方含在途法限
         If bolChkCP142 Then 'Added by Morgan 2023/5/22
            strSql = "update caseprogress set cp142=null where cp09='" & m_CP09 & "' and cp142 is not null"
            cnnConnection.Execute strSql, intI
         End If
         'end 2023/4/19
      Else
      'end 2023/4/12
      
         'Modified by Morgan 2020/8/13 改期限一致仍轉回原程序上已處理後進行送判/發函作業
         ''設定確認日期時間人員
         'strSql = "update InputRecord set ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
         '         " where ir01=" & m_strIR01 & " and ir02=" & m_strIR02 & " and ir03='" & m_strIR03 & "' and ir04='" & m_strIR04 & "' and ir08=0"
         'cnnConnection.Execute strSql, intI
         ''設定可刪除日期
         'strSql = "update PatentInput set pi16=" & strSrvDate(1) & _
         '         " where pi01=" & m_strIR01 & " and pi02=" & m_strIR02 & " and pi03='" & m_strIR03 & "' and pi16=0"
         'cnnConnection.Execute strSql, intI
         strSql = "update InputRecord set ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                  ",ir16='7',ir20='" & strSrvDate(2) & " 2次確認OK;'||decode(ir20,null,null,chr(13)||chr(10)||ir20),ir22=''" & _
                  " where ir01=" & m_strIR01 & " and ir02=" & m_strIR02 & " and ir03='" & m_strIR03 & "' and ir04='" & m_strIR04 & "' and ir08=0"
         cnnConnection.Execute strSql, intI
         'end 2020/8/13
         
      End If 'Added by Morgan 2023/4/12
   Else
      '設定確認日期時間人員
      'Modified by Morgan 2023/4/17 取消 ir22=''
      strSql = "update InputRecord set ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
               ",ir16='8',ir20='" & strSrvDate(2) & " 2次確認退回,法定期限：" & txtCaseField(3) & ";'||decode(ir20,null,null,chr(13)||chr(10)||ir20)" & _
               " where ir01=" & m_strIR01 & " and ir02=" & m_strIR02 & " and ir03='" & m_strIR03 & "' and ir04='" & m_strIR04 & "' and ir08=0"
      cnnConnection.Execute strSql, intI
      
      'EMail通知原程序
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         " values('" & strUserNum & "','" & strCP65 & "',to_char(sysdate,'yyyymmdd')" & _
         ",to_char(sysdate,'hh24miss'),'請再次確認" & Replace(lblCaseField(0), " ", "") & "(" & lblCaseProperty & ")法定期限','如旨')"
      cnnConnection.Execute strSql, intI
      
      'Added by Morgan 2024/9/27
      '大陸專利權期限補償核准退回時要還原基本檔的專用期止日
      If bolCN445 Then
         strSql = "update patent set pa25=(select to_char(to_date(pa25,'yyyymmdd')-(cp71+0),'yyyymmdd') from caseprogress where cp09='" & m_CP09 & "')" & _
            " where (pa01,pa02,pa03,pa04)=(select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & m_CP09 & "' and CP71>0 and to_char(to_date(cp142,'yyyymmdd')-1,'yyyymmdd')=pa25)"
         cnnConnection.Execute strSql, intI
      End If
      'end 2024/9/27
   End If
   
   cnnConnection.CommitTrans
   SaveDatabase = True
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function
