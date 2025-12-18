VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030209_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "電話通知-輸入期限"
   ClientHeight    =   4332
   ClientLeft      =   72
   ClientTop       =   996
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4332
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdInput 
      Caption         =   "多案案號輸入"
      Height          =   400
      Left            =   4860
      TabIndex        =   25
      Top             =   70
      Width           =   1425
   End
   Begin VB.TextBox textNP07 
      Height          =   300
      Left            =   6480
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3180
      Width           =   732
   End
   Begin VB.TextBox textNP07_2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   7290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3210
      Width           =   1692
   End
   Begin VB.TextBox textNP08 
      Height          =   300
      Left            =   3660
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3210
      Width           =   1000
   End
   Begin VB.TextBox textNP09 
      Height          =   300
      Left            =   1230
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3210
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7080
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7920
      TabIndex        =   10
      Top             =   70
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   5
      Top             =   360
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1755
      MaxLength       =   6
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2595
      MaxLength       =   1
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2835
      MaxLength       =   2
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   7
      Left            =   1230
      TabIndex        =   36
      Top             =   720
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "收文號:"
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   35
      Top             =   720
      Width           =   585
   End
   Begin MSForms.TextBox textNP15 
      Height          =   645
      Left            =   1230
      TabIndex        =   4
      Top             =   3540
      Width           =   7845
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13838;1138"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   645
      Left            =   1230
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2400
      Width           =   7845
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13838;1138"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   6
      Left            =   4860
      TabIndex        =   34
      Top             =   2100
      Width           =   4185
      VariousPropertyBits=   27
      Size            =   "7382;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   5
      Left            =   1230
      TabIndex        =   33
      Top             =   2100
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   4
      Left            =   4860
      TabIndex        =   32
      Top             =   1770
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   3
      Left            =   1230
      TabIndex        =   31
      Top             =   1770
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   1230
      TabIndex        =   30
      Top             =   1380
      Width           =   7575
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13361;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   2
      Left            =   4860
      TabIndex        =   29
      Top             =   1050
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   1
      Left            =   1230
      TabIndex        =   28
      Top             =   1050
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   0
      Left            =   4860
      TabIndex        =   27
      Top             =   720
      Width           =   2115
      VariousPropertyBits=   27
      Size            =   "3731;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      Caption         =   "申請人1:"
      Height          =   225
      Left            =   3990
      TabIndex        =   26
      Top             =   2130
      Width           =   675
   End
   Begin VB.Label Label10 
      Caption         =   "下一程序備註 :"
      Height          =   375
      Left            =   180
      TabIndex        =   24
      Top             =   3570
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "進度備註:"
      Height          =   225
      Left            =   180
      TabIndex        =   23
      Top             =   2460
      Width           =   765
   End
   Begin VB.Label Label6 
      Caption         =   "下一程序 :"
      Height          =   255
      Left            =   5610
      TabIndex        =   22
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "本所期限 :"
      Height          =   255
      Left            =   2790
      TabIndex        =   21
      Top             =   3270
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "法定期限 :"
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Top             =   3270
      Width           =   855
   End
   Begin VB.Label Label29 
      Caption         =   "案件性質:"
      Height          =   225
      Left            =   3990
      TabIndex        =   18
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label13 
      Caption         =   "來函收文日:"
      Height          =   225
      Left            =   180
      TabIndex        =   17
      Top             =   2130
      Width           =   945
   End
   Begin VB.Label Label11 
      Caption         =   "智權人員:"
      Height          =   225
      Left            =   3990
      TabIndex        =   16
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label9 
      Caption         =   "承辦人　:"
      Height          =   225
      Left            =   180
      TabIndex        =   15
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號:"
      Height          =   225
      Left            =   150
      TabIndex        =   14
      Top             =   390
      Width           =   765
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號:"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label5 
      Caption         =   "審定號數:"
      Height          =   225
      Left            =   3990
      TabIndex        =   12
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label7 
      Caption         =   "商標名稱:"
      Height          =   225
      Left            =   180
      TabIndex        =   11
      Top             =   1455
      Width           =   765
   End
End
Attribute VB_Name = "frm030209_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2022/07/28
Option Explicit
Dim m_CP09 As String, m_CP13 As String
Dim m_strCP27 As String  '彈訊息詢問「通知代理人」後輸入發文日
Dim m_AttachPath As String
Dim strCPP02List As String, strCPP04List As String, strDateList As String, strTimeList As String

Private Sub cmdOK_Click(Index As Integer)
   
   Select Case Index
      Case 0 '確定
         
         If TxtValidate = False Then Exit Sub
         
         m_strCP27 = ""
         '彈訊息詢問「通知代理人」
         strExc(0) = "請選擇：" & vbCrLf & _
                          "是，已通知代理人，請接著輸入發文日" & vbCrLf & _
                           "否，無需通知代理人更新發文日=111111" & vbCrLf & _
                           "取消，中斷作業。"
         intI = MsgBox(strExc(0), vbInformation + vbYesNoCancel + vbDefaultButton3)
         If intI = vbCancel Then
             Exit Sub
         ElseIf intI = vbYes Then
JumpReInput:
            strExc(1) = UCase(InputBox("已通知代理人，請接著輸入發文日或是取消：" & vbCrLf & "P.S.發文日不可大於系統日", "已通知代理人", strSrvDate(2)))
            If strExc(1) = "" Then
                Exit Sub
            Else
                If Len(strExc(1)) <> 7 Then
                    GoTo JumpReInput
                Else
                   If strExc(1) > strSrvDate(2) Then
                       GoTo JumpReInput
                   Else
                       If ChkDate(strExc(1)) = False Then
                           GoTo JumpReInput
                       End If
                   End If
                End If
                m_strCP27 = DBDATE(strExc(1))
            End If
         ElseIf intI = vbNo Then
              m_strCP27 = "19221111"
         End If
         '先下載檔案
         If strCPP02List <> "" Then
             PUB_KillAnyFile m_AttachPath
         End If
         If cmdInput.Tag <> "" Then
            strCPP02List = "": strCPP04List = "": strDateList = "": strTimeList = ""
            strExc(0) = "SELECT CPP01,CPP02,CPP04,CPP08,CPP09,CPP14 FROM CASEPAPERPDF " & _
                             "WHERE  CPP01='" & m_CP09 & "' AND NVL(CPP10,'N') <> 'D' AND UPPER(CPP02) LIKE '%.INCOM.MSG' " & _
                             "ORDER BY CPP06 DESC, CPP07 DESC "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                RsTemp.MoveFirst
                Do While Not RsTemp.EOF
                   If PUB_GetAttachFile_CPP(RsTemp.Fields("CPP01"), RsTemp.Fields("CPP02"), m_AttachPath) = False Then
                       MsgBox "檔案下載失敗[ " & RsTemp.Fields("CPP02") & " ]！", vbCritical
                       Exit Sub
                   Else
                       strCPP02List = strCPP02List & IIf(strCPP02List <> "", vbCrLf, "") & RsTemp.Fields("cpp02")
                       strCPP04List = strCPP04List & IIf(strCPP04List <> "", vbCrLf, "") & RsTemp.Fields("cpp04")
                       strDateList = strDateList & IIf(strDateList <> "", vbCrLf, "") & RsTemp.Fields("cpp08")
                       strTimeList = strTimeList & IIf(strTimeList <> "", vbCrLf, "") & RsTemp.Fields("cpp09")
                   End If
                   RsTemp.MoveNext
                Loop
            End If
         End If
         '--------------------------
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
                  
         '回到原畫面要清除畫面
         frm030209_01.ClearForm
         Unload Me
      Case 1 '回前畫面
         Unload Me
   End Select
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me

   With frm030209_01
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      m_CP09 = .Tag
   End With
   
   ReadTradeMark
   
   Call Pub_ChkExcelPath
   m_AttachPath = strExcelPath & Me.Name
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
       PUB_KillAnyFile m_AttachPath
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm030209_01.Show
   Set frm030209_02 = Nothing
End Sub

Private Sub ReadTradeMark()
Dim Lbl As Object

   For Each Lbl In lblFM2
      Lbl = ""
   Next
   Combo1.Clear
   textCP64.Text = "": textCP64.Tag = ""
   textNP07.Text = ""
   textNP07_2.Text = "":     textNP07_2.Tag = ""
   textNP08.Text = ""
   textNP09.Text = ""
   cmdInput.Tag = ""
   
   '基本檔
   strExc(0) = "select tm05,tm06,tm07,tm23, nvl(cu05,nvl(cu04,cu06)) cname1,tm12,tm15 " & _
                     "From trademark, customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
                     "and tm01='" & Text1 & "' and tm02='" & Text2 & "' and tm03='" & Text3 & "' and tm04='" & Text4 & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
        lblFM2(1).Caption = "" & RsTemp.Fields("tm12")
        lblFM2(2).Caption = "" & RsTemp.Fields("tm15")
        lblFM2(6).Caption = "" & RsTemp.Fields("cname1")
        If "" & RsTemp.Fields("tm07") <> "" Then
            Combo1.AddItem "外：" & RsTemp.Fields("tm07"), 0
        End If
        If "" & RsTemp.Fields("tm06") <> "" Then
            Combo1.AddItem "英：" & RsTemp.Fields("tm06"), 0
        End If
        Combo1.AddItem "中：" & RsTemp.Fields("tm05"), 0
        Combo1.ListIndex = 0
   End If
   '收文
   strExc(0) = "select sqldatet(cp05) CP05T ,cp09,cpm03,cp14,s1.st02 as cp14t,cp13,s2.st02 as cp13t,cp64" & _
          " from caseprogress, casepropertymap,staff s1,staff s2" & _
          " where cp09='" & m_CP09 & "' and cp158=0 and cp159=0 and cp01=cpm01(+) and cp10=cpm02(+)" & _
          " and cp14=s1.st01(+) and cp13=s2.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       lblFM2(7).Caption = m_CP09
       lblFM2(0).Caption = "" & RsTemp.Fields("cpm03")
       lblFM2(3).Caption = "" & RsTemp.Fields("cp14t")
       lblFM2(4).Caption = "" & RsTemp.Fields("cp13t")
       lblFM2(5).Caption = "" & RsTemp.Fields("cp05t")
       textCP64.Text = "" & RsTemp.Fields("cp64")
       m_CP13 = "" & RsTemp.Fields("CP13")
   End If
   
   'Mark by Lydia 2022/08/12: 111/08/10 阿蓮: 不受INCOM.MSG歸卷限制，如果以後有需要再人工上傳。
   'strExc(0) = "SELECT CPP01,CPP02,CPP14 FROM CASEPAPERPDF " & _
                    "WHERE  CPP01='" & m_CP09 & "' AND NVL(CPP10,'N') <> 'D' AND UPPER(CPP02) LIKE '%.INCOM.MSG' " & _
                    "ORDER BY CPP06 DESC, CPP07 DESC "
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'If intI = 0 Then
   '   MsgBox "卷宗區尚無INCOM.MSG，不可輸入多案案號！", vbExclamation + vbOKOnly, "卷宗區歸卷檢查"
   '   cmdInput.Enabled = False
   'End If
   'end 2022/08/12
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean

   If textNP07 = "" Then
       MsgBox "請輸入下一程序!!", vbExclamation
       textNP07.SetFocus
       Exit Function
   End If
   
   textNP07_Validate Cancel
   If Cancel = True Then
      textNP07.SetFocus
      Exit Function
   End If
      
   If textNP07 <> "" And Trim(textNP08) = "" Then
       MsgBox "請輸入本所期限!!", vbExclamation
       textNP08.SetFocus
       Exit Function
   End If
   If textNP07 <> "" And Trim(textNP09) = "" Then
       MsgBox "請輸入法定期限!!", vbExclamation
       textNP09.SetFocus
       Exit Function
   End If
   If Trim(textNP08) <> "" And textNP08 < strSrvDate(2) Then
       MsgBox "本所期限不可小於系統日!!", vbExclamation
       textNP08.SetFocus
       Exit Function
   End If
   If Trim(textNP09) <> "" And textNP09 < strSrvDate(2) Then
       MsgBox "法定期限不可小於系統日!!", vbExclamation
       textNP09.SetFocus
       Exit Function
   End If
   If Trim(textNP08) <> "" And Trim(textNP09) <> "" And textNP09 < textNP08 Then
       MsgBox "本所期限不可大於法定期限!!", vbExclamation
       textNP08.SetFocus
       Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
Dim strCase(1 To 4) As String
Dim tmpArr As Variant
Dim intR As Integer, intA As Integer
Dim strNCP09 As String, strNP22 As String
Dim strNewCPP02 As String
Dim ArrList1
Dim ArrList2
Dim arrList3
Dim arrList4
Dim m_CP12 As String

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
      
   '1.先處理現在的收文
   If textNP07 <> "" And textNP07_2 <> "" Then
       strNP22 = GetNextProgressNo
       strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                "VALUES ('" & m_CP09 & "','" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "'," & textNP07 & "," & _
                          DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & m_CP13 & "'," & strNP22 & ")"
       cnnConnection.Execute strSql
       'Modified by Lydia 2023/07/05 電話通知1727所限和法限同畫面輸入的下一程序的所限和法限
       'strSql = "update caseprogress set cp27=" & m_strCP27 & " , cp64=" & CNULL(ChgSQL(textCP64)) & " where cp09='" & m_CP09 & "' "
       strSql = "update caseprogress set cp27=" & m_strCP27 & ", CP06=" & CNULL(DBDATE(textNP08), True) & ", CP07=" & CNULL(DBDATE(textNP09), True) & ", cp64=" & CNULL(ChgSQL(textCP64)) & " where cp09='" & m_CP09 & "' "
       cnnConnection.Execute strSql
   End If
   '2.多案案號處理
   If cmdInput.Tag <> "" Then
      m_CP12 = GetST15(m_CP13)
      tmpArr = Split(cmdInput.Tag, ",")
      If strCPP02List <> "" Then
          ArrList1 = Split(strCPP02List, vbCrLf)
          ArrList2 = Split(strCPP04List, vbCrLf)
          arrList3 = Split(strDateList, vbCrLf)
          arrList4 = Split(strTimeList, vbCrLf)
      End If
      For intR = 0 To UBound(tmpArr)
         If Trim(tmpArr(intR)) <> "" Then
             Call ChgCaseNo(Trim(tmpArr(intR)), strCase)
             If strCase(1) <> "" And strCase(2) <> "" Then
                 strNCP09 = AutoNo("C", 6)
                 'Modified by Lydia 2023/07/05 新增案件之收文日與主案相同 strSrvDate(1)=>DBDATE(lblFM2(5))
                 'Modified by Lydia 2023/07/05 電話通知1727所限和法限(CP06,CP07)同畫面輸入的下一程序的所限和法限
                 strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP26,CP32, CP64,CP20,CP27) " & _
                            "VALUES ('" & strCase(1) & "','" & strCase(2) & "','" & strCase(3) & "','" & strCase(4) & "'," & DBDATE(lblFM2(5)) & "," & CNULL(DBDATE(textNP08), True) & "," & CNULL(DBDATE(textNP09), True) & "," & _
                            "'" & strNCP09 & "','1727','" & m_CP12 & "','" & m_CP13 & "','" & m_CP13 & "','N','N','" & ChgSQL(textCP64) & "','N', " & m_strCP27 & " ) "
                 cnnConnection.Execute strSql
                 If textNP07 <> "" And textNP07_2 <> "" Then
                    strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                "VALUES ('" & strNCP09 & "','" & strCase(1) & "','" & strCase(2) & "','" & strCase(3) & "','" & strCase(4) & "'," & textNP07 & "," & _
                                 DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & m_CP13 & "'," & strNP22 & ")"
                    cnnConnection.Execute strSql
                 End If
                 '將之前下載檔案上傳到卷宗區
                 If strCPP02List <> "" Then
                     For intA = 0 To UBound(ArrList1)
                        strNewCPP02 = Replace(ArrList1(intA), PUB_FCPCaseNo2FileName(Text1, Text2, Text3, Text4), PUB_FCPCaseNo2FileName(strCase(1), strCase(2), strCase(3), strCase(4)))
                        If SaveAttFile_PDF(strNCP09, m_AttachPath & "\" & ArrList1(intA), strNewCPP02, Val(arrList3(intA)), Val(arrList4(intA)), False, , , , , , "" & ArrList2(intA)) = True Then
                          'bolAdd = True
                        Else
                           GoTo ErrorHandler
                        End If
                     Next intA
                 End If
             End If
         End If
      Next intR
   End If
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
   End If
End Function

Private Sub textCP64_GotFocus()
   TextInverse textCP64
End Sub

Private Sub textNP07_GotFocus()
   TextInverse textNP07
End Sub

Private Sub textNP08_GotFocus()
   TextInverse textNP08
End Sub

Private Sub textNP09_GotFocus()
   TextInverse textNP09
End Sub

' 下一程序
Private Sub textNP07_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNP07) = False Then
      strSql = "SELECT * FROM CasePropertyMap " & _
               "WHERE CPM01 = '" & Text1 & "' AND " & _
                     "CPM02 = '" & textNP07 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         strTit = "檢核資料"
         strMsg = "下一程序代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP07_GotFocus
         rsTmp.Close
         GoTo EXITSUB
      End If
      
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CPM03")) = False Then
         textNP07_2 = rsTmp.Fields("CPM03")
      ElseIf IsNull(rsTmp.Fields("CPM04")) = False Then
         textNP07_2 = rsTmp.Fields("CPM04")
      End If
      rsTmp.Close
            
      textNP08.Locked = False
      textNP08.TabStop = True
      textNP08.BackColor = &H80000005
      textNP09.Locked = False
      textNP09.TabStop = True
      textNP09.BackColor = &H80000005
   Else
      textNP08 = Empty
      textNP08.Locked = True
      textNP08.TabStop = False
      textNP08.BackColor = &H8000000F
      textNP09 = Empty
      textNP09.Locked = True
      textNP09.TabStop = False
      textNP09.BackColor = &H8000000F
   End If
   Set rsTmp = Nothing
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 本所期限
Private Sub textNP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNP08) = False Then
      ' 本所期限日期不正確
      If CheckIsTaiwanDate(textNP08, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08_GotFocus
      Else
          '本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
          textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08, True), 1)
      End If
   End If
End Sub

' 法定期限
Private Sub textNP09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNP09) = False Then
      ' 法定期限日期不正確
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
      'Added by Lydia 2022/09/27 本所期限請依智慧局來函之規則由系統自動帶
      Else
         '台灣案之本所期限設定=法定期限－２個工作天（不含當日）
         textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
      'end 2022/09/27
      End If
   End If
End Sub

'FCT審查機關來電與回覆流程管制：與內商frm020108主管機關來電處理記錄共用
Private Sub cmdInput_Click()
   Set frm880004.mPreForm = Me
   frm880004.iStiu = 8
   frm880004.m_LCV01 = Text1 & Text2 & Text3 & Text4 & "," & lblFM2(2).Caption & "," & lblFM2(1).Caption
   frm880004.m_TempList = Me.cmdInput.Tag
   frm880004.Show vbModal
End Sub

Private Sub textNP15_GotFocus()
    TextInverse textNP15
End Sub
