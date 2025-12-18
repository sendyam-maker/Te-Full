VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人收到卷宗及本所期限輸入"
   ClientHeight    =   3765
   ClientLeft      =   4245
   ClientTop       =   5280
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7050
   Begin VB.CommandButton cmdOK 
      Caption         =   "接洽單"
      Height          =   345
      Index           =   2
      Left            =   4110
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   3990
      MaxLength       =   1
      TabIndex        =   31
      Top             =   1980
      Width           =   315
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   4890
      TabIndex        =   1
      Top             =   150
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   345
      Index           =   1
      Left            =   5730
      TabIndex        =   2
      Top             =   150
      Width           =   1155
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   4950
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2985
      Width           =   915
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   0
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2430
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "一案兩申請(Y：是)"
      Height          =   255
      Left            =   4380
      TabIndex        =   32
      Top             =   2010
      Width           =   1575
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   10
      Left            =   1080
      TabIndex        =   30
      Top             =   3285
      Width           =   915
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   29
      Top             =   3000
      Width           =   1590
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   28
      Top             =   2730
      Width           =   2340
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4128;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "收卷註記：              (Y：收到卷宗)"
      Height          =   255
      Index           =   27
      Left            =   3960
      TabIndex        =   27
      Top             =   2430
      Width           =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "請承辦人輸入本所期限："
      Height          =   420
      Index           =   24
      Left            =   3800
      TabIndex        =   26
      Top             =   2880
      Width           =   1095
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   7
      Left            =   1515
      TabIndex        =   25
      Top             =   2430
      Width           =   1410
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2487;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   1515
      TabIndex        =   24
      Top             =   2145
      Width           =   600
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1058;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   23
      Top             =   1845
      Width           =   2700
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4762;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   22
      Top             =   1560
      Width           =   1830
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3228;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   21
      Top             =   1260
      Width           =   1590
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   20
      Top             =   990
      Width           =   1590
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "(N：不算)"
      Height          =   255
      Index           =   32
      Left            =   2130
      TabIndex        =   19
      Top             =   2145
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "點　數　："
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   18
      Top             =   3285
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   17
      Top             =   2715
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "專利/商標種類："
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Top             =   2430
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "是否算案件數："
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   15
      Top             =   2145
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   14
      Top             =   1845
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收文日　："
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   12
      Top             =   1260
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   11
      Top             =   975
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   10
      Top             =   690
      Width           =   735
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   690
      Width           =   630
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1111;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   4950
      TabIndex        =   8
      Top             =   690
      Width           =   1590
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "目　　次："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   690
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   960
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
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   945
   End
End
Attribute VB_Name = "frm090201_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; lbl1(index)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim strTot(0 To 500) As String
Dim IntNow As Integer, IntTot As Integer
Dim strSql As String, i As Integer, s As Integer
Dim m_CP06 As String  '程序輸入之本所期限
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
'add by nick 2004/07/16 是否有一案2 請
Dim IsTwo As Boolean
Dim m_CP140 As String 'Add By Sindy 2015/6/24
Dim m_AttachPath As String 'Add By Sindy 2015/6/24


Sub Process(intSitu As Integer)
Dim strText As String

   '總收文號
   strText = strTot(intSitu)
   'Add By Sindy 2015/6/24 +CP140
                        strSql = "SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,PA57,CP06,CP140 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr1 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,DECODE(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,TM29,CP06,CP140 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,TRADEMARK WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr2 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',                            decode(lc15,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,LC08,CP06,CP140 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,LAWCASE WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr3 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,                    CP26,'',                            CPM03,                         S5.ST02,CP18,EP27,EP31,CP10,HC09,CP06,CP140 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,HIRECASE WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr4 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',                            decode(sp09,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,SP15,CP06,CP140 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,SERVICEPRACTICE WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         .MoveFirst
         m_CP140 = "" & .Fields("CP140") 'Add By Sindy 2015/6/24
         For i = 0 To 10
             Lbl1(i) = CheckStr(.Fields(i))
             Lbl1(i).BackColor = &H8000000F 'Added by Lydia 2021/12/21
         Next i
            'add by nick 2004/07/09 一案兩申請
            If SystemNumber(Lbl1(4).Caption, 1) = "P" Or SystemNumber(Lbl1(4).Caption, 1) = "CFP" Then
                Text1.Visible = True
                Label2.Visible = True
                    'add by nick 2004/07/09 檢查一案兩申請
                      strSql = "select ep27 from engineerprogress where ep02 in ("
                      strSql = strSql & " select min(cp09) from casemap,caseprogress where cm01='" & SystemNumber(Lbl1(4).Caption, 1) & "' and cm02='" & SystemNumber(Lbl1(4).Caption, 2) & "' and cm03='" & SystemNumber(Lbl1(4).Caption, 3) & "' and cm04='" & SystemNumber(Lbl1(4).Caption, 4) & "' and cm10='3' and cm05=cp01 and cm06=cp02 and cm07=cp03 and cm08=cp04   " & _
                                     "union select min(cp09) from casemap,caseprogress where cm05='" & SystemNumber(Lbl1(4).Caption, 1) & "' and cm06='" & SystemNumber(Lbl1(4).Caption, 2) & "' and cm07='" & SystemNumber(Lbl1(4).Caption, 3) & "' and cm08='" & SystemNumber(Lbl1(4).Caption, 4) & "' and cm10='3' and cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04   "
                      strSql = strSql & ") "
                     CheckOC3
                     With AdoRecordSet3
                         .CursorLocation = adUseClient
                         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                         If .RecordCount <> 0 Then
                                IsTwo = True
                                Text1.Text = "Y"
                         Else
                                IsTwo = False
                                Text1.Text = ""
                         End If
                     End With
         Else
            Text1.Visible = False
            Label2.Visible = False
         End If
         
         If IsNull(.Fields(14).Value) <> 0 Then
             Me.lblClose.Caption = ""
         Else
             Me.lblClose.Caption = "已閉卷"
         End If
         If IsNull(.Fields("EP27")) Then
            txt1(0) = ""
         Else
            txt1(0) = "Y"
         End If
         If IsNull(.Fields("EP31")) Then
            txt1(1) = ""
         Else
            txt1(1) = .Fields("EP31")
         End If
         If IsNull(.Fields("CP06")) Then
            m_CP06 = ""
         Else
            m_CP06 = .Fields("CP06")
         End If
      Else
         For i = 0 To 10
             Lbl1(i) = ""
             Lbl1(i).BackColor = &H8000000F 'Added by Lydia 2021/12/21
         Next i
         Me.lblClose.Caption = ""
         m_CP06 = ""
         txt1(0) = ""
         txt1(1) = ""
      End If
   End With
   CheckOC
   IntNow = IntNow + 1
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim rsA As New ADODB.Recordset
Dim stFileName As String
Dim hLocalFile As Long
   
On Error GoTo CheckingErr
   
   Select Case Index
      Case 0 '確定
         Screen.MousePointer = vbHourglass
         '重新檢查欄位有效性
         If TxtValidate = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
'         'add by nick 2004/07/09 檢查一案兩申請
'         If Text1.Visible = True Then
'         strSQL = "select cp09 from casemap,caseprogress where cm01='" & SystemNumber(lbl1(4).Caption, 1) & "' and cm02='" & SystemNumber(lbl1(4).Caption, 2) & "' and cm03='" & SystemNumber(lbl1(4).Caption, 3) & "' and cm04='" & SystemNumber(lbl1(4).Caption, 4) & "' and cm10='3' and cm05=cp01 and cm06=cp02 and cm07=cp03 and cm08=cp04 " & _
'                        "union select cp09 from casemap,caseprogress where cm05='" & SystemNumber(lbl1(4).Caption, 1) & "' and cm06='" & SystemNumber(lbl1(4).Caption, 2) & "' and cm07='" & SystemNumber(lbl1(4).Caption, 3) & "' and cm08='" & SystemNumber(lbl1(4).Caption, 4) & "' and cm10='3' and cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04 "
'        CheckOC
'        With adoRecordset
'            .CursorLocation = adUseClient
'            .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If .RecordCount <> 0 And .RecordCount > 0 Then
                   If Text1.Text = "" And IsTwo = True Then
                       MsgBox "此案有一案兩申請，請核取！", , "User 輸入錯誤！"
                       Screen.MousePointer = vbDefault
                       Exit Sub
                    End If
'            Else
                   If (Text1.Text = "Y" Or Text1.Text = "y") And IsTwo = False Then
                       MsgBox "此案沒有一案兩申請，請取消核取！", , "User 輸入錯誤！"
                       Screen.MousePointer = vbDefault
                       Exit Sub
                    End If
'            End If
'        End With
'       End If
         'add end 2004/07/09
         If txt1(0) <> "" Or txt1(1) <> "" Then
            DoEvents
            cnnConnection.BeginTrans
            'edit by nick 2004/07/09 修正
            'strSQL = "Update EngineerProgress Set EP27=" & IIf(ChangeTStringToWString(txt1(0)) = "", "NULL", ServerDate) & ",EP31=" & IIf(ChangeTStringToWString(txt1(1)) = "", "NULL", ChangeTStringToWString(txt1(1))) & " Where EP02='" & lbl1(2).Caption & "' "
            'add by nick 2004/07/09 當一案2 申請時，要順便更新相關案
            If Text1.Visible = True And Text1.Text = "Y" Then
            ' edit by nick 2004/07/16 不用修正
'                 If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'                     strSQL = "Update EngineerProgress Set EP27=" & IIf(Txt1(0) = "", "NULL", ServerDate) & ",EP31=" & IIf(ChangeTStringToWString(Txt1(1)) = "", "NULL", ChangeTStringToWString(Txt1(1))) & " Where EP02='" & CheckStr(adoRecordset.Fields(0).Value) & "' "
'                     cnnConnection.Execute strSQL
'                 End If
            End If
            strSql = "Update EngineerProgress Set EP27=" & IIf(txt1(0) = "", "NULL", strSrvDate(1)) & ",EP31=" & IIf(ChangeTStringToWString(txt1(1)) = "", "NULL", ChangeTStringToWString(txt1(1))) & " Where EP02='" & Lbl1(2).Caption & "' "
            cnnConnection.Execute strSql
            cnnConnection.CommitTrans
         End If
         'add by nick 2004/07/09
         CheckOC
         '下一筆
         If IntNow <> IntTot Then
            txt1(0).SetFocus
            Process IntNow
         Else
            Unload frm090201_6
            Screen.MousePointer = vbHourglass
            frm090201_5.RefreshData
            Screen.MousePointer = vbDefault
            If frm090201_5.TextOk = True Then
               frm090201_5.Show
            'Add by Morgan 2010/11/17
            Else
               frm090201_5.strContinue = False
               Unload frm090201_5
            'end 2010/11/17
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1 '回前畫面
         Screen.MousePointer = vbHourglass
         frm090201_5.RefreshData
         Screen.MousePointer = vbDefault
         If frm090201_5.TextOk = True Then frm090201_5.Show
         Unload Me
      Case 2 '接洽單
         Screen.MousePointer = vbHourglass
         If m_CP140 <> "" Then
            '查詢接洽記錄單
            'Modify By Sindy 2022/12/23 改用共用函數
            Call PUB_Queryfrm090801(m_CP140, DBDATE(Lbl1(3).Caption), Me)
'            'Modify By Sindy 2022/9/5
'            If DBDATE(Replace(lbl1(3).Caption, "/", "")) >= 接洽單電子收文啟用日 Then
'               '查詢接洽記錄單
'               frm090801_Q.SetParent Me
'               frm090801_Q.m_blnCallPrint = True
'               frm090801_Q.Text5 = m_CP140
'               Call frm090801_Q.cmdOK_Click(4)
'               frm090801_Q.Show
'            Else
'            '2022/9/5 END
'               frm090801.SetParent Me
'               frm090801.m_blnCallPrint = True 'Add By Sindy 2022/10/19
'               frm090801.Text5 = m_CP140
'               frm090801.m_blnCallPrint_CRL119 = True '是否列印特殊收據頁
'               Call frm090801.cmdOK_Click(4)
'               frm090801.cmdOK(2).Visible = False
'               frm090801.cmdOK(0).Visible = False
'               frm090801.txtPCnt.Visible = False
'            End If
'            Me.Hide
            '2022/12/23 END
            txt1(0) = "Y" 'Add By Sindy 2021/2/18 改為開啟接洽單後，系統自動上收卷註記。但還是要按確認鍵一併更新,因有其他資料。
            '2022/12/23 END
         Else
            '檢查是否有接洽單.pdf
            strExc(0) = "select *" & _
                        " From casepaperpdf" & _
                        " where cpp01='" & Lbl1(2) & "' and instr(upper(cpp02),upper('" & EMP_接洽單 & ".pdf'))>0 and cpp10<>'D'"
            rsA.CursorLocation = adUseClient
            rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               '讀取檔案名稱
               stFileName = rsA.Fields("cpp02")
      '         If GetAttachFile_CPP(m_EEP01, stFileName, m_AttachPath & "\" & stFileName) = False Then
      '            MsgBox "無法儲存欲開啟的檔案[ " & stFileName & " ]！"
      '         End If
               If PUB_GetAttachFile_CPP(Lbl1(2), stFileName, m_AttachPath) = True Then
                  '開啟檔案
                  ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
                  txt1(0) = "Y" 'Add By Sindy 2021/2/18 改為開啟接洽單後，系統自動上收卷註記。但還是要按確認鍵一併更新,因有其他資料。
               End If
            Else
               MsgBox "無接洽單！"
            End If
            rsA.Close
            Set rsA = Nothing
         End If
         Screen.MousePointer = vbDefault
   End Select
Exit Sub
CheckingErr:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_SqlGrpStr1 = SQLGrpStr("", 1)
   m_SqlGrpStr2 = SQLGrpStr("", 2)
   m_SqlGrpStr3 = SQLGrpStr("", 3)
   m_SqlGrpStr4 = SQLGrpStr("", 4)
   m_SqlGrpStr5 = SQLGrpStr("", 5)
   With frm090201_5.GRD1
      IntTot = 0
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "V" Then
            strTot(IntTot) = .TextMatrix(i, 22) '收文號
            IntTot = IntTot + 1
         End If
      Next
   End With
   IntNow = 0
   Process IntNow
   'Modify by Amy 2014/09/22 取消工程師輸入本所期限
   Label1(24).Visible = False
   txt1(1).Visible = False
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2015/6/24
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q") = True Then
      Unload frm090801_Q
   End If
   '2022/12/17 END
   
   Set frm090201_6 = Nothing
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
      Select Case Trim(Text1)
      Case "Y", ""
      Case Else
         s = MsgBox("一案兩申請只能輸入 Y !!", , "USER 輸入錯誤")
         Text1.SetFocus
         Text1.SelStart = 0
         Text1.SelLength = Len(Text1)
         Cancel = True
         Exit Sub
      End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
   Case 0 '收卷註記
      Select Case Trim(txt1(0))
      Case "Y", ""
      Case Else
         s = MsgBox("收卷註記只能輸入 Y !!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         txt1(0).SelStart = 0
         txt1(0).SelLength = Len(txt1(0))
         Cancel = True
         Exit Sub
      End Select
   'Mark by Amy 2014/09/22 取消工程師輸入本所期限
'   Case 1 '工程師輸入本所期限
'      If Len(Txt1(Index)) <> 0 Then
'         If CheckIsTaiwanDate(Txt1(Index).Text) = False Then
'            Txt1(Index).SetFocus
'            Txt1(Index).SelLength = Len(Txt1(Index))
'            Cancel = True
'            Exit Sub
'         End If
'         If Txt1(Index).Text <> ChangeWStringToTString(m_CP06) Then
'            s = MsgBox("輸入本所期限與程序輸入本所期限不同!!", , "USER 輸入錯誤")
'            Txt1(Index).SetFocus
'            Txt1(Index).SelLength = Len(Txt1(Index))
'            Cancel = True
'            Exit Sub
'         End If
'         Txt1(0) = "Y"
'      End If
   'end2014/09/22
   End Select
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Me.txt1
      If objTxt.Enabled = True Then
         Cancel = False
         txt1_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   TxtValidate = True
End Function


