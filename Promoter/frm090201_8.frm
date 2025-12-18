VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人 預定會稿日 輸入"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7050
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   5925
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2100
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   5790
      TabIndex        =   2
      Top             =   75
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4980
      TabIndex        =   1
      Top             =   75
      Width           =   800
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   11
      Left            =   4875
      TabIndex        =   29
      Top             =   1800
      Width           =   1920
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3387;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   0
      Left            =   3705
      TabIndex        =   28
      Top             =   1800
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
      Left            =   4020
      TabIndex        =   27
      Top             =   1515
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   255
      Index           =   12
      Left            =   180
      TabIndex        =   26
      Top             =   2955
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "目　　次："
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   25
      Top             =   645
      Width           =   960
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   4875
      TabIndex        =   24
      Top             =   645
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
      Index           =   0
      Left            =   1140
      TabIndex        =   23
      Top             =   645
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
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   255
      Index           =   8
      Left            =   4020
      TabIndex        =   22
      Top             =   645
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   255
      Index           =   21
      Left            =   180
      TabIndex        =   21
      Top             =   930
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收文日　："
      Height          =   255
      Index           =   20
      Left            =   180
      TabIndex        =   20
      Top             =   1215
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   19
      Left            =   180
      TabIndex        =   19
      Top             =   1515
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   18
      Left            =   180
      TabIndex        =   18
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "是否算案件數："
      Height          =   255
      Index           =   17
      Left            =   180
      TabIndex        =   17
      Top             =   2100
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "專利/商標種類："
      Height          =   255
      Index           =   16
      Left            =   180
      TabIndex        =   16
      Top             =   2385
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   15
      Top             =   2670
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "點　數　："
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   14
      Top             =   3240
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "(N：不算)"
      Height          =   255
      Index           =   32
      Left            =   2190
      TabIndex        =   13
      Top             =   2100
      Width           =   1065
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   12
      Top             =   945
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
      Index           =   3
      Left            =   1140
      TabIndex        =   11
      Top             =   1215
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
      Index           =   4
      Left            =   1140
      TabIndex        =   10
      Top             =   1515
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
      Index           =   5
      Left            =   1140
      TabIndex        =   9
      Top             =   1800
      Width           =   2400
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4233;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   1575
      TabIndex        =   8
      Top             =   2100
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
      Index           =   7
      Left            =   1575
      TabIndex        =   7
      Top             =   2385
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "請承辦人輸入 預定會稿日："
      Height          =   255
      Index           =   24
      Left            =   3690
      TabIndex        =   6
      Top             =   2100
      Width           =   2205
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   1140
      TabIndex        =   5
      Top             =   2685
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
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   9
      Left            =   1140
      TabIndex        =   4
      Top             =   2955
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
      Index           =   10
      Left            =   1140
      TabIndex        =   3
      Top             =   3240
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
End
Attribute VB_Name = "frm090201_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; lbl1(index)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'create by nickc 2006/02/08 copy from frm090201_6
Option Explicit
Dim strTot(0 To 500) As String
Dim IntNow As Integer, IntTot As Integer
Dim strSql As String, i As Integer, s As Integer
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
Dim m_CP48 As String  '紀錄該筆承辦期限
Dim P_DateLine As String
Dim CFP_DateLine As String
Dim m_CP13 As String

Sub Process(intSitu As Integer)
Dim strText As String

   '總收文號
   strText = strTot(intSitu)
   
                        strSql = "SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,PA57,CP06,cp48,cp13 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr1 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,DECODE(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,TM29,CP06,cp48,cp13 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,TRADEMARK WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr2 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',                            decode(lc15,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,LC08,CP06,cp48,cp13 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,LAWCASE WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr3 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,                    CP26,'',                            CPM03,                         S5.ST02,CP18,EP27,EP31,CP10,HC09,CP06,cp48,cp13 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,HIRECASE WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr4 & ") "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02,CP09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',                            decode(sp09,'000',cpm03,cpm04),S5.ST02,CP18,EP27,EP31,CP10,SP15,CP06,cp48,cp13 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,SERVICEPRACTICE WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         .MoveFirst
         For i = 0 To 10
             lbl1(i) = CheckStr(.Fields(i))
             lbl1(i).BackColor = &H8000000F 'Added by Lydia 2021/12/21
         Next i
         If IsNull(.Fields(14).Value) <> 0 Then
             Me.lblClose.Caption = ""
         Else
             Me.lblClose.Caption = "已閉卷"
         End If
'         If IsNull(.Fields("EP31")) Then
'            Txt1(0) = ""
'         Else
'            Txt1(0) = .Fields("EP31")
'         End If
         m_CP48 = CheckStr(.Fields("cp48"))
         lbl1(11).Caption = ChangeWStringToTDateString(m_CP48)
         P_DateLine = CompWorkDay(5, m_CP48, 0)
         CFP_DateLine = CompWorkDay(10, m_CP48, 0)
         m_CP13 = CheckStr(.Fields("cp13"))
      Else
         For i = 0 To 11
             lbl1(i) = ""
             lbl1(i).BackColor = &H8000000F 'Added by Lydia 2021/12/21
         Next i
         Me.lblClose.Caption = ""
         Txt1(0) = ""
         P_DateLine = ""
         CFP_DateLine = ""
         m_CP13 = ""
      End If
   End With
   CheckOC
   IntNow = IntNow + 1
End Sub

Private Sub cmdOK_Click(Index As Integer)
On Error GoTo CheckingErr
   Select Case Index
      Case 0 '確定
         Screen.MousePointer = vbHourglass
         '重新檢查欄位有效性
         If TxtValidate = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         If Txt1(0) <> "" Then
            DoEvents
            cnnConnection.BeginTrans
            strSql = "Update EngineerProgress Set EP28=" & IIf(Txt1(0) = "", "NULL", ChangeTStringToWString(Txt1(0))) & " Where EP02='" & lbl1(2).Caption & "' "
            cnnConnection.Execute strSql
            cnnConnection.CommitTrans
            '輸入預定會稿日要發 mail
            If m_CP13 <> "" And Trim(Txt1(0)) <> "" Then
                PUB_SendMail strUserNum, m_CP13, "", lbl1(4) & "已輸入預定會稿日！", "本所案號：" & lbl1(4) & vbCrLf & "收文號：" & lbl1(2) & vbCrLf & "案件名稱：" & lbl1(5) & vbCrLf & "預定會稿日：" & ChangeTStringToTDateString(Txt1(0)), ""
            End If
         End If
         CheckOC
         '下一筆
         If IntNow <> IntTot Then
            Txt1(0).SetFocus
            Process IntNow
         Else
            Unload frm090201_8
            Screen.MousePointer = vbHourglass
            frm090201_7.RefreshData
            Screen.MousePointer = vbDefault
            If frm090201_7.TextOk = True Then frm090201_7.Show
         End If
         Screen.MousePointer = vbDefault
      Case 1 '回前畫面
         Screen.MousePointer = vbHourglass
         frm090201_7.RefreshData
         Screen.MousePointer = vbDefault
         If frm090201_7.TextOk = True Then frm090201_7.Show
         Unload Me
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
   With frm090201_7.GRD1
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090201_8 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   Txt1(Index).SelStart = 0
   Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
   Case 0 '預定會稿日
      If Len(Txt1(Index)) <> 0 Then
         If ChkWork(ChangeTStringToWString(Txt1(Index))) = False Then
            Txt1(Index).SetFocus
            Txt1(Index).SelStart = 0
            Txt1(Index).SelLength = Len(Txt1(Index))
            Cancel = True
            Exit Sub
         End If
         If (SystemNumber(lbl1(4), 1) = "P" And ChangeTStringToWString(Txt1(Index)) > P_DateLine) Or (SystemNumber(lbl1(4), 1) = "CFP" And ChangeTStringToWString(Txt1(Index)) > CFP_DateLine) Then
            '2008/8/25 modify by sonia 王協理操作不檢查
            'Modified by Lydia 2023/04/24 修改王副總退休之相關控制
            'If strUserNum <> "71011" Then
            If strSrvDate(1) >= "20230501" Then
               strExc(1) = "71011;73022"
            Else
               strExc(1) = "71011"
            End If
            If InStr(strExc(1), strUserNum) = 0 Then
            'end 2023/04/24
               MsgBox "P 案上限 5 工作天，CFP 案上限 10 作天！", vbCritical, "錯誤！"
               Txt1(Index).SetFocus
               Txt1(Index).SelStart = 0
               Txt1(Index).SelLength = Len(Txt1(Index))
               Cancel = True
               Exit Sub
            End If
         End If
         If CheckIsTaiwanDate(Txt1(Index).Text) = False Then
            Txt1(Index).SetFocus
            Txt1(Index).SelStart = 0
            Txt1(Index).SelLength = Len(Txt1(Index))
            Cancel = True
            Exit Sub
         End If
      End If
   End Select
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Me.Txt1
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




