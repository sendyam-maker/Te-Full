VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm077003_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "放棄案源"
   ClientHeight    =   3990
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   6980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6980
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   0
      Left            =   5760
      TabIndex        =   2
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "確認放棄"
      Height          =   350
      Index           =   1
      Left            =   4530
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   60
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   90
      TabIndex        =   3
      Top             =   450
      Width           =   6825
      Begin MSForms.ListBox lstUsers 
         Height          =   825
         Left            =   4530
         TabIndex        =   4
         Top             =   570
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;1455"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtContent 
         Height          =   765
         Left            =   1260
         TabIndex        =   16
         Top             =   1470
         Width           =   5445
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "9604;1349"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAbortReason 
         Height          =   645
         Left            =   1260
         TabIndex        =   0
         Top             =   2730
         Width           =   5445
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "9604;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主題："
         Height          =   180
         Index           =   3
         Left            =   3120
         TabIndex        =   20
         Top             =   2370
         Width           =   540
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   6
         Left            =   3690
         TabIndex        =   19
         Top             =   2370
         Width           =   2985
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "5265;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   1260
         TabIndex        =   18
         Top             =   2370
         Width           =   1365
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "2408;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法律所案號："
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   2370
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "放棄原因："
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   2730
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹內容："
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   1470
         Width           =   900
      End
      Begin MSForms.Label Label2 
         Height          =   510
         Index           =   4
         Left            =   1260
         TabIndex        =   13
         Top             =   870
         Width           =   3105
         BackColor       =   16777215
         Size            =   "5477;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   1260
         TabIndex        =   12
         Top             =   570
         Width           =   1725
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "3043;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹人："
         Height          =   180
         Index           =   9
         Left            =   3510
         TabIndex        =   11
         Top             =   570
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "管制日期："
         Height          =   180
         Index           =   5
         Left            =   3510
         TabIndex        =   10
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業務區："
         Height          =   180
         Index           =   10
         Left            =   150
         TabIndex        =   9
         Top             =   570
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹日期："
         Height          =   180
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   900
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   7
         Top             =   270
         Width           =   1725
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "3043;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹客戶："
         Height          =   180
         Index           =   11
         Left            =   150
         TabIndex        =   6
         Top             =   870
         Width           =   900
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   4530
         TabIndex        =   5
         Top             =   270
         Width           =   1725
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "3043;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   22
      Top             =   120
      Width           =   1845
      BackColor       =   16777215
      VariousPropertyBits=   27
      Size            =   "3254;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總收文號："
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   21
      Top             =   157
      Width           =   900
   End
End
Attribute VB_Name = "frm077003_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; lstUsers、Label2(indes)、txtContent、txtAbortReason
'Created by Morgan 2020/4/24
Option Explicit

Public strLOS15 As String
Dim strMailTo As String
Dim strLOS02 As String, strLOS10 As String 'Added by Morgan 2022/6/9
Dim strCRL01 As String 'Add By Sindy 2022/10/3


Private Sub cmdOK_Click(Index As Integer)
   If Index = 1 Then
      If Trim(txtAbortReason) = "" Then
         MsgBox "請輸入放棄原因！", vbExclamation
         txtAbortReason.SetFocus
         Exit Sub
      End If
      'Added by Lydia 2022/02/11 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True, "TextBox") = False Then
          Exit Sub
      End If
      'end 2022/02/11
      If FormSave = False Then
         Exit Sub
      End If
   End If
   frm077003.iReturn = Index
   Unload Me
End Sub

Private Sub Form_Load()
   Dim oLabel As Control
   
   MoveFormToCenter Me, True
   
   For Each oLabel In Label2
      oLabel.BackColor = &H8000000F
   Next
   
   ReadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache
   Set frm077003_1 = Nothing
End Sub

Private Sub ReadData()
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   'Modify By Sindy 2022/10/3 + ,CRL01
   stSQL = "select los01 ""總收文號"",sqldatet(los12) ""介紹日期"",sqldatet(cp06) ""管制日期"" " & _
      ",a0902 ""業務區"",los04 ""介紹人"",NVL(CRA07,CRA08) ""介紹客戶"",CRL57 ""介紹內容"" " & _
      ",CRL07||decode(CRL08,'','新案','-'||CRL08||decode(CRL09||CRL10,'000','','-'||CRL09||'-'||CRL10)) ""法律所案號""" & _
      ",CRL17 ""案件名稱"",los02,los10,CRL01 from lawofficesource,caseprogress,acc090,ConsultRecordList,ConsultRecApp" & _
      " where los15='" & strLOS15 & "' and cp09(+)=los01 and a0901(+)=cp12" & _
      " and crl01(+)=los17 and cra01(+)=crl01"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With RsQ
      Label2(0) = "" & .Fields("總收文號")
      Label2(1) = "" & .Fields("介紹日期")
      Label2(2) = "" & .Fields("管制日期")
      Label2(3) = "" & .Fields("業務區")
      Label2(4) = "" & .Fields("介紹客戶")
      txtContent = "" & .Fields("介紹內容")
      Label2(5) = "" & .Fields("法律所案號")
      Label2(6) = "" & .Fields("案件名稱")
      strMailTo = "" & .Fields("介紹人")
      'Added by Morgan 2022/6/10
      strLOS02 = "" & .Fields("los02")
      strLOS10 = "" & .Fields("los10")
      'end 2022/6/10
      strCRL01 = "" & .Fields("CRL01") 'Add By Sindy 2022/10/3
      SetlstUsers strMailTo
      End With
   End If
   Set RsQ = Nothing
End Sub

Private Sub SetlstUsers(p_stNums As String)
   lstUsers.Clear
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0 order by instr('" & p_stNums & "',st01) desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         .MoveFirst
         Do While Not .EOF
            lstUsers.AddItem "" & .Fields("st02"), 0
            .MoveNext
         Loop
         End With
      End If
   End If
End Sub

Private Sub txtAbortReason_GotFocus()
   TextInverse txtAbortReason
End Sub

Private Function FormSave() As Boolean
   Dim strSubject As String, strContent As String, strCC As String, strPS As String
   Dim strCP10 As String 'Add By Sindy 2023/12/13
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   strSql = "update lawofficesource set los06=los06  where los15='" & strLOS15 & "' and los06||los07 is not null"
   cnnConnection.Execute strSql, intI
   If intI <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox "該案源已收文/放棄，不可放棄！", vbExclamation
      Exit Function
   End If
   
   'Added by Morgan 2022/6/10
   'TT案取消收文
   If strLOS10 <> "" Then
      strSql = "update caseprogress set cp27=null,cp57=" & strSrvDate(1) & ",cp58='99'" & _
         ",cp64='放棄原因:" & ChgSQL(txtAbortReason) & ";'||cp64" & _
         " where cp09='" & strLOS10 & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2022/6/10
      
   strSql = "update lawofficesource set los07=" & strSrvDate(1) & _
      ",los08='" & strUserNum & "',los09='" & ChgSQL(txtAbortReason) & "'" & _
      " where los15='" & strLOS15 & "'"
   cnnConnection.Execute strSql, intI
   
   strSubject = Label2(1) & "介紹之案源" & IIf(InStr(Label2(5), "-") > 0, "(" & Label2(5) & ")", "") & "經律師評量後,認定不宜承辦,特此通知."
   strContent = "法律所案號：" & Label2(5) & vbCrLf & _
      "主題：" & Label2(6) & vbCrLf & vbCrLf & _
      "放棄原因：" & txtAbortReason & vbCrLf & vbCrLf & _
      "介紹人：　" & GetUsers() & vbCrLf & _
      "介紹日期：" & Label2(1) & vbCrLf & _
      "介紹客戶：" & Label2(4) & vbCrLf & _
      "介紹內容：" & txtContent
   
   strCC = PUB_GetLos04Man(strMailTo)
   
   'Modified by Morgan 2022/6/9 B1類才需要自動取消收文,另外再告知智權會自動銷案並CC給承辦
   If strLOS02 = "B1" Then
      
      'P/T案取消收文
      strSql = "update caseprogress set cp57=" & strSrvDate(1) & ",cp58='99'" & _
         ",cp64='放棄原因:" & ChgSQL(txtAbortReason) & ";'||cp64" & _
         " where cp162='" & strLOS15 & "'"
      cnnConnection.Execute strSql, intI
            
      strExc(0) = "select distinct cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,cp14,cp01,cp02,cp03,cp04" & _
         " from caseprogress where cp162='" & strLOS15 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         '閉卷
         If InStr(.Fields("cp01"), "T") > 0 Then
            strSql = "update trademark set tm29='Y',tm30=" & strSrvDate(1) & ",tm31='99'" & _
               ",tm58='放棄原因:" & ChgSQL(txtAbortReason) & ";'||tm58" & _
               " where tm01='" & .Fields("cp01") & "' and tm02='" & .Fields("cp02") & "'" & _
               " and tm03='" & .Fields("cp03") & "' and tm04='" & .Fields("cp04") & "'" & _
               " and not exists(select * from caseprogress where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp57 is null)"
         Else
            strSql = "update patent set pa57='Y',pa58=" & strSrvDate(1) & ",pa59='99'" & _
               ",pa91='放棄原因:" & ChgSQL(txtAbortReason) & ";'||pa91" & _
               " where pa01='" & .Fields("cp01") & "' and pa02='" & .Fields("cp02") & "'" & _
               " and pa03='" & .Fields("cp03") & "' and pa04='" & .Fields("cp04") & "'" & _
               " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp57 is null)"
         End If
         cnnConnection.Execute strSql, intI
         
         strContent = strContent & vbCrLf & vbCrLf & "※" & .Fields("CNo") & "已自動銷案！"
         Do While Not .EOF
            If Not IsNull(.Fields("cp14")) Then
               strCC = .Fields("cp14") & ";" & strCC
            End If
            .MoveNext
         Loop
         End With
      End If
   End If
   'end 2022/6/9
   
   'Add By Sindy 2022/10/3
   If strCRL01 <> "" Then
      strExc(0) = "select * from flow002 where F0201='" & strCRL01 & "' and F0202='A4'" 'and F0207 is null
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '簽核檔-放棄
         strSql = "update FLOW002 set " & _
                  "F0205='" & strSrvDate(1) & "'" & _
                  ",F0206='" & Right("000000" & ServerTime, 6) & "'" & _
                  ",F0207='4',F0204='" & strUserNum & "'" & _
                  " where F0201='" & strCRL01 & "' and F0202='A4' and F0207 is null "
         cnnConnection.Execute strSql
         '表單主檔
         strSql = "update FLOW003 set " & _
                  "F0307='" & strUserNum & "'" & _
                  ",F0308=F0316" & _
                  ",F0309='" & Flow_放棄案源 & "'" & _
                  " where F0301='" & strCRL01 & "'"
         cnnConnection.Execute strSql
         'Add By Sindy 2023/8/8
         '流程備註檔
         strSql = GetInsertFLOW004Sql(strCRL01, strUserNum, strSrvDate(1), Right("000000" & ServerTime, 6), Flow_放棄案源, ChgSQL(Trim(txtAbortReason.Text)))
         cnnConnection.Execute strSql
         '2023/8/8 END
         'Add By Sindy 2023/12/13 有附件則歸到TT總收文號
         strExc(0) = "select cp09,cp10 from caseprogress where cp09='" & strLOS10 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCP10 = RsTemp.Fields("cp10")
         End If
         strSql = "update casepaperpdf set cpp01='" & strLOS10 & "',cpp10='X'" & _
                  ",cpp02=replace(cpp02,'" & strCRL01 & "','TT999999." & strCP10 & "') where cpp11='" & strCRL01 & "'"
         cnnConnection.Execute strSql, intI
         '2023/12/13 END
      End If
   End If
   '2022/10/3 END
   
   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " values ('" & strUserNum & "','" & strMailTo & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
            ",'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "','" & strCC & "')"
   cnnConnection.Execute strSql, intI
      
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Function GetUsers() As String
   Dim strNameList As String, ii As Integer
   
   For ii = 0 To lstUsers.ListCount - 1
      strNameList = strNameList & " " & lstUsers.List(ii)
   Next
   GetUsers = strNameList
End Function
