VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090908 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專新案認領區"
   ClientHeight    =   4788
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   9048
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4788
   ScaleWidth      =   9048
   Begin VB.ComboBox Combo2 
      Height          =   276
      ItemData        =   "frm090908.frx":0000
      Left            =   5088
      List            =   "frm090908.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   4422
      Width           =   3060
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "外專新案認領-批次"
      Height          =   405
      Left            =   1020
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   270
      Left            =   2784
      TabIndex        =   10
      Text            =   "txtInput"
      Top             =   4392
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "原始檔(&P)"
      Height          =   400
      Index           =   2
      Left            =   5367
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&E)"
      Height          =   400
      Index           =   1
      Left            =   3390
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   795
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGRD1 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   6160
      _Version        =   393216
      Cols            =   11
      AllowUserResizing=   3
      FormatString    =   "V|認領(Y/N)| 認領期限| 認領組|  認領人員 | 不認領組   | 收 文 日 |本所案號    | 案 件 性 質|  譯畢期限|急件分組"
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區(&C)"
      Height          =   400
      Index           =   4
      Left            =   7114
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   3
      Left            =   6188
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   4170
      TabIndex        =   2
      Top             =   120
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   7935
      TabIndex        =   1
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色說明："
      Height          =   180
      Left            =   4152
      TabIndex        =   13
      Top             =   4470
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "符號說明：▲非英說案"
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   2430
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      TabIndex        =   9
      Top             =   150
      Width           =   1800
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3175;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblMemo 
      Caption         =   "認領狀態：Y認領，N不認領"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   570
      Width           =   2355
   End
   Begin VB.Label lblSname 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   230
      Width           =   900
   End
End
Attribute VB_Name = "frm090908"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Lydia 2022/02/22 外專新案認領區
Option Explicit
Public cmdState As Integer
Public mRole As String   'M-主管核判
Private Const cntFixed As Integer = 5
Dim intLastRow As Integer
Dim nRow As Integer '本次點選列數
Dim nCol As Integer '本次點選欄數
Dim m_UserNo As String, m_UserName As String, m_UserSt16 As String '操作員工編號,名稱,工程師組別
Dim rsAD As New ADODB.Recordset
Dim intP As Integer, strP1 As String
Dim mGrpName(1 To 4) '工程師組別名稱
Dim colCp09 As Integer, colInNew As Integer, colInOLD As Integer
Dim colTCN23 As Integer, colTCN20 As Integer, colPA10 As Integer
Dim colCaseNo As Integer, colCP01 As Integer, colCP02 As Integer, colCP03 As Integer, colCP04 As Integer, colCP05 As Integer, colCP06 As Integer
Dim colGrpY As Integer, colGrpYman As Integer, colGrpN As Integer
Dim colTCT02 As Integer, colTCT03 As Integer, colTCN13 As Integer

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   txtInput.Visible = False '跳離編輯
   If Index = 0 Then '查詢
      If doQuery(True) = False Then
      End If
   Else
      cmdState = Index
      PubShowNextData
   End If
End Sub

Public Sub PubShowNextData()
Dim inX As Integer, inY As Integer
Dim Str01 As String
Dim lngColor As Long

    Me.Enabled = False
    For inX = 1 To MGRD1.Rows - 1
       MGRD1.row = inX
       MGRD1.col = 0
       '更新採用全部檢查
       If (mRole = "M" And Trim(MGRD1.Text) = "V") Or (mRole <> "M" And ((cmdState <> 1 And Trim(MGRD1.Text) = "V") Or cmdState = 1)) Then
           If Trim(MGRD1.Text) = "V" Then
               MGRD1.Text = ""
               MGRD1.col = 0
               MGRD1.CellBackColor = MGRD1.BackColor
               MGRD1.col = cntFixed + 1
               lngColor = MGRD1.CellBackColor
               For inY = 0 To cntFixed
                   MGRD1.col = inY
                   MGRD1.CellBackColor = lngColor
               Next inY
           End If

           '本所案號
           Str01 = Trim(MGRD1.TextMatrix(inX, colCP01)) & "-" & Trim(MGRD1.TextMatrix(inX, colCP02)) & "-" & Trim(MGRD1.TextMatrix(inX, colCP03)) & "-" & Trim(MGRD1.TextMatrix(inX, colCP04))
           If Replace(Str01, "-", "") <> "" Then
               'Move by Lydia 2023/06/26 從Str01上方搬下來
               If cmdState > 1 Then
                  If cmdState = 2 Then
                      If PUB_CheckFormExist("frm100101_M") Then
                          MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
                          GoTo JumpToExit
                      End If
                  End If
                  If cmdState = 4 Then
                      If PUB_CheckFormExist("frm100101_L") Then
                          MsgBox "請先關閉共同查詢〔卷宗區〕畫面！"
                          GoTo JumpToExit
                      End If
                  End If
                  If fnSaveParentForm(Me) = False Then
                      GoTo JumpToExit
                  End If
               End If
               'End --- 'Move by Lydia 2023/06/26 從Str01上方搬下來
                Select Case cmdState
                    Case 1 '確定=>更新；M-主管核判
                        If mRole = "M" Then
                           Me.Hide
                           Call frm090908_1.SetParent(Me, Str01, Trim(MGRD1.TextMatrix(inX, colCp09)), m_UserNo)
                           frm090908_1.Show
                        Else
                           'Added by Lydia 2023/05/08
                           If TxtValidate(inX, "" & MGRD1.TextMatrix(inX, colInNew)) = False Then
                              GoTo JumpToExit
                           End If
                           'end 2023/05/08
                           If "" & MGRD1.TextMatrix(inX, colInNew) <> "" Then
                               If "" & MGRD1.TextMatrix(inX, colInNew) <> "" & MGRD1.TextMatrix(inX, colInOLD) Then
                                  Screen.MousePointer = vbHourglass
                                    'Modified by Lydia 2023/06/14 ＋申請日PA10,非英說案的狀態(外文本的對應英/中說)TCN13
                                    If SaveDatabase(MGRD1.TextMatrix(inX, colCp09), MGRD1.TextMatrix(inX, colTCN23), MGRD1.TextMatrix(inX, colInNew), MGRD1.TextMatrix(inX, colCP01), MGRD1.TextMatrix(inX, colCP02), MGRD1.TextMatrix(inX, colCP03), MGRD1.TextMatrix(inX, colCP04), MGRD1.TextMatrix(inX, colPA10), MGRD1.TextMatrix(inX, colTCN13)) = True Then
                                    End If
                                  Screen.MousePointer = vbDefault
                               End If
                           ElseIf "" & MGRD1.TextMatrix(inX, colInOLD) <> "" Then
                               MsgBox MGRD1.TextMatrix(inX, colCaseNo) & "請輸入認領狀態Y或N ！" & vbCrLf & "原本認領狀態=" & MGRD1.TextMatrix(inX, colInOLD), vbCritical + vbOKOnly, "輸入檢查"
                               If doQuery(False) = False Then '避免其他案件有更新
                               End If
                               GoTo JumpToExit
                           End If
                        End If
                    Case 2 '原始檔
                        Call ChgCaseNo(Replace(Str01, "-", ""), strExc)
                        If PUB_ChkCPExist(strExc, cntEnglish_Vers, , strExc(0), , "D") = True Then 'English_Vers992
                            Screen.MousePointer = vbHourglass
                            frm100101_M.m_strKey = strExc(0)
                            frm100101_M.SetParent Me
                            If frm100101_M.QueryData = True Then
                               frm100101_M.Show
                               Me.Hide
                            End If
                            Screen.MousePointer = vbDefault
                        Else
                           MsgBox strExc(1) & "-" & strExc(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
                        End If
                    Case 3 '基本檔
                         Screen.MousePointer = vbHourglass
                         frm100101_3.Show
                         frm100101_3.Tag = Pub_RplStr(Str01)
                         frm100101_3.StrMenu
                         Screen.MousePointer = vbDefault
                    Case 4 '卷宗區
                        Screen.MousePointer = vbHourglass
                        frm100101_L.m_strKey = Str01
                        frm100101_L.SetParent Me
                        If frm100101_L.QueryData = True Then
                           frm100101_L.Show
                           Me.Hide
                        End If
                        Screen.MousePointer = vbDefault
                End Select
           End If
       End If
    Next inX
    
    '確定後更新畫面
    If cmdState = 1 Then
       'Modified by Lydia 2023/05/11 +True 改用QPGMR發mail
       PUB_SendMailCache True  '避免畫面停留過久,沒有發email
       If doQuery(True) = False Then
       End If
    End If
    
JumpToExit:
    Me.Enabled = True

End Sub

Private Sub Combo1_Click()
   '直接查詢
   If Combo1.Tag <> "" And Combo1.Tag <> Combo1.Text Then
       If doQuery(True) = False Then
       End If
   End If
   Combo1.Tag = Combo1.Text
End Sub

Private Sub Command1_Click()
Dim strR1 As String, intR As Integer
Dim strQ1 As String, intQ As Integer
Dim rsRd As New ADODB.Recordset
Dim rsQD As New ADODB.Recordset
Dim strEDate As String, strETime As String, strNewType As String
Dim strTo As String, strCC As String, strSpecSub As String
Dim bolConn As Boolean
Dim xlsPrintList
Dim wksPrint
Dim strTitle As Variant, strTitleW As Variant
Dim strCont As String

'GoTo JumpToFirst '直接跳每個工作日下午２點通知核判主管
   
   If ChkWorkDay(strSrvDate(1)) = False Then
       Exit Sub  '非工作日不執行
   End If
   
   '認領階段
   'Modified by Lydia 2023/06/14 (5/19 Email):昨會後經David建議，新案認領組別(非急件)逾期通知 :排除最高主管核判TCN24
   'strR1 = "select tct01,tct04,tcn21,tcn22,tcn23,cp01,cp02,cp03,cp04,pa10 from trackingcasename,transcasetitle,caseprogress,patent " & _
               "where tcn05=tct01(+) and tct04 is null and tct01 is not null and tcn05=cp09(+) and cp159=0 and cp05>=20230501 " & _
               "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and tcn23 in ('0','1','2','3') " & _
               "and (tcn21<to_char(sysdate,'yyyymmdd') or (tcn21=to_char(sysdate,'yyyymmdd') and tcn22<=substr(lpad(to_char(sysdate,'hh24miss'),6,'0'),1,4))) "
   strR1 = "select tct01,tct04,tcn21,tcn22,cp66,cp67,tcn23,cp01,cp02,cp03,cp04,pa10,tcn25,tcn13,pa75,nvl(fa05,nvl(fa04,fa06)) pa75n,pa26,nvl(cu05,nvl(cu04,cu06)) pa26n " & _
               "From trackingcasename, transcasetitle, caseprogress, patent, fagent, customer " & _
               "where tcn05=tct01(+) and tct04 is null and tct01 is not null and tcn05=cp09(+) and cp159=0 and cp05>=to_char(sysdate,'yyyymmdd')-10000 and cp10 in (" & GetAddStr(FcpAddTct) & ")" & _
               "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and tcn23 in ('0','1','2','3','4','5') and nvl(tcn24,'N') <> 'Y' and nvl(tcn25,0) < 2 " & _
               "and (tcn21<to_char(sysdate,'yyyymmdd') or (tcn21=to_char(sysdate,'yyyymmdd') and tcn22<=substr(lpad(to_char(sysdate,'hh24miss'),6,'0'),1,4))) " & _
               "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
   intR = 1
   Set rsRd = ClsLawReadRstMsg(intR, strR1)
   If intR = 1 Then
      rsRd.MoveFirst
      Do While Not rsRd.EOF
         'Modified by Lydia 2023/06/14 Email主旨開頭改成模組
         strSpecSub = PUB_GetTCNmTitle(rsRd.Fields("cp01"), rsRd.Fields("cp02"), rsRd.Fields("cp03"), rsRd.Fields("cp04"), "" & rsRd.Fields("pa10"), "" & rsRd.Fields("tcn13"), "SPEC")
         strCont = "代理人：" & IIf("" & rsRd.Fields("PA75") <> "", rsRd.Fields("PA75") & " " & rsRd.Fields("PA75N"), "（空白）") & vbCrLf & _
                       "申請人：" & IIf("" & rsRd.Fields("PA26") <> "", rsRd.Fields("PA26") & " " & rsRd.Fields("PA26N"), "（空白）") & vbCrLf
         strNewType = ""
         Select Case "" & rsRd.Fields("tcn23")
             'Modifed by Lydia 2023/06/14
             'Case "0", "2", "3" '急件認領+職代認領(最後)+協調認領
             Case "0"
                 '逾期檢查: 是否有人認領
                 'Modified by Lydia 2023/06/14
                 'If "" & rsRd.Fields("tcn23") = "2" Or "" & rsRd.Fields("tcn23") = "3" Then
                 '    strQ1 = "select st01,st02,tfa05,st16 from transfeeassign,staff where tfa01='" & rsRd.Fields("tct01") & "' and tfa04=st01(+) " & _
                                  "and tfa05='Y' and tfa09=" & CNULL(IIf("" & rsRd.Fields("tcn23") = "2", "1", "2"))
                       strQ1 = "select st01,st02,tfa05,st16 from transfeeassign,staff where tfa01='" & rsRd.Fields("tct01") & "' and tfa04=st01(+) " & _
                                  "and tfa05='Y' and tfa09=" & CNULL(rsRd.Fields("tcn23"))
                 'end 2023/06/14
                     intQ = 1
                     Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
                     If intQ = 1 Then
                        If rsQD.RecordCount = 1 Then
                           bolConn = True
                           cnnConnection.BeginTrans
                             strSql = "Update TrackingCaseName Set TCN20='" & rsQD.Fields("st16") & "' Where TCN05='" & rsRd.Fields("tct01") & "' "
                             cnnConnection.Execute strSql
                             If PUB_UpdateTCNstate("2", rsRd.Fields("cp01") & rsRd.Fields("cp02") & rsRd.Fields("cp03") & rsRd.Fields("cp04")) = False Then
                                GoTo ErrHandle
                             End If
                           cnnConnection.CommitTrans
                           strNewType = "U"
                        End If
                     End If
                 'End If 'Mark by Lydia 2023/06/14
                 '沒人認領或超過1組=>最高主管進行核判TCN24
                 If strNewType = "" Then
                     'strNewType = "4" 'Mark by Lydia 2023/06/14
                     '核判期限至隔日下班前
                     strEDate = CompWorkDay(2, rsRd.Fields("tcn21"))
                     strETime = "1700"
                     'Added by Lydia 2023/06/14
                     strSql = "Update TrackingCaseName Set TCN24='Y', TCN21=" & strEDate & ", TCN22=" & strETime & " Where TCN05='" & rsRd.Fields("tct01") & "' "
                     cnnConnection.Execute strSql
                     'end 2023/06/14
                     'Email通知
                     If PUB_GetTCNEmail(rsRd.Fields("cp01"), rsRd.Fields("cp02"), rsRd.Fields("cp03"), rsRd.Fields("cp04"), IIf(rsRd.Fields("tcn23") = "0", "0", "1")) = True Then
                     End If
                 End If
             Case "1"  '主管認領2H=>職代認領+1H
                 strNewType = "2"
                 Call PUB_CompWorkTime("" & rsRd.Fields("tcn22"), 60, strETime, "" & rsRd.Fields("tcn21"), strEDate)
                 'Email通知參考PUB_UpdateTCNstate
                 strTo = PUB_GetEngGrpMan(strCC)
                 strExc(1) = Replace(strSpecSub, "SPEC", "") & "，請協助確認組別，謝謝！"
                 
                 strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                                " values( '" & strUserNum & "','" & strCC & "',to_char(sysdate,'yyyymmdd')" & _
                                ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & ChgSQL(strCont) & "')"
                 cnnConnection.Execute strSql
             'Added by Lydia 2023/06/14
             Case "2", "3", "4", "5" '已到職代認領(2)或協調(3)、非英說(4)認領階段=> 逾期通知: 1=逾3小時,2=逾1天
                 strExc(3) = ""
                 If Val("" & rsRd.Fields("tcn25")) = 0 Then '以最後認領期限計算:逾3小時
                     Call PUB_CompWorkTime("" & rsRd.Fields("tcn22"), 180, strETime, "" & rsRd.Fields("tcn21"), strEDate)
                     If strSrvDate(1) & Format(Now, "hhmm") >= strEDate & Left(strETime, 4) Then
                        strExc(3) = "1"
                     End If
                 Else '以建檔日期計算:逾1天=> (6/1認領日期+時間逾1天)
                     strExc(2) = CompWorkDay(2, "" & rsRd.Fields("tcn21"))
                     If strSrvDate(1) & Format(Now, "hhmm") >= strExc(2) & Format(rsRd.Fields("tcn22"), "0000") Then
                        strExc(3) = "2"
                     End If
                 End If
                 If strExc(3) <> "" Then
                     strSql = "Update TrackingCaseName Set TCN25='" & strExc(3) & "' Where TCN05='" & rsRd.Fields("TCT01") & "' "
                     cnnConnection.Execute strSql
                     strTo = PUB_GetEngGrpMan(strCC)
                     '同時CC職代+程序
                     strExc(2) = PUB_GetFCPHandler("" & rsRd.Fields("CP01"), "" & rsRd.Fields("CP02"), "" & rsRd.Fields("CP03"), "" & rsRd.Fields("CP04"))
                     strExc(1) = Replace(strSpecSub, "SPEC", IIf(strExc(3) = "1", "-已逾3小時通知", "-已逾24小時通知")) & "，請協助確認組別，謝謝！"
                     strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                  " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                                  ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & ChgSQL(strCont) & "','" & strCC & ";" & strExc(2) & "')"
                     cnnConnection.Execute strSql
                 End If
             'end by Lydia 2023/06/14
         End Select
         If strNewType <> "" And strEDate <> "" And strETime <> "" Then
             strSql = "Update TrackingCaseName Set TCN23='" & strNewType & "', TCN21=" & CNULL(strEDate, True) & ", TCN22=" & CNULL(Left(strETime, 4), True) & " Where TCN05='" & rsRd.Fields("TCT01") & "' "
             cnnConnection.Execute strSql
         End If
         rsRd.MoveNext
      Loop
   End If
   
   '每個工作日下午２點(14:00)若有前日未核判之新案(非提申急件)，由系統寄email提醒通知國外部最高主管進行核判。
   'Modified by Lydia 2023/05/17 因為放在最後面，1400無法整點執行，改用記錄判斷
   'If Val(Format(Now, "hhmm")) = 1400 Then
   If Val(Format(Now, "hhmm")) >= 1400 And Val(Format(Now, "hhmm")) <= 1430 Then
JumpToFirst:
       strR1 = "select * from addressa4list where aal01='TFA' and aal02=to_char(sysdate,'yyyymmdd') "
       intR = 1
       Set rsRd = ClsLawReadRstMsg(intR, strR1)
       If intR = 0 Then
         strSql = "Delete from addressa4list where aal01='TFA' and aal02<to_char(sysdate,'yyyymmdd') "
         cnnConnection.Execute strSql
         strSql = "Insert into addressa4list (aal01,aal02,aal03,aal04) values ('TFA',to_char(sysdate,'yyyymmdd'),'1','QPGMR') "
         cnnConnection.Execute strSql
   'end 2023/05/17
         strSpecSub = ""
         strTitle = Split("承辦人員,收文日,本所案號,代　理　人,申　請　人,本所期限,法定期限", ",")
         strTitleW = Split("13,10,12,30,30,10,10", ",")
         
         'Modified by Lydia 2023/06/14 最高主管只管急件逾期或協調不過; and (tcn23='4' or pa10 is null) => and nvl(tcn24,'N')='Y'
         strR1 = "select tcn20,tcn21,tcn22,tcn23, tct01,tct02,tct03,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp10,cp13,st02 as cp13n,st16 " & _
                    ",pa75,nvl(fa04,nvl(fa05,fa06)) pa75n,pa26,nvl(cu04,nvl(cu05,cu06)) pa26n " & _
                    "From transcasetitle, caseprogress, staff, trackingcasename, patent, fagent, customer " & _
                    "where tct04 is null and cp159=0 and tct01=cp09(+) and cp13=st01(+) and cp66 < to_char(sysdate,'yyyymmdd') and cp05>=to_char(sysdate,'yyyymmdd')-10000 and cp05<" & strSrvDate(1) & _
                    " and tct01=tcn05(+) and nvl(tcn16,'N')<>'Y' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                    " and nvl(tcn24,'N')='Y' and nvl(tcn23,'0') <> '9' and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                    " order by cp05"
         intR = 1
         Set rsRd = ClsLawReadRstMsg(intR, strR1)
         If intR = 1 Then
            rsRd.MoveFirst
            Do While Not rsRd.EOF
               'If "" & rsRd.Fields("tcn23") = "4" Or Not (("" & rsRd.Fields("cp06") <> "" And strSrvDate(1) >= "" & rsRd.Fields("cp06")) Or ("" & rsRd.Fields("cp07") <> "" And strSrvDate(1) >= "" & rsRd.Fields("cp07"))) Then 'Mark by Lydia 2023/06/14
                   If strSpecSub = "" Then
                       strSpecSub = App.path & "\外專新案尚未認領清單" & MsgText(43)
                       If Dir(strSpecSub) <> "" Then
                           Kill strSpecSub
                       End If
                       intQ = 1
                       Set xlsPrintList = CreateObject("Excel.Application")
                       xlsPrintList.SheetsInNewWorkbook = 1
                       xlsPrintList.Workbooks.add
                       Set wksPrint = xlsPrintList.Worksheets(1)
                       wksPrint.Activate
                       'xlsPrintList.Visible = True
                       For intQ = 0 To UBound(strTitle)
                          wksPrint.Range(Chr(65 + intQ) & "1").Value = Trim(strTitle(intQ))
                          wksPrint.Range(Chr(65 + intQ) & ":" & Chr(65 + intQ)).ColumnWidth = Val(strTitleW(intQ))
                       Next intQ
                       wksPrint.Range("A1:" & Chr(65 + UBound(strTitle)) & "1").Font.Bold = True
                       intQ = 2
                   End If
                   wksPrint.Range("A" & intQ).Value = rsRd.Fields("CP13") & " " & rsRd.Fields("cp13n")
                   wksPrint.Range("B" & intQ).Value = ChangeWStringToTDateString(rsRd.Fields("CP05"))
                   wksPrint.Range("C" & intQ).Value = rsRd.Fields("CP01") & "-" & rsRd.Fields("CP02") & IIf(rsRd.Fields("CP03") & rsRd.Fields("CP04") <> "000", "-" & rsRd.Fields("CP03") & "-" & rsRd.Fields("CP04"), "")
                   wksPrint.Range("D" & intQ).Value = rsRd.Fields("PA75") & " " & rsRd.Fields("PA75n")
                   wksPrint.Range("E" & intQ).Value = rsRd.Fields("PA26") & " " & rsRd.Fields("PA26n")
                   wksPrint.Range("F" & intQ).Value = ChangeWStringToTDateString(rsRd.Fields("CP06"))
                   wksPrint.Range("G" & intQ).Value = ChangeWStringToTDateString(rsRd.Fields("CP07"))
               'End If 'Mark by Lydia 2023/06/14
               intQ = intQ + 1
               rsRd.MoveNext
            Loop
            If strSpecSub <> "" Then
               If Val(xlsPrintList.Version) < 12 Then
                   xlsPrintList.Workbooks(1).SaveAs FileName:=strSpecSub, FileFormat:=-4143
               Else
                   xlsPrintList.Workbooks(1).SaveAs FileName:=strSpecSub, FileFormat:=56
               End If
               xlsPrintList.Workbooks.Close
               xlsPrintList.Quit
               Set xlsPrintList = Nothing
               Set wksPrint = Nothing
               strQ1 = Pub_GetSpecMan("外專新案命名核判主管")
               If strQ1 <> "" Then
                   PUB_SendMail strUserNum, strQ1, "", ChangeTStringToTDateString(strSrvDate(2)) & "外專新案尚未認領清單", "請參考附件", , strSpecSub, , , , , , , , , , , , False
               End If
            End If
         End If
       End If 'Added by Lydia 2023/05/17 改用記錄判斷
   End If
   
   Set rsRd = Nothing
   Set rsQD = Nothing
   PUB_SendMailCache
   MsgBox "End"
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
       If bolConn = True Then
          cnnConnection.RollbackTrans
       End If
       MsgBox Err.Description, vbCritical, "上線後拿掉MsgBox"
   End If
End Sub

Private Sub Form_Load()
Dim strTmp As String

    MoveFormToCenter Me
    txtInput.Visible = False
    
    '參考PUB_GetFCPGrpName
    mGrpName(1) = "電子"
    mGrpName(2) = "化學"
    mGrpName(3) = "日文"
    mGrpName(4) = "機械"
    
    Command1.Visible = False
    If mRole = "M" Then '主管核判
       Me.Caption = "外專新案認領區-主管核判"
       Combo1.AddItem strUserNum & " " & strUserName
       Combo1.ListIndex = 0
       Combo1.Visible = False
       lblSname.Visible = False
       cmdOK(1).Caption = "明細(&E)"
       Label3.Visible = False
       Combo2.Visible = False
       If Pub_StrUserSt03 = "M51" Then
          Command1.Visible = True
       End If
    Else
      Label3.Visible = True
      Combo2.Visible = True
      mRole = "A"
      If Pub_StrUserSt03 = "M51" Then
JumpToInput:
         strTmp = UCase(InputBox("請輸入欲操作的工程師組別(1~4)或輸入5全部？", , "1"))
         If Val(strTmp) < 1 And Val(strTmp) > 5 Then
             GoTo JumpToInput
         End If
      End If
      Combo1.Clear
      If strTmp <> "" Then
         If Val(strTmp) < 5 Then
            Combo1.AddItem Pub_GetFCPGrpMan(strTmp) & " " & GetStaffName(Pub_GetFCPGrpMan(strTmp))
         Else
            For intI = LBound(mGrpName) To UBound(mGrpName)
                If intI <> 3 Then
                   Combo1.AddItem Pub_GetFCPGrpMan(intI) & " " & GetStaffName(Pub_GetFCPGrpMan(intI))
                End If
            Next intI
         End If
      Else
         Combo1.AddItem strUserNum & " " & strUserName
      End If
      Combo1.ListIndex = 0
    End If
    
    '(案件)顏色說明
    Combo2.Clear
    Combo2.AddItem "紅色：急件認領"
    Combo2.AddItem "黃色：已提申"
    Combo2.AddItem "綠色：協調認領"
    Combo2.AddItem "橘色：非英說案協調認領"
    Combo2.ListIndex = 0
    
    If doQuery(False) = False Then
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache True 'Modified by Lydia 2023/05/11 +True 改用QPGMR發mail
   Set frm090908 = Nothing
End Sub

Public Function doQuery(ByVal bolMsg As Boolean) As Boolean
Dim strQuery As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
    
   SetGrd True
   m_UserNo = Trim(Left(Combo1.Text, 6))
   m_UserSt16 = PUB_GetStaffST16(m_UserNo)
   If m_UserSt16 = "" Then m_UserSt16 = "1"  '預設
   txtInput.Text = "": txtInput.Tag = ""
   
   If m_UserNo <> "" Then
      '2=協調認領; 檢查前一階段TFA09=1 'AND TCN23 IN ('0','1','2','3')
      'Modified by Lydia 2023/06/14 排除暫不認領AND NVL(TCN21,'0') <> '99999999'
      'strQuery = "SELECT '' AS V,'' AS 認領狀態,DECODE(TCN21,NULL,'',SUBSTR(SQLDATET(TCN21),1,9)||' '||SUBSTR(SQLTIME6(TCN22||'00'),1,5)) AS 認領期限 " & _
                  ",'' AS 認領組,'' AS 認領人員,'' AS 不認領組,SUBSTR(SQLDATET(CP05),1,9) AS 收文日,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) AS 本所案號 " & _
                  ",DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,DECODE(TCT02,NULL,'',SUBSTR(SQLDATET(TCT02),1,9)||' '||SUBSTR(SQLTIME6(TCT03||'00'),1,5)) AS 譯畢期限 " & _
                  ",TCN17 AS 相似舊案,TCN20 AS 急件分組,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,PA10,TCN23,'' AS 認領狀態OLD,TCT02,TCT03 " & _
                  "From TRANSCASETITLE, CASEPROGRESS, CASEPROPERTYMAP, PATENT, TRACKINGCASENAME, CUSTOMER " & _
                  "WHERE TCT04 IS NULL AND TCT01=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL " & _
                  "AND PA108 IS NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND TCT01=TCN05(+) AND NVL(TCN16,'N')='N' AND TCN23 IN ('0','1','2') " & _
                  "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) "
      strQuery = "SELECT '' AS V,'' AS 認領狀態,DECODE(TCN21,NULL,'',SUBSTR(SQLDATET(TCN21),1,9)||' '||SUBSTR(SQLTIME6(TCN22||'00'),1,5)) AS 認領期限 " & _
                  ",'' AS 認領組,'' AS 認領人員,'' AS 不認領組,SUBSTR(SQLDATET(CP05),1,9) AS 收文日,DECODE(TCN13,0,'',NULL,'','▲')||CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) AS 本所案號 " & _
                  ",DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,DECODE(TCT02,NULL,'',SUBSTR(SQLDATET(TCT02),1,9)||' '||SUBSTR(SQLTIME6(TCT03||'00'),1,5)) AS 譯畢期限 " & _
                  ",TCN17 AS 相似舊案,TCN20 AS 急件分組,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,PA10,TCN23,'' AS 認領狀態OLD,TCT02,TCT03,TCN13 " & _
                  "From TRANSCASETITLE, CASEPROGRESS, CASEPROPERTYMAP, PATENT, TRACKINGCASENAME, CUSTOMER " & _
                  "WHERE TCT04 IS NULL AND TCT01=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL " & _
                  "AND PA108 IS NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND TCT01=TCN05(+) AND NVL(TCN16,'N')='N' AND NVL(TCN21,'0') <> '99999999' AND TCN23 IN ('0','1','2') " & _
                  "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) "
      If mRole = "M" Then
         'Modified by Lydia 2023/06/14 最高主管進行核判 AND TCN23='4'=> AND TCN24='Y'
         strSql = Replace(strQuery, " AND TCN23 IN ('0','1','2')", "") & " AND TCN24='Y' "
      Else
         strQuery = strQuery & " AND TCN24 IS NULL "
         '2=協調認領:抓一般認領階段=Y的組別
         'Modified by Lydia 2023/06/14 拿掉人員編號
         'strSql = strQuery & " UNION " & Replace(strQuery, " AND TCN23 IN ('0','1','2')", " AND TCN23='3' AND CP09 IN (SELECT TFA01 FROM TRANSFEEASSIGN,STAFF WHERE TFA04=ST01(+) AND TFA09='1' AND TFA05='Y' AND ST16='" & m_UserSt16 & "' AND TFA01=CP09 AND TFA04='" & m_UserNo & "')")
         strSql = strQuery & " UNION " & Replace(strQuery, " AND TCN23 IN ('0','1','2')", " AND TCN23='3' AND CP09 IN (SELECT TFA01 FROM TRANSFEEASSIGN,STAFF WHERE TFA04=ST01(+) AND TFA09='1' AND TFA05='Y' AND ST16='" & m_UserSt16 & "' AND TFA01=CP09 )")
         'Added by Lydia 2023/06/14 非英說案協調認領=2次認領(已預設新增空白記錄TFA09=4)
         strSql = strSql & " UNION " & Replace(strQuery, " AND TCN23 IN ('0','1','2')", " AND TCN23='4' AND CP09 IN (SELECT TFA01 FROM TRANSFEEASSIGN,STAFF WHERE TFA04=ST01(+) AND TFA09='4' AND ST16='" & m_UserSt16 & "' AND TFA01=CP09)")
         'Added by Lydia 2023/06/14 非英說案2次認領,有2人以上=>非英說案協調認領=5
         'Mark by Lydia 2023/06/14 (保留)
         'strSql = strSql & " UNION " & Replace(strQuery, " AND TCN23 IN ('0','1','2')", " AND TCN23='5' AND CP09 IN (SELECT TFA01 FROM TRANSFEEASSIGN,STAFF WHERE TFA04=ST01(+) AND TFA09='4' AND TFA05='Y' AND ST16='" & m_UserSt16 & "' AND TFA01=CP09 )")
      End If
      If bolMsg = True Then
         intQ = 0
      Else
         intQ = 1
      End If
      Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
      MGRD1.FixedCols = 0
      
      If intQ = 1 Then
         doQuery = True
         Set MGRD1.Recordset = rsQuery
         SetGrd False
         MGRD1.FixedCols = cntFixed
      Else
         doQuery = False
      End If
   End If
   
   Set rsQuery = Nothing
End Function

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, intR As Integer
   Dim pTime As String
   Dim lngColor As Long
   
   'V|認領狀態|認領期限|認領組|認領人員|不認領組|收文日|本所案號|案件性質|譯畢期限|急件分組|CP01|CP02|CP03|CP04|CP05|CP06|CP07|CP09|CP10|CP13PA10|TCN23|認領狀態OLD
   pTime = Mid(Format(ServerTime, "000000"), 1, 4)
   'Modified by Lydia 2023/06/14 +TCN13
   arrGridHeadText = Array("V", "狀態", "認領期限", "認領組", "認領人員", "不認領組", "收文日", "本所案號", "案件性質", "譯畢期限", "相似舊案", "急件分組", "CP01", "CP02", "CP03", "CP04", "CP05", "CP06", "CP07", "CP09", "CP10", "CP13", "PA10", "TCN23", "認領狀態OLD", "TCT02", "TCT03", "TCN13")
   If mRole = "M" Then '主管核判
       arrGridHeadWidth = Array(300, 0, 1240, 1000, 0, 1000, 840, 1100, 840, 1000, 1000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   Else
       arrGridHeadWidth = Array(300, 500, 1240, 1000, 1100, 1000, 840, 1100, 840, 1000, 1000, 1000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   End If
   MGRD1.Visible = False
   MGRD1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
        MGRD1.Clear
        MGRD1.Rows = 2
   End If

   For iRow = 0 To MGRD1.Cols - 1
      MGRD1.row = 0
      MGRD1.col = iRow
      MGRD1.Text = arrGridHeadText(iRow)
      MGRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGRD1.CellAlignment = flexAlignCenterCenter
   Next
   If colCP01 = 0 Then
      colInNew = PUB_MGridGetId("狀態", MGRD1)
      colInOLD = PUB_MGridGetId("認領狀態OLD", MGRD1)
      colTCN23 = PUB_MGridGetId("TCN23", MGRD1)      '認領期限狀態
      colTCN20 = PUB_MGridGetId("急件分組", MGRD1)      '急件分組：急件提申發文時取消案件之工程師組別，重新進入認領流程(只將TCT04=null, TCN20保持急件分組)
      colGrpY = PUB_MGridGetId("認領組", MGRD1)
      colGrpYman = PUB_MGridGetId("認領人員", MGRD1)
      colGrpN = PUB_MGridGetId("不認領組", MGRD1)
      colTCN13 = PUB_MGridGetId("TCN13", MGRD1)       'Added by Lydia 2023/06/14 非英說案的狀態(外文本的對應英/中說)
      colCaseNo = PUB_MGridGetId("本所案號", MGRD1)
      colCP01 = PUB_MGridGetId("CP01", MGRD1)
      colCP02 = PUB_MGridGetId("CP02", MGRD1)
      colCP03 = PUB_MGridGetId("CP03", MGRD1)
      colCP04 = PUB_MGridGetId("CP04", MGRD1)
      colCP05 = PUB_MGridGetId("CP05", MGRD1)
      colCP06 = PUB_MGridGetId("CP06", MGRD1)
      colCp09 = PUB_MGridGetId("CP09", MGRD1)
      colPA10 = PUB_MGridGetId("PA10", MGRD1)  '申請日
      colTCT02 = PUB_MGridGetId("TCT02", MGRD1)
      colTCT03 = PUB_MGridGetId("TCT03", MGRD1)
   End If
   
   For intR = 1 To MGRD1.Rows - 1
      MGRD1.row = intR
      '取得欄位內容
      lngColor = &H80000005
      If "" & MGRD1.TextMatrix(intR, colCp09) <> "" Then
         strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(3) = ""
         If GetStateTFA(MGRD1.TextMatrix(intR, colCp09), MGRD1.TextMatrix(intR, colTCN23), strExc(1), strExc(2), strExc(3), strExc(4)) = True Then
             '認領狀態
             MGRD1.TextMatrix(intR, colInNew) = strExc(1)
             MGRD1.TextMatrix(intR, colInOLD) = strExc(1)
             '認領組,人員
             MGRD1.TextMatrix(intR, colGrpY) = strExc(2)
             MGRD1.TextMatrix(intR, colGrpYman) = strExc(3)
             '不認領組
             MGRD1.TextMatrix(intR, colGrpN) = strExc(4)
         End If
         '急件分組
         If "" & MGRD1.TextMatrix(intR, colTCN20) <> "" Then
             MGRD1.TextMatrix(intR, colTCN20) = mGrpName(Val(MGRD1.TextMatrix(intR, colTCN20)))
         End If
         If mRole = "M" Then '核判主管=最高主管
            '急件提申發文時取消案件之工程師組別，重新進入認領流程
            If "" & MGRD1.TextMatrix(intR, colPA10) <> "" Then
                 lngColor = vbYellow
            '有譯畢期限並且系統時間距離期限小於2小時並且尚未有主管認領，則那條記錄顯示為紅色
            '當日以前收文最高主管尚未核判而提申期限為當日者，也視做提申急件
            ElseIf (strExc(2) = "" And Val("" & MGRD1.TextMatrix(intR, colTCT02)) > 0 And Val("" & MGRD1.TextMatrix(intR, colTCT02) & MGRD1.TextMatrix(intR, colTCT02 + 1)) - Val(strSrvDate(1) & pTime) < 200) Or _
                    ("" & MGRD1.TextMatrix(intR, colCP05) < strSrvDate(1) And "" & MGRD1.TextMatrix(intR, colCP06) <> "" And "" & MGRD1.TextMatrix(intR, colCP06) <= strSrvDate(1)) Then
                 lngColor = vbRed
            End If
         Else '工程師主管
            'Added by Lydia 2023/06/14 非英說案協調認領，整列資料顯示為橘色
            If "" & MGRD1.TextMatrix(intR, colTCN23) = "4" Then
                 lngColor = &H80FF&
            'end 2023/06/14
            '協調認領，整列資料顯示為綠色；排除M-主管核判
            ElseIf "" & MGRD1.TextMatrix(intR, colTCN23) = "3" Then
                 lngColor = &H80FF80
            '急件提申發文時取消案件之工程師組別，重新進入認領流程(只將TCT04=null, TCN20保持急件分組)；在認領畫面以黃色標示已提申
            ElseIf "" & MGRD1.TextMatrix(intR, colPA10) <> "" Then
                 lngColor = vbYellow
            '提申急件沒有預設組別處於可認領狀態，整列資料顯示為紅色
            ElseIf "" & MGRD1.TextMatrix(intR, colTCN23) = "0" Then
                 lngColor = vbRed
            End If
         End If
      End If
      For iRow = 0 To MGRD1.Cols - 1
         MGRD1.col = iRow
         MGRD1.CellBackColor = lngColor
         '置中
         If iRow = colInNew Then
            MGRD1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intR
   
   MGRD1.Visible = True
End Sub

Private Sub MGRD1_Click()
Dim intRow As Integer, intCol As Integer
Dim lngColor As Long
   With MGRD1
       If .MouseRow > 0 Then
          intRow = .MouseRow
          intCol = .MouseCol
          .row = intRow
          .col = cntFixed + 1 '還原底色
          lngColor = .CellBackColor
          '----單選
          GridClick MGRD1, intLastRow, 0, 0, cntFixed, "V", lngColor
          intLastRow = intRow
          .col = intCol
          If "" & MGRD1.TextMatrix(intRow, 0) = "V" And intCol = colInNew Then
              SetBox
          Else
              txtInput.Visible = False
          End If
       End If
   End With
End Sub

'讀取目前認領狀態
'Modified by Lydia 2023/06/14 +目前非英說案的狀態(外文本的對應英/中說)
Private Function GetStateTFA(ByVal pCP09 As String, ByVal pState As String, ByRef nowACK As String, Optional ByRef nGrpYes As String, Optional ByRef nGrpYesMan As String, Optional ByRef nGrpNO As String, Optional ByRef nTCN13 As String) As Boolean
'nowAck : 工程師所屬組別之認領狀態
'nGrpYes,nGrpYesMan：目前認領組, 認領人員
'nGrpNO：不認領組
Dim strAll As String

   nowACK = "": nGrpYes = "": nGrpYesMan = "": nGrpNO = ""
   If m_UserSt16 <> "" And pState <> "" Then
       If mRole = "M" Then '主管核判
         strP1 = "select st01,st02,tfa05,st16,'' as nstate from transfeeassign,staff " & _
                    "where tfa01='" & pCP09 & "' and tfa04=st01(+) order by tfa09 desc, st16 "
       Else
         '新案認領狀態: 0=急件認領, 1=認領, 2=協調認領
         If pState = "1" Or pState = "2" Then  '認領階段:1主管期限+2通知職代認領
            strP1 = "and tfa09='1' "
         ElseIf pState = "3" Then
            strP1 = "and tfa09='2' "
         Else
            strP1 = "and tfa09='" & pState & "' "
         End If
         strP1 = "select st01,st02,tfa05,st16,decode(st16,'" & m_UserSt16 & "','Y','') as nstate from transfeeassign,staff " & _
                    "where tfa01='" & pCP09 & "' and tfa04=st01(+) " & strP1
         strP1 = strP1 & "order by tfa05,st16,st01"
       End If
       intP = 1
       Set rsAD = ClsLawReadRstMsg(intP, strP1)
       If intP = 1 Then
           rsAD.MoveFirst
           Do While Not rsAD.EOF
               If InStr("," & strAll, "," & rsAD.Fields("st16")) = 0 Then
                  If "" & rsAD.Fields("tfa05") = "Y" And InStr(nGrpYes & ",", mGrpName(Val(rsAD.Fields("st16")))) = 0 Then
                      nGrpYes = nGrpYes & "," & mGrpName(Val(rsAD.Fields("st16")))
                      nGrpYesMan = nGrpYesMan & "," & rsAD.Fields("st02")
                  ElseIf "" & rsAD.Fields("tfa05") = "N" And InStr(nGrpNO & ",", mGrpName(Val(rsAD.Fields("st16")))) = 0 Then
                      nGrpNO = nGrpNO & "," & mGrpName(Val(rsAD.Fields("st16")))
                  End If
                  '工程師所屬組別之認領狀態
                  If "" & rsAD.Fields("nstate") = "Y" Then
                      nowACK = "" & rsAD.Fields("tfa05")
                  End If
                  strAll = strAll & "," & rsAD.Fields("st16")
               End If
               rsAD.MoveNext
           Loop
           If nGrpYes <> "" Then
              nGrpYes = Mid(nGrpYes, 2)
              nGrpYesMan = Mid(nGrpYesMan, 2)
           End If
           If nGrpNO <> "" Then
              nGrpNO = Mid(nGrpNO, 2)
           End If
           If nowACK & nGrpYes & nGrpNO <> "" Then
               GetStateTFA = True
           End If
       End If
       'Added by Lydia 2023/06/14 目前非英說案的狀態(外文本的對應英/中說)
       strP1 = "Select TCN13 from TrackingCaseName Where TCN05='" & pCP09 & "' "
       intP = 1
       Set rsAD = ClsLawReadRstMsg(intP, strP1)
       If intP = 1 Then
          nTCN13 = "" & rsAD.Fields("TCN13")
       End If
       'end 2023/06/14
       Set rsAD = Nothing
   End If
End Function

Private Sub SetBox()
Dim lngLeft As Long, lngTop As Long, ii As Integer
'參考Promoter\frm090220
   With MGRD1
      If .row > 0 And .col = colInNew Then
         If .TextMatrix(.row, colCp09) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            nRow = .row: nCol = .col
            txtInput.Visible = True
            txtInput.SetFocus
            TextInverse txtInput
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
         End If
      End If
   End With
End Sub

Private Sub MGRD1_DblClick()
   If mRole = "M" And cmdOK(1).Enabled = True Then
       Call cmdok_Click(1)
   End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)

   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   Else
      'Added by Lydia 2023/05/08
      If TxtValidate(nRow, "" & Chr(UpperCase(KeyAscii))) = False Then
          KeyAscii = 0
          Beep
          Exit Sub
      End If
      'end 2023/05/08
      If KeyAscii = vbKeyReturn Then
         MGRD1.TextMatrix(nRow, nCol) = txtInput.Text
         GoNext
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   End If
End Sub

Private Sub GoNext()
   With MGRD1
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      SetBox
   End With
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
   MGRD1.TextMatrix(nRow, nCol) = txtInput.Text
End Sub

Private Sub txtInput_GotFocus()
   TextInverse txtInput
End Sub

'Modified by Lydia 2023/06/14 ＋申請日pPA10, 非英說案的狀態(外文本的對應英/中說)pTCN13
Private Function SaveDatabase(ByVal pCP09 As String, ByVal pState As String, ByVal nowACK As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pPA10 As String, ByVal pTCN13 As String) As Boolean
'nAck : 工程師所屬組別之認領狀態
'nGrpYes,nGrpYesMan：目前認領組, 認領人員
'nGrpNO：不認領組
Dim intCnt As Integer
Dim nGrpACK As String, nGrpYesMan As String, strCon As String, oldACK As String
Dim m_Grp As String, m_GrpMan As String
Dim strCaseNo As String, strAttFList As String
Dim bolUpd As Boolean
Dim strSpecSub As String 'Added by Lydia 2023/06/14
Dim intMaxCnt As Integer, intAgree 'Added by Lydia 2023/06/14 需要認領、目前認領=Y的組別數

   If m_UserSt16 <> "" And pState <> "" Then
       '新案認領狀態: 0=急件認領, 1=認領, 2=協調認領
       If pState = "1" Or pState = "2" Then  '認領階段:1主管期限+2通知職代認領
          strCon = "1"
          'Added by Lydia 2023/06/14
          'Modified by Lydia 2024/02/27 外專機械設計組人員異動調整程式：新案認領組別，請取消機械設計組，只留電子電機組及化學組
          'intMaxCnt = 3
          intMaxCnt = FCPforEngNum
          intAgree = GetTFAcnt(pCP09, strCon, "Y")
          'end 2023/06/14
       ElseIf pState = "3" Then
          strCon = "2"
          'Added by Lydia 2023/06/14
          intMaxCnt = GetTFAcnt(pCP09, "1", "Y")
          intAgree = GetTFAcnt(pCP09, "2", "Y")
          'end 2023/06/14
       'Added by Lydia 2023/06/14 非英說案協調認領
       'Mark by Lydia 2023/06/14 (保留)
       'ElseIf pState = "5" Then
       '   strCon = "5"
       '   intMaxCnt = GetTFAcnt(pCP09, "4", "Y")
       '   intAgree = GetTFAcnt(pCP09, "5", "Y")
       'end 2023/06/14
       Else
          strCon = pState
          'Added by Lydia 2023/06/14
          intMaxCnt = GetTFAcnt(pCP09, strCon)
          intAgree = GetTFAcnt(pCP09, strCon, "Y")
          'end 2023/06/14
       End If
       If pTCN13 = "" Then pTCN13 = "0" 'Added by Lydia 2023/06/14

       'Modified by Lydia 2023/06/14 Email主旨開頭改成模組
       strSpecSub = PUB_GetTCNmTitle(pCP01, pCP02, pCP03, pCP04, pPA10, pTCN13, "")
       '目前認領的狀況
       strExc(1) = "select st01,st02,tfa05,st16,decode(st16,'" & m_UserSt16 & "','Y','') as nstate from transfeeassign,staff " & _
                  "where tfa01='" & pCP09 & "' and tfa04=st01(+) and tfa09=" & CNULL(strCon)
       intP = 1
       strExc(2) = ""
       Set rsAD = ClsLawReadRstMsg(intP, strExc(1))
       If intP = 1 Then '已有人認領
           rsAD.MoveFirst
           Do While Not rsAD.EOF
               If InStr(nGrpACK & ",", rsAD.Fields("st16")) = 0 Then
                  If "" & rsAD.Fields("tfa05") = "Y" Then
                      nGrpYesMan = nGrpYesMan & "," & rsAD.Fields("st01")
                  End If
                  If "" & rsAD.Fields("tfa05") <> "" Then 'Added by Lydia 2023/06/14 非英說案第2次認領會先產生空白記錄
                    nGrpACK = nGrpACK & "," & rsAD.Fields("st16")
                    intCnt = intCnt + 1
                  End If
               End If
               '工程師所屬組別之認領人員
               If "" & rsAD.Fields("st16") = m_UserSt16 Then
                   strExc(2) = strExc(2) & "," & rsAD.Fields("st01")
                   If "" & rsAD.Fields("tfa05") <> nowACK Then
                       oldACK = "" & rsAD.Fields("tfa05")
                   End If
               End If
               rsAD.MoveNext
           Loop
       End If
       strCaseNo = pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "")
       
           If oldACK <> "" Then
              If MsgBox("原認領狀態=" & oldACK & "，是否要修改為" & nowACK & " ？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  SaveDatabase = True
                  Exit Function
              Else
                  If oldACK = "Y" And nowACK <> oldACK Then
                      nGrpYesMan = Replace(nGrpYesMan, strExc(2), "")
                  End If
                  If oldACK = "N" And nowACK <> oldACK Then
                      nGrpYesMan = nGrpYesMan & "," & strUserNum
                  End If
              End If
           End If
           If InStr(nGrpACK & ",", m_UserSt16) = 0 Then
               If nowACK = "Y" Then
                   nGrpYesMan = nGrpYesMan & "," & m_UserNo
               End If
               nGrpACK = nGrpACK & "," & m_UserSt16
               intCnt = intCnt + 1
           End If
           
           If intCnt = intMaxCnt And nGrpYesMan = "" Then
              'Modified by Lydia 2024/06/18
              'If MsgBox(strCaseNo & "無人認領，是否要修改為認領？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
              If MsgBox(strCaseNo & "無人認領，是否要修改為認領？" & vbCrLf & vbCrLf & "選「是」：重新輸入認領Y/N" & vbCrLf & "選「否」：" & IIf(strCon = "1", "進入協調認領階段", "請最高主管進行核判"), vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                  SaveDatabase = True
                  Exit Function
              End If
           End If
           
           'Mark by Lydia 2025/10/03 FCP-074449(10/2)遇見英說參考本的第2次主管認領(TFA09=4)的時分秒一致，造成認領未完成分案
           'If bolUpd = False Then bolUpd = True
           'cnnConnection.BeginTrans
           'end 2025/10/03
             If nGrpYesMan <> "" Then nGrpYesMan = Mid(nGrpYesMan, 2)
             If strExc(2) <> "" Then
                'Modified by Lydia 2025/10/03 debug:加上 tfa01='" & pcp09 & "' and
                strSql = "delete from TransFeeAssign where tfa01='" & pCP09 & "' and tfa04 in (" & GetAddStr(Mid(strExc(2), 2)) & ")  and tfa09=" & CNULL(strCon)
                cnnConnection.Execute strSql
             End If
             strSql = "INSERT INTO TRANSFEEASSIGN(TFA01,TFA02,TFA03,TFA04,TFA05,TFA09) VALUES ('" & pCP09 & "'," & _
                         " TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE, 'HH24MISS')," & _
                         " '" & m_UserNo & "','" & nowACK & "','" & strCon & "')"
             cnnConnection.Execute strSql
            'Added by Lydia 2025/10/03 重新抓「目前認領的狀況」;Ex.FCP-074449
               '目前認領的狀況
               nGrpYesMan = ""
               nGrpACK = ""
               intCnt = 0
               strExc(1) = "select st01,st02,tfa05,st16,decode(st16,'" & m_UserSt16 & "','Y','') as nstate from transfeeassign,staff " & _
                          "where tfa01='" & pCP09 & "' and tfa04=st01(+) and tfa09=" & CNULL(strCon)
               intP = 1
               Set rsAD = ClsLawReadRstMsg(intP, strExc(1))
               If intP = 1 Then '已有人認領
                   rsAD.MoveFirst
                   Do While Not rsAD.EOF
                       If InStr(nGrpACK & ",", rsAD.Fields("st16")) = 0 Then
                          If "" & rsAD.Fields("tfa05") = "Y" Then
                              nGrpYesMan = nGrpYesMan & "," & rsAD.Fields("st01")
                          End If
                          If "" & rsAD.Fields("tfa05") <> "" Then 'Added by Lydia 2023/06/14 非英說案第2次認領會先產生空白記錄
                            nGrpACK = nGrpACK & "," & rsAD.Fields("st16")
                            intCnt = intCnt + 1
                          End If
                       End If
                       rsAD.MoveNext
                   Loop
               End If
               If InStr(nGrpACK & ",", m_UserSt16) = 0 Then
                   If nowACK = "Y" Then
                       nGrpYesMan = nGrpYesMan & "," & m_UserNo
                   End If
                   nGrpACK = nGrpACK & "," & m_UserSt16
                   intCnt = intCnt + 1
               End If
               If nGrpYesMan <> "" Then nGrpYesMan = Mid(nGrpYesMan, 2)
           If bolUpd = False Then bolUpd = True
           cnnConnection.BeginTrans
           'end 2025/10/03
             '新案認領狀態: 0=急件認領, 1=認領, 2=協調認領, 4=非英說案件認領(2023/06/14)
             Select Case strCon
                 Case "0" '0=急件認領: 系統發Email通知各組主管和職代由第一個認領的組別先命名。
                     'Modified by Lydia 2023/05/10 可能有人先輸入不認領;
                     'If intCnt = 1 And nowACK = "Y" Then
                     If intCnt > 0 And nowACK = "Y" Then
                         strExc(3) = PUB_GetEngGrpMan(strExc(4))
                         If InStr(strExc(3) & ";" & strExc(4), m_UserNo) = 0 Then
                            strExc(3) = strExc(3) & ";" & m_UserNo '操作人員非工程師主管或職代
                         End If
                         strSql = "Update TrackingCaseName Set TCN20='" & m_UserSt16 & "' Where TCN05='" & pCP09 & "' "
                         cnnConnection.Execute strSql
                         m_Grp = m_UserSt16
                         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                " values( '" & strUserNum & "','" & strExc(3) & "',to_char(sysdate,'yyyymmdd')" & _
                                ",to_char(sysdate,'hh24miss'),'新案急件" & strCaseNo & "【" & PUB_GetFCPGrpName(m_UserSt16) & "已認領】，請協助翻譯發明名稱以先進行提申','同主旨','" & strExc(4) & "')"
                         cnnConnection.Execute strSql
                     End If
                     'Added by Lydia 2024/08/12 若其他組都不認領，由最後一組先翻譯；ex.FCP-072185因為Wilson不認領，然後Red放著直到認領逾期發通知給主管
                     If FCPforEngNum - intCnt = 1 And nowACK = "N" Then
                        strExc(3) = PUB_GetEngGrpMan(strExc(4))
                        strExc(1) = "select st01,st02,st16 from staff where st16<> '" & m_UserSt16 & "' and st16 is not null and instr('" & strExc(3) & "',st01) > 0 " & _
                                    "and st01 not in (select tfa04 from transfeeassign where tfa09='" & strCon & "' and tfa01='" & pCP09 & "') "
                        intP = 1
                        Set rsAD = ClsLawReadRstMsg(intP, strExc(1))
                        If intP = 1 Then
                           'Added by Lydia 2024/11/11 秒數99=系統自動產生；FCP-072700由Wilison先不認領，系統自動分給Red，然後在提申後又重新認領；
                           strSql = "INSERT INTO TRANSFEEASSIGN(TFA01,TFA02,TFA03,TFA04,TFA05,TFA09) VALUES ('" & pCP09 & "'," & _
                                    " TO_CHAR(SYSDATE,'YYYYMMDD'),substr(TO_CHAR(SYSDATE, 'HH24MISS'),1,4)||'99'," & _
                                    " '" & rsAD.Fields("st01") & "','Y','" & strCon & "')"
                           cnnConnection.Execute strSql
                           'end 2024/11/11
                           strSql = "Update TrackingCaseName Set TCN20='" & rsAD.Fields("st16") & "' Where TCN05='" & pCP09 & "' "
                           cnnConnection.Execute strSql
                           m_Grp = "" & rsAD.Fields("st16")
                           strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                  " values( '" & strUserNum & "','" & rsAD.Fields("st01") & "',to_char(sysdate,'yyyymmdd')" & _
                                  ",to_char(sysdate,'hh24miss'),'新案急件" & strCaseNo & "【" & PUB_GetFCPGrpName("" & rsAD.Fields("st16")) & "已認領】，請協助翻譯發明名稱以先進行提申','因為其他組不認領，所以本案直接交由" & PUB_GetFCPGrpName("" & rsAD.Fields("st16")) & "協助翻譯。','" & strExc(4) & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                     'end 2024/08/12
                 Case "1" '1=認領
                        'Modified by Lydia 2023/06/14 改成變數intMaxCnt
                        If intCnt = intMaxCnt Then
                        '認領期限內最後一個組輸入即可提前結束該階段，進入系統判斷階段。
                        If nGrpYesMan <> "" Then
                            'Modified by Lydia 2023/06/14 等同於一般認領:非英說案=0 or 已確定(3=確定已收文件、4=確定無文件)
                            'If InStr(nGrpYesMan, ",") > 0 Then
                            If InStr(nGrpYesMan, ",") > 0 And InStr("0,3,4", pTCN13) > 0 Then
                                '2組以上認領：系統寄email通知有認領之工程師主管（職代）進行協調和再認領，記錄協調開始起算時間為認領期限再加2小時(若遇中午休息跳過該時段），Email主旨：新案FCP0*****/P******請協調以確認組別，謝謝！
                                '　● 2小時到未再認領則視為放棄，最後一個組輸入即可提前結束該階段。
                                Call PUB_CompWorkTime(Left(Format(ServerTime, "000000"), 4), 120, strExc(4), strSrvDate(1), strExc(3))
                                If strExc(4) <> "" Then
                                   'Added by Lydia 2023/06/14 一般認領+非英說案已確收(TCN13=3,4)，顯示方式比照非英說案協調認領，參考PUB_UpdateReTCN補空白記錄
                                   If InStr("3,4", pTCN13) > 0 Then
                                      strSql = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05,tfa09) " & _
                                               "select tfa01, to_char(sysdate,'yyyymmdd') tfa02, to_char(sysdate, 'HH24MISS') tfa03,tfa04,'' tfa05,'4' as tfa09 from transfeeassign where tfa01='" & pCP09 & "' and tfa09='1' and tfa05='Y' "
                                      cnnConnection.Execute strSql
                                      strSql = "Update trackingcasename Set tcn23='4', tcn25=0, tcn21=" & CNULL(strExc(3)) & ", tcn22=" & CNULL(Left(strExc(4), 4)) & " where tcn05='" & pCP09 & "' "
                                   Else
                                   'end 2023/06/14
                                      strSql = "Update trackingcasename Set tcn23='3', tcn25=0, tcn21=" & CNULL(strExc(3)) & ", tcn22=" & CNULL(Left(strExc(4), 4)) & " where tcn05='" & pCP09 & "' "
                                   End If 'Added by Lydia 2023/06/14
                                   cnnConnection.Execute strSql
                                   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                                           " values( '" & strUserNum & "','" & Replace(nGrpYesMan, ",", ";") & "',to_char(sysdate,'yyyymmdd')" & _
                                           ",to_char(sysdate,'hh24miss'),'" & strSpecSub & "，請協調以確認組別，謝謝！','同主旨')"
                                   cnnConnection.Execute strSql
                                End If
                            Else
                                'Added by Lydia 2023/06/14 非英說案第一次認領階段需要三組主管都有輸入，若有兩組以上認領，則先給第一組進行命名，等待客戶提供文件後進入第二認領階段。
                                If InStr("1,2", pTCN13) > 0 Then '有=1 or 待確定=2
                                   strExc(1) = "select tfa02,tfa03,tfa04,st16 from transfeeassign,staff where tfa01='" & pCP09 & "' and tfa09=" & CNULL(strCon) & " and tfa05='Y' and tfa04=st01(+) order by tfa02 asc,tfa03 asc, st16 asc "
                                   intP = 1
                                   Set rsAD = ClsLawReadRstMsg(intP, strExc(1))
                                   If intP = 1 Then
                                      rsAD.MoveFirst
                                      nGrpYesMan = "" & rsAD.Fields("tfa04")
                                   End If
                                End If
                                'end 2023/06/14
                                '僅1組認領→等同於程序人員於外專新案建檔作業設定工程師組別後，自動發email通知主管進行命名分案作業
                                m_Grp = PUB_GetStaffST16(nGrpYesMan)
                                strSql = "Update TrackingCaseName Set TCN20='" & m_Grp & "' Where TCN05='" & pCP09 & "' "
                                cnnConnection.Execute strSql
                            End If
                        Else
                            '沒有人認領=>主管核判
                            GoTo JumpToMan
                        End If
                     End If
                 'Mark by Lydia 2023/06/14
                 'Case "2", "5" '2=協調認領; 檢查前一階段TFA09=1、5=非英說案協調認領(2次認領有2人以上---2023/06/14)
                 Case "2" '2=協調認領
                     'Modified by Lydia 2023/06/14 改成變數intMaxCnt
                     If intCnt = intMaxCnt Then
                         If InStr(nGrpYesMan, ",") > 0 Or nGrpYesMan = "" Then '超過一組或沒有組認領
JumpToMan:
                           '第2階段未有共識則系統會寄email通知國外部最高主管進行核判，核判期限至隔日下班前；Email主旨：新案FCP0*****/P******請協助確認組別，謝謝！，並且將已存在原始檔區ORI.PDF當做附件。
                           '改成共用模組
                           strAttFList = App.path & "\" & strUserNum
                           Call Pub_ChkExcelPath(strAttFList)
                           '核判期限至隔日下班前
                           strExc(3) = CompWorkDay(2, strSrvDate(1))
                           strExc(4) = "1700"
                           'Modified by Lydia 2023/05/31 TCN23='4' => TCN24='Y'
                           strSql = "Update TrackingCaseName Set TCN24='Y', TCN21=" & strExc(3) & ", TCN22=" & strExc(4) & " Where TCN05='" & pCP09 & "' "
                           cnnConnection.Execute strSql
                        Else
                           '僅1組認領→等同於程序人員於外專新案建檔作業設定工程師組別後，自動發email通知主管進行命名分案作業
                           m_Grp = PUB_GetStaffST16(nGrpYesMan)
                           strSql = "Update TrackingCaseName Set TCN20='" & m_Grp & "' Where TCN05='" & pCP09 & "' "
                           cnnConnection.Execute strSql
                        End If
                     End If
                 'Added by Lydia 2023/06/14
                 Case "4"   '4=非英說案協調認領(2次認領)
                     If intCnt = intMaxCnt Then
                        '認領期限內最後一個組輸入即可提前結束該階段，進入系統判斷階段。
                        If nGrpYesMan <> "" Then
                            If Len(nGrpYesMan) > 10 Then '認領有一人以上，進入協調
                                'Mark by Lydia 2023/06/14 保留; (秀玲)非英說案協調認領(2次認領)視同於一般案的協調認領，直接送最高主管核判。
                                ''非英說協調認領期限+1h
                                'Call PUB_CompWorkTime(Left(Format(ServerTime, "000000"), 4), 60, strExc(4), strSrvDate(1), strExc(3))
                                'If strExc(4) <> "" Then
                                '   strSql = "Update trackingcasename Set tcn23='5',tcn25=0 , tcn21=" & CNULL(strExc(3)) & ", tcn22=" & CNULL(Left(strExc(4), 4)) & " where tcn05='" & pCP09 & "' "
                                '   cnnConnection.Execute strSql
                                '   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                                '           " values( '" & strUserNum & "','" & Replace(nGrpYesMan, ",", ";") & "',to_char(sysdate,'yyyymmdd')" & _
                                '           ",to_char(sysdate,'hh24miss'),'" & strSpecSub & "，請協調以確認組別，謝謝！','同主旨')"
                                '   cnnConnection.Execute strSql
                                'End If
                                GoTo JumpToMan
                                'end 2023/06/14
                            Else
                                '僅1組認領→等同於程序人員於外專新案建檔作業設定工程師組別後，自動發email通知主管進行命名分案作業
                                m_Grp = PUB_GetStaffST16(nGrpYesMan)
                                strSql = "Update TrackingCaseName Set TCN20='" & m_Grp & "' Where TCN05='" & pCP09 & "' "
                                cnnConnection.Execute strSql
                            End If
                        Else
                            '沒有人認領=>主管核判
                            GoTo JumpToMan
                        End If
                     End If
             End Select
             If m_Grp <> "" Then
                 If PUB_UpdateTCNstate("2", pCP01 & pCP02 & pCP03 & pCP04) = False Then
                     GoTo ErrorHandle
                 End If
             End If

           cnnConnection.CommitTrans
           SaveDatabase = True
           If strAttFList <> "" Then
              '改成共用模組
              If PUB_GetTCNEmail(pCP01, pCP02, pCP03, pCP04, "2", strAttFList) = True Then
              End If
           End If
           
       Set rsAD = Nothing
   End If
   
   Exit Function
   
ErrorHandle:
   If bolUpd = True Then
       cnnConnection.RollbackTrans
   End If
   If Err.Number <> 0 Then
       MsgBox "存檔失敗：" & strCaseNo & vbCrLf & Err.Description, vbCritical
   End If
End Function

'Added by Lydia 2023/05/08 讀取目前認領狀態：檢查
Private Function TxtValidate(ByVal pRow As Integer, ByVal pTxt As String) As Boolean
   TxtValidate = False
   strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = "": strExc(9) = ""
   If GetStateTFA(MGRD1.TextMatrix(pRow, colCp09), MGRD1.TextMatrix(pRow, colTCN23), strExc(1), strExc(2), strExc(3), strExc(4), strExc(5)) = True Then
      If "" & MGRD1.TextMatrix(pRow, colTCN23) = "0" And pTxt <> "" And strExc(2) <> "" Then
          strExc(9) = "急件翻譯名稱已被" & strExc(2) & "組認領，"
      End If
      If strExc(1) <> MGRD1.TextMatrix(pRow, colInOLD) Then
          strExc(9) = "目前組別之認領狀態為" & IIf(strExc(1) = "Y", "「" & strExc(1) & "=認領」", "「" & strExc(1) & "=不認領」") & "，"
      End If
   End If
   'Added by Lydia 2023/06/14 判斷目前非英說案的狀態(外文本的對應英/中說)
   If "" & MGRD1.TextMatrix(pRow, colTCN13) <> strExc(5) Then
      Select Case strExc(5)
         Case "3"    '參考來源PUB_GetTCNmTitle
            strExc(9) = strExc(9) & "〔非英說案: 已收參考本〕"
         Case "4"
            strExc(9) = strExc(9) & "〔非英說案: 確定無參考本〕"
      End Select
   End If
   'end 2023/06/14
   If strExc(9) <> "" Then
      MsgBox strExc(9) & vbCrLf & "現在自動執行畫面更新！", vbInformation + vbOKOnly, "輸入檢查"
      Call cmdok_Click(0)
      Exit Function
   End If
   TxtValidate = True
End Function

'Added by Lydia 2023/06/14
Private Function GetTFAcnt(ByVal pCP09 As String, ByVal pGrade As String, Optional ByVal pAns As String) As Integer
  
  GetTFAcnt = 0
  strExc(0) = "SELECT COUNT(*) CNT FROM TRANSFEEASSIGN,STAFF WHERE TFA04=ST01(+) AND TFA09='" & pGrade & "' " & IIf(pAns <> "", "AND TFA05='" & pAns & "'", "") & " AND TFA01='" & pCP09 & "' "
  intI = 1
  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
  If intI = 1 Then
     GetTFAcnt = Val("" & RsTemp.Fields("cnt"))
  End If
End Function

