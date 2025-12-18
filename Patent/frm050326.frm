VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm050326 
   BorderStyle     =   1  '單線固定
   Caption         =   "未收文期限提醒E-Mail"
   ClientHeight    =   5750
   ClientLeft      =   1170
   ClientTop       =   3300
   ClientWidth     =   6500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   6500
   Begin VB.ComboBox cboNP 
      Height          =   300
      ItemData        =   "frm050326.frx":0000
      Left            =   960
      List            =   "frm050326.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   2460
      Width           =   1905
   End
   Begin VB.ComboBox cboNP24 
      Height          =   300
      ItemData        =   "frm050326.frx":0004
      Left            =   960
      List            =   "frm050326.frx":0006
      Style           =   2  '單純下拉式
      TabIndex        =   34
      Top             =   2640
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   735
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   4
         Left            =   2895
         MaxLength       =   2
         TabIndex        =   2
         Top             =   75
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   3
         Left            =   2655
         MaxLength       =   1
         TabIndex        =   1
         Top             =   75
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   2
         Left            =   1815
         MaxLength       =   6
         TabIndex        =   0
         Top             =   75
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   1
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "CFP"
         Top             =   75
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "申請案號："
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Top             =   400
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "本所案號："
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   31
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   0
         Left            =   1335
         MaxLength       =   20
         TabIndex        =   3
         Top             =   400
         Width           =   1935
      End
   End
   Begin VB.CheckBox chkByOutLook 
      Caption         =   "密件副本：操作者本人"
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Value           =   1  '核取
      Width           =   2535
   End
   Begin VB.ComboBox cboRecv 
      Height          =   300
      ItemData        =   "frm050326.frx":0008
      Left            =   1410
      List            =   "frm050326.frx":000A
      TabIndex        =   7
      Text            =   "cboRecv"
      Top             =   3420
      Width           =   1905
   End
   Begin VB.OptionButton Option1 
      Caption         =   "其他收件人："
      Height          =   270
      Index           =   1
      Left            =   30
      TabIndex        =   6
      Top             =   3450
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "智權人員："
      Height          =   270
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   3120
      Value           =   -1  'True
      Width           =   1260
   End
   Begin VB.TextBox Text4 
      Height          =   1176
      Left            =   1245
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   8
      Text            =   "frm050326.frx":000C
      Top             =   4080
      Width           =   5055
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   5400
      TabIndex        =   11
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "E-Mail(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   4380
      TabIndex        =   10
      Top             =   45
      Width           =   975
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   5130
      Top             =   2460
      _ExtentX        =   953
      _ExtentY        =   953
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   4500
      Top             =   2460
      _ExtentX        =   953
      _ExtentY        =   953
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "收件人："
      Height          =   270
      Index           =   4
      Left            =   30
      TabIndex        =   29
      Top             =   2790
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   270
      Index           =   3
      Left            =   30
      TabIndex        =   28
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblApplicant 
      Height          =   270
      Left            =   975
      TabIndex        =   27
      Top             =   2160
      Width           =   5460
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   270
      Index           =   2
      Left            =   30
      TabIndex        =   26
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label lblNation 
      Height          =   270
      Left            =   975
      TabIndex        =   25
      Top             =   1860
      Width           =   5460
   End
   Begin VB.Label lblSubject 
      Caption         =   "本所案號　通知期限"
      Height          =   270
      Left            =   1260
      TabIndex        =   24
      Top             =   3780
      Width           =   5025
   End
   Begin VB.Label Label1 
      Caption         =   "主旨："
      Height          =   270
      Index           =   12
      Left            =   30
      TabIndex        =   23
      Top             =   3780
      Width           =   1185
   End
   Begin VB.Label lblNP09 
      Height          =   270
      Left            =   3900
      TabIndex        =   20
      Top             =   2010
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "下一程序："
      Height          =   270
      Index           =   11
      Left            =   30
      TabIndex        =   22
      Top             =   2490
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "期限："
      Height          =   270
      Index           =   10
      Left            =   3000
      TabIndex        =   21
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "E-Mail內容："
      Height          =   270
      Index           =   9
      Left            =   30
      TabIndex        =   19
      Top             =   4080
      Width           =   1185
   End
   Begin VB.Label lblSaleName 
      Height          =   270
      Left            =   1290
      TabIndex        =   18
      Top             =   3120
      Width           =   1620
   End
   Begin VB.Label lblSaleZone 
      Height          =   270
      Left            =   3900
      TabIndex        =   17
      Top             =   3120
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   270
      Index           =   7
      Left            =   2955
      TabIndex        =   16
      Top             =   3120
      Width           =   915
   End
   Begin VB.Label lblCaseName 
      Caption         =   "(日)："
      Height          =   270
      Index           =   2
      Left            =   825
      TabIndex        =   15
      Top             =   1545
      Width           =   5625
   End
   Begin VB.Label lblCaseName 
      Caption         =   "(英)："
      Height          =   270
      Index           =   1
      Left            =   825
      TabIndex        =   14
      Top             =   1260
      Width           =   5625
   End
   Begin VB.Label lblCaseName 
      Caption         =   "(中)："
      Height          =   270
      Index           =   0
      Left            =   825
      TabIndex        =   13
      Top             =   960
      Width           =   5625
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱"
      Height          =   270
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   960
      Width           =   1185
   End
End
Attribute VB_Name = "frm050326"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Create By Sindy 2012/2/29
Option Explicit

Dim m_blnTxtValidate As Boolean
Dim pa(4) As String
'Add By Sindy 2018/2/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Dim m_Done As Boolean
Public m_PrevForm As Form '前一畫面
'2018/2/22 END


''Add By Sindy 2018/2/22
'Public Sub SetParent(ByRef fm As Form)
'   Set m_PrevForm = fm
'End Sub

'Add By Sindy 2013/3/19
Private Sub cboRecv_LostFocus()
   If Trim(cboRecv.Text) <> "" And Len(Trim(cboRecv.Text)) = 5 Then
      cboRecv.Text = Trim(cboRecv.Text) & " " & GetPrjSalesNM(Trim(cboRecv.Text))
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim strTo As String
   
   Select Case Index
      Case 0 '確定
         'Add By Sindy 2018/2/22
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> Text1(1) & Text1(2) & Text1(3) & Text1(4) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2018/2/22 END
         
         Call cboRecv_LostFocus 'Add By Sindy 2013/3/19
         
         'Added by Lydia 2018/01/16 + 申請案號
         If Option2(0).Value = True Then
            If Me.Text1(0).Text = "" Then
               MsgBox "請輸入申請案號!!!", vbExclamation + vbOKOnly
               Me.Text1(0).SetFocus
               TextInverse Me.Text1(0)
               Exit Sub
            End If
         Else
         'end 2018/01/16
            If Me.Text1(1).Text = "" Then
               MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
               Me.Text1(1).SetFocus
               TextInverse Me.Text1(1)
               Exit Sub
            End If
            If Me.Text1(2).Text = "" Then
               MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
               Me.Text1(2).SetFocus
               TextInverse Me.Text1(2)
               Exit Sub
            End If
         End If  'end 2018/01/16
         
         If Option1(0).Value = True Then
            strTo = Left(Trim(lblSaleName), 5)
         ElseIf Option1(1).Value = True Then
            strTo = Left(Trim(cboRecv.Text), 5)
         End If
         If strTo = "" Then
            MsgBox "收件人空白，無法寄送！"
            Exit Sub
         End If
         
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
                  
         '加可選擇要有寄件備份
         If chkByOutLook.Value = 1 Then
            'Modified by Lydia 2020/03/13 因為OutLook過去和現在版本不同,所以改用密件副本保留
'            DoEvents
'            MAPISession1.LogonUI = False
'            MAPISession1.UserName = strUserNum
'            Err.Clear
'On Error Resume Next
'            MAPISession1.SignOn
'            If Err.NUMBER <> 0 Then
'               MsgBox "EMail發送失敗!!請啟動 OutLook 後重試!!"
'               Screen.MousePointer = vbDefault
'               Exit Sub
'            End If
'            MAPIMessages1.SessionID = MAPISession1.SessionID
'            MAPIMessages1.MsgIndex = -1
'            MAPIMessages1.Compose
'            'Modify By Sindy 2014/1/16
'            'MAPIMessages1.MsgSubject = "◎系統代發◎" & Trim(lblSubject)
'            MAPIMessages1.MsgSubject = "◎" & IIf(Pub_StrUserSt03 = "M51" And PUB_GetST05(strUserNum) <> "", PUB_GetDbTerminal, "") & Trim(lblSubject)
'            '2014/1/16 END
'            MAPIMessages1.MsgNoteText = Text4
'            MAPIMessages1.RecipIndex = 0
'            MAPIMessages1.RecipDisplayName = ChkMailId(strTo)
'            MAPIMessages1.ResolveName
'            MAPIMessages1.Send
'            MAPISession1.SignOff
            PUB_SendMail strUserNum, strTo, "", Trim(lblSubject), Text4, "", , , , , , , , , , , strUserNum
         Else
            PUB_SendMail strUserNum, strTo, "", Trim(lblSubject), Text4, ""
         End If
         
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         
         'Add by Sindy 2018/2/22
         If m_strIR01 <> "" Then
            PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050326"
            Unload Me
            Exit Sub
         End If
         '2018/2/22 END
         
         Call ClearForm(True)
         
      Case 1
         Unload Me
   End Select
End Sub

'2016/3/9 add by sonia 郭說P案改用本所期限
Private Sub Form_Activate()
   If Text1(1) = "P" Or Text1(1) = "PS" Then
      Label1(10) = "本所期限："
      If Text1(2) = "" Then Me.Text4.Text = "　　上述案件，之前曾通知你及客戶下一程序本所期限(本所期限)，今代理人又來函通知前述事情，故再次提醒。"
   Else
      Label1(10) = "法定期限："
      If Text1(2) = "" Then Me.Text4.Text = "　　上述案件，之前曾通知你及客戶下一程序法定期限(法定期限)，今代理人又來函通知前述事情，故再次提醒。"
   End If
   
   'Added by Sindy 2018/2/22
   If m_strIR01 <> "" And m_Done = False Then
      Text1(1).Text = m_strCP01: Text1(1).Locked = True
      Text1(2).Text = m_strCP02: Text1(2).Locked = True
      Text1(3).Text = m_strCP03: Text1(3).Locked = True
      Text1(4).Text = m_strCP04: Text1(4).Locked = True
      'cmdOK(0).Value = True
      Call Text1_Validate(4, False)
      Option2(0).Enabled = False
      Text1(0).Locked = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/2/22 END
End Sub
'2016/3/9 end

Private Sub Form_Load()
   MoveFormToCenter Me
   'Added by Lydia 2018/01/16
   Frame1.BackColor = &H8000000F
   'Remove by Lydia 2018/01/26 改成預設本所案號
   'Option2(0).Value = True
   'end 2018/01/16
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2018/2/23
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
      End If
   End If
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2018/2/23 END
   
   Set frm050326 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
Case 4
   If m_blnTxtValidate = False Then
      Me.Text1(1).SetFocus
      m_blnTxtValidate = True
   End If
End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim stCP13 As String
Dim stCon As String 'Added by Lydia 2018/01/16

Select Case Index
Case 1 '系統類別
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   StrSQLa = "Select * From SystemKind Where SK01='" & Me.Text1(1).Text & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
      Cancel = True
      Me.Text1(1).SetFocus
      TextInverse Me.Text1(1)
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

'Modified by Lydia 2018/01/16 + 申請案號
'Case 4
Case 0, 4
   If Index = 0 Then
        If Trim(Text1(Index)) = "" Then
             Exit Sub
        ElseIf Option2(0).Value = False Then
             Option2(0).Value = True
        End If
   ElseIf Index = 4 Then
        If Text1(2).Text <> "" And Option2(1).Value = False Then
            Option2(1).Value = True
        End If
   End If
'end 2018/01/16

   m_blnTxtValidate = True
   ClearQueryLog (Me.Name)
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   Call ClearForm(False)
   Me.cmdOK(0).Enabled = False
   
   'Added by Lydia 2018/01/16 + 申請案號
   Erase pa
   If Option2(0).Value = True Then
       pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Text1(0).Text
       stCon = "PA11='" & Text1(0).Text & "' "
   Else
   'end 2018/01/16
        pa(1) = Me.Text1(1).Text
        pa(2) = Me.Text1(2).Text
        pa(3) = IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text)
        pa(4) = IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text)
        'Modified by Lydia 2018/01/16
        'pub_QL05 = pub_QL05 & ";" & Label1(0) & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
        pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
        stCon = "PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' "
   End If 'end 2018/01/16
    
   'Modified by Lydia 2018/01/16
   'StrSQLa = "Select PATENT.*,NVL(na03,na04) nation,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) cname " & _
             "From PATENT,Nation,CUSTOMER " & _
             "Where PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' " & _
             "and SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
             "and PA09=NA01(+) "
   StrSQLa = "Select PATENT.*,NVL(na03,na04) nation,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) cname " & _
             "From PATENT,Nation,CUSTOMER " & _
             "Where " & stCon & " and SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
             "and PA09=NA01(+) "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      InsertQueryLog (0)
      MsgBox "資料庫無此案號資料!!!", vbExclamation + vbOKOnly
      Me.lblCaseName(0).Caption = "(中)："
      Me.lblCaseName(1).Caption = "(英)："
      Me.lblCaseName(2).Caption = "(日)："
      Me.lblNation.Caption = ""
      Me.lblApplicant.Caption = ""
      m_blnTxtValidate = False
   Else
      InsertQueryLog (rsA.RecordCount)
      Me.lblCaseName(0).Caption = "(中)：" & rsA("PA05").Value
      Me.lblCaseName(1).Caption = "(英)：" & rsA("PA06").Value
      Me.lblCaseName(2).Caption = "(日)：" & rsA("PA07").Value
      Me.lblNation.Caption = rsA("PA09").Value & " " & rsA("nation").Value
      Me.lblApplicant.Caption = rsA("PA26").Value & " " & rsA("cname").Value
      
      'Added by Lydia 2018/01/16 本所案號
      pa(1) = rsA.Fields("PA01")
      pa(2) = rsA.Fields("PA02")
      pa(3) = rsA.Fields("PA03")
      pa(4) = rsA.Fields("PA04")
      'end 2018/01/16
      
      '下一程序
      '2016/3/9 modify by sonia 郭說P案改用本所期限
      'strExc(0) = "Select distinct NP07,NP09,cpm04 " & _
                  "From NextProgress,casepropertymap " & _
                  "Where NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' " & _
                  "AND NP06 IS NULL " & strNpSqlOfNoSalesDuty & " and cpm01(+)=np02 and cpm02(+)=np07 " & _
                  "Order By NP09 desc "
      'Modify By Sindy 2018/9/7 + NP24
      'modify by sonia 2025/1/24 decode(np02,'P',NP08,'PS',NP08,NP09)改為,decode(np02,'P',NP08,'PS',NP08,nvl(NP09,np08)) (CFP-029236補文件無法限會錯誤)
      strExc(0) = "Select distinct NP07,decode(np02,'P',NP08,'PS',NP08,nvl(NP09,np08)) NP09,cpm04,NP24 " & _
                  "From NextProgress,casepropertymap " & _
                  "Where NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' " & _
                  "AND NP06 IS NULL " & strNpSqlOfNoSalesDuty & " and cpm01(+)=np02 and cpm02(+)=np07 " & _
                  "Order By NP09 desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While Not RsTemp.EOF
            cboNP.AddItem "" & RsTemp("cpm04"), 0
            cboNP24.AddItem "" & RsTemp("NP24"), 0 'Add By Sindy 2018/9/7
            cboNP.ItemData(0) = RsTemp("np09")
            RsTemp.MoveNext
         Loop
         cboNP.ListIndex = 0
         cboNP_Click
      Else
         'Add By Sindy 2018/9/7 檢查是否結案中
         strExc(0) = "Select * " & _
                     "From flow003 " & _
                     "Where F0303='" & pa(1) & pa(2) & pa(3) & pa(4) & "' AND F0304 is null AND F0309<>'" & Flow_歸檔 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "注意：此案號正在結案中！", vbExclamation
         End If
      End If
      
      '主旨
      Me.lblSubject.Caption = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & "　通知期限 (" & "" & rsA("PA05").Value & ")"
      Me.cmdOK(0).Enabled = True
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   Me.lblSaleZone.Caption = ""
   Me.lblSaleName.Caption = ""
   stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   If stCP13 <> "" Then
      strExc(0) = "select A0902,ST02,ST06,ST15,ST01 from staff,acc090 where st01='" & stCP13 & "' and a0901(+)=st15"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Me.lblSaleZone.Caption = RsTemp("ST15") & " " & RsTemp("A0902")
         Me.lblSaleName.Caption = RsTemp("ST01") & " " & RsTemp("ST02")
         
         '其他收件人
         If "" & RsTemp("ST15") <> "" Then
            strExc(0) = "select A0902,ST02,ST06,ST15,ST01 " & _
                          "from staff,acc090 " & _
                        "where st15='" & RsTemp("ST15") & "' and st01<>'" & stCP13 & "' and st04='1' and length(st01)=5 and substr(st01,1,1) in(" & ST01CodeNum1 & ") and a0901(+)=st15 order by st01 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Do While Not RsTemp.EOF
                  cboRecv.AddItem RsTemp("ST01") & " " & RsTemp("ST02")
                  RsTemp.MoveNext
               Loop
               cboRecv.ListIndex = 0
            End If
         End If
      End If
   End If
End Select
If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub cboNP_Click()
Dim str0 As String
   
   If cboNP.Tag <> Trim(cboNP.ListIndex) Then
      intI = cboNP.ListIndex
      If intI >= 0 Then
         If cboNP.ItemData(intI) > 0 Then
            Me.lblNP09.Caption = ChangeTStringToTDateString(cboNP.ItemData(intI) - 19110000)
            
            If Me.lblNP09.Caption <> "" Then
               str0 = DBYEAR(Me.lblNP09.Caption) - 1911 & "年" & DBMONTH(Me.lblNP09.Caption) & "月" & DBDAY(Me.lblNP09.Caption) & "日"
            End If
            '2016/3/9 add by sonia 郭說P案改用本所期限
            'Text4 = "　　上述案件，之前曾通知你及客戶" & cboNP & "期限(" & str0 & ")，今代理人又來函通知前述事情，故再次提醒。"
            If Text1(1) = "P" Or Text1(1) = "PS" Then
               Text4 = "　　上述案件，之前曾通知你及客戶" & cboNP & "本所期限(" & str0 & ")，今代理人又來函通知前述事情，故再次提醒。"
            Else
               Text4 = "　　上述案件，之前曾通知你及客戶" & cboNP & "法定期限(" & str0 & ")，今代理人又來函通知前述事情，故再次提醒。"
            End If
            '2016/3/9 end
         End If
         'Add By Sindy 2018/9/7
         If cboNP24.List(intI) <> "" And Len(cboNP24.List(intI)) = 8 Then
            MsgBox "注意：此案號正在結案中！", vbExclamation
         End If
         '2018/9/7 END
      End If
      cboNP.Tag = intI
   End If
End Sub

'清除畫面欄位資料
Private Sub ClearForm(bolKey As Boolean)
   If bolKey = True Then
      Me.Text1(2).Text = ""
      Me.Text1(3).Text = ""
      Me.Text1(4).Text = ""
      Me.Text1(0).Text = "" 'Added by Lydia 2018/01/16
   End If
   Me.lblCaseName(0).Caption = "(中)："
   Me.lblCaseName(1).Caption = "(英)："
   Me.lblCaseName(2).Caption = "(日)："
   Me.lblNation.Caption = ""
   Me.lblApplicant.Caption = ""
   Me.cboNP.Clear
   Me.cboNP24.Clear 'Add By Sindy 2018/9/7
   Me.cboNP.Tag = ""
   Me.lblNP09.Caption = ""
   Me.Option1(0).Value = True
   Me.lblSaleName.Caption = ""
   Me.lblSaleZone.Caption = ""
   Me.cboRecv.Clear
   Me.lblSubject.Caption = "本所案號　通知期限"
   '2016/3/9 add by sonia 郭說P案改用本所期限
   'Me.Text4.Text = "　　上述案件，之前曾通知你及客戶下一程序期限(法定期限)，今代理人又來函通知前述事情，故再次提醒。"
   If Text1(1) = "P" Or Text1(1) = "PS" Then
      Label1(10) = "本所期限："
      Me.Text4.Text = "　　上述案件，之前曾通知你及客戶下一程序本所期限(本所期限)，今代理人又來函通知前述事情，故再次提醒。"
   Else
      Label1(10) = "法定期限："
      Me.Text4.Text = "　　上述案件，之前曾通知你及客戶下一程序法定期限(法定期限)，今代理人又來函通知前述事情，故再次提醒。"
   End If
   '2016/3/9 end
   Me.chkByOutLook.Value = 1 'Modify By Sindy 2012/7/12 慧汶說要預設為"要寄件備份"
   
   'Added by Lydia 2018/01/16 +申請案號
   If Option2(0).Value = True Then
      Me.Text1(0).SetFocus
   Else
   'end 2018/01/16
      Me.Text1(2).SetFocus
   End If 'end 2018/01/16
End Sub

'Added by Lydia 2018/01/16
Private Sub Option2_Click(Index As Integer)
    If Index = 0 Then
        Text1(0).Enabled = True
        Text1(1).Enabled = False
        Text1(2).Enabled = False
        Text1(3).Enabled = False
        Text1(4).Enabled = False
    Else
        Text1(0).Enabled = False
        Text1(1).Enabled = True
        Text1(2).Enabled = True
        Text1(3).Enabled = True
        Text1(4).Enabled = True
        Me.Text1(2).SetFocus
        Text1_GotFocus 2
    End If
End Sub
