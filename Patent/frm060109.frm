VERSION 5.00
Begin VB.Form frm060109 
   BorderStyle     =   1  '單線固定
   Caption         =   "整批重新委任收/發文"
   ClientHeight    =   3105
   ClientLeft      =   255
   ClientTop       =   990
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   9375
   Begin VB.CommandButton cmdOK 
      Caption         =   "補件清單(P)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   6
      Left            =   7875
      TabIndex        =   27
      Top             =   2550
      Width           =   1305
   End
   Begin VB.TextBox txtAppNo 
      Height          =   270
      Left            =   1935
      TabIndex        =   4
      Top             =   2730
      Width           =   1665
   End
   Begin VB.TextBox txtCP27 
      Height          =   270
      Left            =   1305
      MaxLength       =   6
      TabIndex        =   2
      Top             =   2100
      Width           =   1125
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   540
      Width           =   660
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請書(&F)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   3
      Left            =   4680
      TabIndex        =   17
      Top             =   60
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發文室送件清單(&L)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   5
      Left            =   135
      TabIndex        =   16
      Top             =   60
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "整批發文(&D)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   2
      Left            =   6210
      TabIndex        =   15
      Top             =   60
      Width           =   1155
   End
   Begin VB.TextBox txtLR09 
      Height          =   270
      Left            =   2475
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtLR01 
      Height          =   270
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   0
      Top             =   540
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件清單(P)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   4
      Left            =   3420
      TabIndex        =   7
      Top             =   60
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   8505
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "整批收文(&R)"
      Enabled         =   0   'False
      Height          =   400
      Index           =   1
      Left            =   7380
      TabIndex        =   5
      Top             =   60
      Width           =   1110
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   2
      Left            =   3510
      TabIndex        =   26
      Top             =   1080
      Width           =   4275
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   1
      Left            =   3510
      TabIndex        =   25
      Top             =   840
      Width           =   4275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(日):"
      Height          =   180
      Index           =   0
      Left            =   3105
      TabIndex        =   24
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   3105
      TabIndex        =   23
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   3105
      TabIndex        =   22
      Top             =   585
      Width           =   345
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "總委任書申請案號:"
      Height          =   180
      Left            =   315
      TabIndex        =   21
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   315
      TabIndex        =   20
      Top             =   2130
      Width           =   585
   End
   Begin VB.Label lblAppQty3 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      Height          =   180
      Left            =   2115
      TabIndex        =   19
      Top             =   1860
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "已發文案件數:"
      Height          =   180
      Left            =   315
      TabIndex        =   18
      Top             =   1875
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "已收文未發文案件數:"
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   14
      Top             =   1620
      Width           =   1665
   End
   Begin VB.Label lblAppQty2 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      Height          =   180
      Left            =   2115
      TabIndex        =   13
      Top             =   1605
      Width           =   945
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   0
      Left            =   3510
      TabIndex        =   12
      Top             =   585
      Width           =   4275
   End
   Begin VB.Label lblAppQty 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      Height          =   180
      Left            =   2115
      TabIndex        =   11
      Top             =   1380
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "未收文案件數:"
      Height          =   180
      Left            =   315
      TabIndex        =   10
      Top             =   1365
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "是否管制委任書期限4個月            ( Y:是 )"
      Height          =   180
      Left            =   315
      TabIndex        =   9
      Top             =   2430
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號:"
      Height          =   180
      Left            =   315
      TabIndex        =   8
      Top             =   585
      Width           =   765
   End
End
Attribute VB_Name = "frm060109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan2010/8/13 日期欄已修改
'Add by Morgan 2007/6/23
Option Explicit
Dim rsData As New ADODB.Recordset
Dim PLeft() As Integer, PColName() As String
Dim iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim m_stSubTitle As String
Dim m_strRefDate As String '報表上所謂的當日

Private Sub cmdok_Click(Index As Integer)
   Dim strAppDate As String
   Screen.MousePointer = vbHourglass
   '結束
   If Index = 0 Then
      Unload Me
   Else
      Select Case Index
         Case 1 '整批收文
            If TxtValidate = True Then
               Set rsData = GetRst(1, intI)
               If intI = 1 Then
                  If FormSave = False Then
                     MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
                  Else
                     m_strRefDate = strSrvDate(1)
                     PrintCaseList
                     cmdSearch_Click
                  End If
               End If
            End If
         
         Case 2 '整批發文
            If TxtValidate(2) = True Then
               Set rsData = GetRst(5, intI)
               If intI = 1 Then
                  If FormSave1 = True Then
                     m_strRefDate = DBDATE(txtCP27)
                     PrintCaseList
                     cmdSearch_Click
                  End If
               End If
            End If
            
         Case 3 '申請書
            If TxtValidate(3) = True Then
               PrintAppForm
            End If
            
         Case 4 '案件清單
            '2007/7/6 MODIFY BY SONIA
            'm_strRefDate = strSrvDate(1)
            'm_strRefDate = DBDATE(txtCP27)
            'Modify by Morgan 2007/8/13 改若有輸發文日時用發文日否則用系統日判斷當日
            If txtCP27 <> "" Then
               m_strRefDate = DBDATE(txtCP27)
            Else
               m_strRefDate = strSrvDate(1)
            End If
            'End 2007/8/13
            PrintCaseList
         
         Case 5 '送件清單
            strAppDate = InputBox("發文日：", "請輸入發文日期", strSrvDate(2), Me.Left, Me.Top + Me.Height + 1000)
            If strAppDate <> "" Then
               If ChkDate(strAppDate) = True Then
                  PrintAppList strAppDate
               End If
            End If
            
         'Add by Morgan 2007/8/17
         Case 6 '補件清單
            PrintCaseList1
            
      End Select
   End If
   Screen.MousePointer = vbDefault
End Sub
'設定客戶名稱
Private Sub setCustName()
   strExc(0) = "select cu04 n1,cu05||' '||cu88||' '||cu89||' '||cu90 n2,cu06 n3 from customer where cu01='" & Left(txtLR01, 8) & "' and cu02='" & Mid(txtLR01, 9, 1) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      lblCustName(0) = "" & RsTemp.Fields(0)
      lblCustName(1) = "" & RsTemp.Fields(1)
      lblCustName(2) = "" & RsTemp.Fields(2)
   Else
      lblCustName(0) = ""
      lblCustName(1) = ""
      lblCustName(2) = ""
   End If
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If txtLR01 = "" Then
      MsgBox "客戶編號不可空白！"
      txtLR01.SetFocus
   Else
      txtLR01 = Left(txtLR01 & "000", 9)
      setCustName
      txtLR01.Tag = txtLR01
      SetCaseQty
      SetEnable True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   FormClear
   'txtCP27 = strSrvDate(2)   '2007/7/9 靜芳說不預設
   If CheckUse("frm060109", strPrint, False) = True Then
      cmdOK(5).Enabled = True
   Else
      cmdOK(5).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060109 = Nothing
End Sub

Private Sub txtAppNo_GotFocus()
   TextInverse txtAppNo
End Sub

Private Sub txtCP27_GotFocus()
   TextInverse txtCP27
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If txtCP27 <> "" Then
      If Not ChkDate(txtCP27) Then
        Cancel = True
      End If
   End If
End Sub

Private Sub txtLR01_Change()
   If txtLR01.Tag <> "" Then
      FormClear
   End If
End Sub

Private Sub txtLR01_GotFocus()
   TextInverse txtLR01
End Sub

Private Sub txtLR01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtLR09_GotFocus()
   TextInverse txtLR09
End Sub

Private Sub txtLR09_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function TxtValidate(Optional p_iAct As Integer = 1) As Boolean
   Dim bCancel As Boolean
   If txtLR01 = "" Then
      MsgBox "客戶編號不可空白！", vbExclamation
      txtLR01.SetFocus
      Exit Function
   End If
   If Len(txtLR01) <> 6 And Len(txtLR01) <> 9 Then
      MsgBox "客戶編號錯誤！", vbExclamation
      txtLR01_GotFocus
      txtLR01.SetFocus
      Exit Function
   End If
   If txtLR01.Tag <> txtLR01 Then
      MsgBox "客戶編號已變更，請重新按查詢鈕！", vbExclamation
      cmdSearch.SetFocus
      Exit Function
   End If
   '發文
   If p_iAct = 2 Then
      If txtCP27 = "" Then
         MsgBox "請輸入發文日！", vbExclamation
         txtCP27.SetFocus
         Exit Function
      End If
      txtCP27_Validate bCancel
      If bCancel = True Then
         txtCP27_GotFocus
         txtCP27.SetFocus
         Exit Function
      End If
      If txtLR09 = "" Then
         If MsgBox("是否確定不管制期限？", vbYesNo + vbDefaultButton2) = vbNo Then
            txtLR09.SetFocus
            Exit Function
         End If
      End If
   End If
   '申請書
   If p_iAct = 3 Then
      If txtAppNo = "" Then
         If MsgBox("是否確定沒有總委任書申請案號？", vbYesNo + vbDefaultButton2) = vbNo Then
            txtAppNo.SetFocus
            Exit Function
         End If
      Else
         strExc(0) = "select pa26,pa27,pa28,pa29,pa30 from patent where pa11='" & txtAppNo & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For intI = 0 To 4
               If "" & RsTemp.Fields(intI) = txtLR01 Then
                  Exit For
               End If
            Next
            If intI = 5 Then
               MsgBox "總委任書申請案號並非該客戶案件！", vbExclamation
               txtAppNo_GotFocus
               txtAppNo.SetFocus
               Exit Function
            End If
         Else
            MsgBox "總委任書申請案號不存在！", vbExclamation
            txtAppNo_GotFocus
            txtAppNo.SetFocus
            Exit Function
         End If
      End If
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   
   Dim cp(1 To 110) As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   With rsData
      .MoveFirst
      Do While Not .EOF
         cp(1) = .Fields("lc01")
         cp(2) = .Fields("lc02")
         cp(3) = .Fields("lc03")
         cp(4) = .Fields("lc04")
         cp(5) = strSrvDate(1)
         cp(9) = AutoNo("B", 6)
         cp(10) = "928"
         cp(13) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
         cp(12) = GetSalesArea(cp(13))
         cp(14) = strUserNum
         cp(20) = "N"
         cp(26) = "N"
         cp(110) = "76012,81040"
         
         strSql = "insert into caseprogress(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP110)" & _
            " VALUES(" & CNULL(cp(1)) & "," & CNULL(cp(2)) & "," & CNULL(cp(3)) & "," & CNULL(cp(4)) & "," & cp(5) & _
            "," & CNULL(cp(9)) & "," & CNULL(cp(10)) & "," & CNULL(cp(12)) & "," & CNULL(cp(13)) & "," & CNULL(cp(14)) & _
            "," & CNULL(cp(20)) & "," & CNULL(cp(26)) & "," & CNULL(cp(110)) & ")"
            
         cnnConnection.Execute strSql, intI
         .MoveNext
      Loop
   End With
   cnnConnection.CommitTrans
   FormSave = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Function FormSave1() As Boolean
   Dim lMax As Long
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
   
   With rsData
      .MoveFirst
      Do While Not .EOF
         If IsNull(.Fields("cp27")) Then
            If txtLR09 = "Y" Then
               '檢查是否已有委任書補文件且相關總收文號為重新委任
               strExc(0) = "select * from nextprogress,caseprogress where np02='" & .Fields("cp01") & "' and np03='" & .Fields("cp02") & "' and np04='" & .Fields("cp03") & "' and np05='" & .Fields("cp04") & "' and np07='202' and np15='委任書' and cp09(+)=np01 and cp10='928'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  'Modify by Morgan 2007/7/19 改40天
                 'strExc(1) = CompDate(1, 1, strSrvDate(1))
                 'Modify by Morgan 2007/8/24 改發文日起4個月
                 'strExc(1) = CompDate(2, 40, strSrvDate(1))
                 strExc(1) = CompDate(1, 4, DBDATE(txtCP27))
                 strExc(2) = CompDate(2, -2, strExc(1))
                 lMax = GetNextProgressNo
                 
                 strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
                    " VALUES ('" & .Fields("cp09") & "','" & .Fields("cp01") & "','" & .Fields("cp02") & "','" & .Fields("cp03") & "','" & .Fields("cp04") & "'" & _
                    ",'202'," & strExc(2) & "," & strExc(1) & ",'" & .Fields("cp13") & "','委任書'," & _
                 lMax & ")"
                 
                  cnnConnection.Execute strSql, intI
               End If
            End If
            strSql = "Update CaseProgress Set CP27=" & DBDATE(txtCP27) & _
               " where cp09='" & .Fields("cp09") & "' and cp10='928' and cp57 is null and cp27 is null"
            cnnConnection.Execute strSql, intI
         Else
            strSql = "Update CaseProgress Set CP27=null" & _
               " where cp09='" & .Fields("cp09") & "' and cp10='928' and cp57 is null and cp27=19221111"
            cnnConnection.Execute strSql, intI
         End If
         .MoveNext
      Loop
   End With
   cnnConnection.CommitTrans
   FormSave1 = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub FormClear()
   txtLR01.Tag = ""
   lblCustName(0) = ""
   lblCustName(1) = ""
   lblCustName(2) = ""
   txtLR09 = ""
   lblAppQty = ""
   lblAppQty2 = ""
   lblAppQty3 = ""
   SetEnable False
End Sub

Private Sub SetEnable(Optional p_bolEnable As Boolean = True)
   Dim ii As Integer
   For ii = 1 To 4
      cmdOK(ii).Enabled = False
   Next
   cmdOK(6).Enabled = False
   If p_bolEnable = True Then
      '有未收文
      If Val(lblAppQty) > 0 Then
         cmdOK(1).Enabled = True '整批收文
      End If
      '有未發文
      If Val(lblAppQty2) > 0 Then
         cmdOK(2).Enabled = True '整批發文
         cmdOK(3).Enabled = True '申請書
      End If
      If Val(lblAppQty) = 0 And (Val(lblAppQty2) > 0 Or Val(lblAppQty3) > 0) Then
         cmdOK(4).Enabled = True '案件清單
      End If
      
      Set rsData = GetRst(6, intI)
      If intI = 1 Then
         cmdOK(6).Enabled = True '補文件清單
      End If
   End If
End Sub

Private Function GetRst(p_iType As Integer, p_iRlt As Integer) As ADODB.Recordset
   Dim stCustNo As String
   stCustNo = Left(txtLR01 & "000", 9)
   Select Case p_iType
      Case 1 '未收文(未閉卷,有專用,未收文928)
'2007/7/6 MODIFY BY SONIA 案件程序中僅有:專利調查,調卷,鑑定報告,翻譯之程序,則刪除該案,即若還有其他程序者,則該案要保留.--96/7/6靜芳
'         strExc(0) = "select X.* from (" & _
'            " select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc07='" & stCustNo & "'" & _
'            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc08='" & stCustNo & "'" & _
'            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc09='" & stCustNo & "'" & _
'            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc10='" & stCustNo & "'" & _
'            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc11='" & stCustNo & "'" & _
'            ") X,patent" & _
'            " where pa01(+)=lc01 and pa02(+)=lc02 and pa03(+)=lc03 and pa04(+)=lc04" & _
'            " and pa57 is null and (pa25 is null or pa25>" & strSrvDate(1) & ")" & _
'            " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and cp10='928')"
'Modify by Morgan 2007/7/27 加已銷卷的也不要
'Modify by Morgan 2007/7/30 改已銷卷的但專用期還在的仍要
         strExc(0) = "select X.* from (" & _
            " select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc07='" & stCustNo & "'" & _
            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc08='" & stCustNo & "'" & _
            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc09='" & stCustNo & "'" & _
            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc10='" & stCustNo & "'" & _
            " union select lc01,lc02,lc03,lc04 from lincase where lc01='FCP' and lc11='" & stCustNo & "'" & _
            ") X,patent" & _
            " where pa01(+)=lc01 and pa02(+)=lc02 and pa03(+)=lc03 and pa04(+)=lc04" & _
            " and pa57 is null and ((pa25 is null and pa108 is null) or pa25>" & strSrvDate(1) & ")" & _
            " and not exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and cp10='928')" & _
            " and exists(select * from caseprogress where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and cp10 NOT IN ('201','927','903','904','906'))"
      
      Case 2 '已收文未發文
         strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp13,decode(cp27,null,2,1) C1,substr(cp09,1,1) C2,decode(pa16,'Y',1,2) C3" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||cp04) C4 from (" & _
            " select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa26='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa27='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa28='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa29='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa30='" & stCustNo & "' and pa57 is null" & _
            ") X,caseprogress" & _
            " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and cp27 is null and cp57 is null" & _
            " order by C1,C2,C3,C4"
      
      Case 3 '已發文
         strExc(0) = "select 1 from (" & _
            " select pa01,pa02,pa03,pa04 from patent where pa01='FCP' and pa26='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04 from patent where pa01='FCP' and pa27='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04 from patent where pa01='FCP' and pa28='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04 from patent where pa01='FCP' and pa29='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04 from patent where pa01='FCP' and pa30='" & stCustNo & "' and pa57 is null" & _
            ") X,caseprogress" & _
            " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and cp27>0"
      
      Case 4 '案件清單
         strExc(0) = "select decode(cp27,null,2,1) C1,substr(cp09,1,1) C2,decode(pa16,'1',1,2) C3,cp05,s1.st02 S1,cp27,pa75,na03" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||cp04) C4,X1,s2.st02 S2" & _
            " from (select pa01,pa02,pa03,pa04,pa16,pa75,decode(pa27,null,'','Y') X1 from patent where pa01='FCP' and pa26='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16,pa75,'Y' from patent where pa01='FCP' and pa27='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16,pa75,'Y' from patent where pa01='FCP' and pa28='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16,pa75,'Y' from patent where pa01='FCP' and pa29='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16,pa75,'Y' from patent where pa01='FCP' and pa30='" & stCustNo & "' and pa57 is null" & _
            ") X,caseprogress,staff s1,fagent,nation,staff s2" & _
            " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928'" & _
            " and s1.st01(+)=cp14 and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1) and na01(+)=fa10" & _
            " and s2.st01(+)=na16 order by C1,C2,C3,C4"
            
      Case 5 '整批發文
         strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp13,decode(cp27,null,2,1) C1,substr(cp09,1,1) C2,decode(pa16,'Y',1,2) C3" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||cp04) C4,cp27 from (" & _
            " select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa26='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa27='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa28='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa29='" & stCustNo & "' and pa57 is null" & _
            " union select pa01,pa02,pa03,pa04,pa16 from patent where pa01='FCP' and pa30='" & stCustNo & "' and pa57 is null" & _
            ") X,caseprogress" & _
            " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and (cp27 is null or cp27=19221111) and cp57 is null" & _
            " order by C1,C2,C3,C4"
            
      Case 6 '補件清單
         'Modify by Morgan 2010/8/13 百年蟲 sqldatet(cp27) -->substrb(' '||sqldatet(cp27),-9)
         strExc(0) = "select sqldatet(cp27) C1,pa75 C2,na03 C3,s1.st02 C4,s2.st02 C5" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||cp04) C6" & _
            ",decode(pa27,null,'','Y') C7,sqldatet(np08) C8,sqldatet(np09) C9" & _
            " from patent,caseprogress,nextprogress,FAGENT,nation,staff s1,staff s2" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and cp27>20070000" & _
            " and np01(+)=cp09 and np07='202' and np06 is null and Fa01(+)=SUBSTR(pa75,1,8) AND FA02(+)=SUBSTR(PA75,9,1) AND NA01(+)=FA10" & _
            " and s1.st01(+)=cp14 and s2.st01(+)=na16 order by 1,6"
            
   End Select
   p_iRlt = 1
   Set GetRst = ClsLawReadRstMsg(p_iRlt, strExc(0))
   
End Function

Private Sub SetCaseQty()
   
   lblAppQty = ""
   lblAppQty2 = ""
   lblAppQty3 = ""
   
   Set RsTemp = GetRst(1, intI)
   If intI = 1 Then
      lblAppQty = RsTemp.RecordCount
   Else
      lblAppQty = 0
   End If
   
   Set RsTemp = GetRst(2, intI)
   If intI = 1 Then
      lblAppQty2 = RsTemp.RecordCount
   Else
      lblAppQty2 = 0
   End If
   
      Set RsTemp = GetRst(3, intI)
   If intI = 1 Then
      lblAppQty3 = RsTemp.RecordCount
   Else
      lblAppQty3 = 0
   End If
End Sub
'申請書
Private Sub PrintAppForm()
   Set rsData = GetRst(2, intI)
   If intI = 1 Then
      With rsData
         Do While Not .EOF
            EndLetter "01", .Fields("cp09"), "01", strUserNum
             
            If txtAppNo <> "" Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('01','" & .Fields("cp09") & "','01','" & strUserNum & "'" & _
                  ",'總委任書申請案號','" & txtAppNo & "')"
               cnnConnection.Execute strSql, intI
            End If
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('01','" & .Fields("cp09") & "','01','" & strUserNum & "'" & _
               ",'流水號','" & Format(.AbsolutePosition, "000000") & "')"
            cnnConnection.Execute strSql, intI
            NowPrint .Fields("cp09"), "01", "01", False, strUserNum
            .MoveNext
         Loop
         PUB_BatchPrint "5"
      End With
   End If
End Sub
'清單
Private Sub PrintCaseList1()
   Set rsData = GetRst(6, intI)
   If intI = 1 Then
      DoPrint1
   End If
End Sub

'案件清單
Private Sub PrintCaseList()
   Set rsData = GetRst(4, intI)
   If intI = 1 Then
      DoPrint
   End If
End Sub
'送件清單
Private Sub PrintAppList(Optional p_AppDate As String)
   Dim strTmp As String, iRec As Integer, iRecs As Integer, iCol As Integer
   Dim lngTot As Long
   Dim iCaseNo As Integer '案件筆數
   Dim iPage As Integer '頁次
   Dim nCopys As Integer '份數
   Dim iCopys As Integer
   Dim stCon As String
   
   If p_AppDate <> "" Then
      stCon = " and cp27=" & DBDATE(p_AppDate)
   Else
      stCon = " and cp27=" & strSrvDate(1)
   End If
   
   strExc(0) = "SELECT LPAD(CP01||'-'||CP02||'-'||CP03||'-'||CP04,15,' ') C01, 0 C02, 0 C03" & _
            ",RPAD(NVL(PA11,' '),12,' ') C04, RPAD(CPM03,12,' ') C05, RPAD(NVL(CU04,' '),20,' ') C06" & _
            ",RPAD(PA05,40,' ') C07" & _
            " from caseprogress,PATENT,customer,casepropertymap" & _
            " where cp01='FCP' and cp10='928'" & stCon & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and cu01(+)=substr(pa26,1,8) and cu02=substr(pa26,9,1)" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
            " order by 1"
            
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '2007/7/6 modify by sonia 張弘郁說改一份
   'nCopys = 2
   nCopys = 1
   If intI = 1 Then
      Printer.Orientation = 2
      Printer.Font = "細明體"
      For iCopys = 1 To nCopys
         If iCopys > 1 Then Printer.NewPage
         iPage = 1: iRec = 0: lngTot = 0: iCaseNo = 0: iRecs = 0
         With RsTemp
         .MoveFirst
         PrintHead
         Do While Not .EOF
            iRec = iRec + 1: iRecs = iRecs + 1
            If iRec > 26 Then
               PrintTail iPage
               Printer.NewPage
               iPage = iPage + 1
               PrintHead
               iRec = 0
            End If
            strTmp = ""
            For iCol = 0 To 6
               Select Case iCol
                  Case 1
                     strTmp = strTmp & Right(Space(9) & Format(Val("" & .Fields(iCol)), "###,###"), 9) & Space(1)
                  Case 2
                     strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(iCol)), "###,###"), 8) & Space(1)
                  Case Else
                     strTmp = strTmp & .Fields(iCol) & Space(1)
               End Select
            Next
            Printer.CurrentY = Printer.CurrentY + 60
            Printer.Print strTmp
            iCaseNo = iCaseNo + 1
            .MoveNext
         Loop
         PrintTail iPage, lngTot, iRecs, iCaseNo
         End With
      Next
      Printer.EndDoc
      MsgBox "列印完成！"
   Else
      MsgBox "無可列印資料！", vbInformation
   End If
  
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      Printer.CurrentX = PLeft(iCol)
      Printer.CurrentY = iPrint
      Printer.Print strData(iCol)
    Next
End Sub

Private Sub PrintSubTotal(p_iCount As Integer, Optional p_bSubTotal As Boolean = True, Optional p_ExCount As Integer = -1)
   PrintNewLine
   
   If iPrint + 600 > lngPageHeight Then
      PrintPageHeader
      PrintPageHeader1
   End If
   
   DrawLine
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   If p_bSubTotal = True Then
      strExc(0) = "小計： " & p_iCount & " 筆"
      If p_ExCount > -1 Then
         strExc(0) = strExc(0) & " (當日發文 " & p_ExCount & " 筆)"
      End If
      Printer.Print strExc(0)
   Else
      Printer.Print "總計： " & p_iCount & " 筆"
   End If
   iPrint = iPrint + 300
End Sub

Private Sub DoPrint()
   Dim Grp1 As String, Grp2 As String, Grp3 As String, iCount As Integer, iTotal As Integer, iCurCount As Integer
   Dim subTitle As String, iNo As Integer
   Dim iOrientation As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With rsData
      GetPleft
      ReDim strTemp(1 To 8)
      iPage = 0
      iCount = 0
      iTotal = 0
      iNo = 0
      m_stSubTitle = ""
      PrintPageHeader
      Do While Not .EOF
         If Grp1 <> .Fields("C1") Or (.Fields("C1") <> 1 And (Grp2 <> .Fields("C2") Or Grp3 <> .Fields("C3"))) Then
            If iCount > 0 Then
               PrintSubTotal iCount, , iCurCount
               iTotal = iTotal + iCount
               iCount = 0
               iCurCount = -1
            End If
            If .Fields("C1") = 1 Then
               m_stSubTitle = "***已發文***"
               iCurCount = 0
            Else
               If .Fields("C3") = "1" Then
                  m_stSubTitle = "***" & .Fields("C2") & "類未發文(已准)***"
               Else
                  m_stSubTitle = "***" & .Fields("C2") & "類未發文(未准)***"
               End If
            End If
            PrintPageHeader1
            Grp1 = .Fields("C1")
            Grp2 = .Fields("C2")
            Grp3 = .Fields("C3")
         End If
         iCount = iCount + 1
         If "" & .Fields("cp27") = m_strRefDate Then
            iCurCount = iCurCount + 1
         End If
         strTemp(1) = ChangeWStringToTDateString(Format("" & .Fields("cp05"), ""))
         strTemp(2) = "" & .Fields("S1")
         strTemp(3) = ChangeWStringToTDateString(Format("" & .Fields("cp27"), ""))
         strTemp(4) = "" & .Fields("pa75")
         strTemp(5) = "" & .Fields("na03")
         strTemp(6) = "" & .Fields("S2")
         '未發文未核准的加流水號
         If .Fields("C1") = 2 And .Fields("C3") = "2" Then
            iNo = iNo + 1
            strTemp(7) = iNo & "." & .Fields("C4")
         Else
            strTemp(7) = "" & .Fields("C4")
         End If
         strTemp(8) = "" & .Fields("X1")
         PrintDetail strTemp
         .MoveNext
      Loop
      If iCount > 0 Then
         PrintSubTotal iCount, , iCurCount
         iTotal = iTotal + iCount
         PrintSubTotal iTotal, False
      End If
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
End Sub

Sub PrintPageHeader(Optional iOpt As Integer = 0)
   Dim strPTmp As String
   Dim strName(3) As String, idx As Integer
   
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   If iOpt = 1 Then
      strPTmp = "重新委任補文件清單"
   Else
      strPTmp = "重新委任案件清單"
   End If
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   iPrint = iPrint + 100
   strPTmp = "(客戶編號：" & txtLR01 & ")"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   idx = 0
   '中
   If lblCustName(0) <> "" Then
      idx = idx + 1
      strName(idx) = Left(lblCustName(0), 20)
   End If
   '英
   If lblCustName(1) <> "" Then
      idx = idx + 1
      strName(idx) = Left(lblCustName(1), 40)
   End If
   '日
   If lblCustName(2) <> "" Then
      idx = idx + 1
      strName(idx) = Left(lblCustName(2), 20)
   End If
   
   If strName(1) <> "" Then
      PrintNewLine
      strPTmp = strName(1)
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   If strName(2) <> "" Then
      strPTmp = strName(2)
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
  
   If strName(3) <> "" Then
      strPTmp = strName(3)
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   iPage = iPage + 1
   Printer.Print "頁    次：" & str(iPage)
End Sub

Sub PrintPageHeader1()
   
   If iPrint + 1000 > lngPageHeight Then
      Printer.NewPage
      PrintPageHeader
   End If
   
   If m_stSubTitle <> "" Then
      PrintNewLine
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print m_stSubTitle
   End If
   
   PrintNewLine
   DrawLine
   
   For intI = 1 To UBound(PLeft)
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print PColName(intI)
   Next
   
   PrintNewLine
   DrawLine
   iPrint = iPrint - 300
   
End Sub

Sub GetPleft(Optional iOpt As Integer = 0)
   Dim iUB As Integer
   
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   If iOpt = 0 Then
      iUB = 8
   Else
      iUB = 11
   End If
   ReDim PLeft(1 To iUB)
   ReDim PColName(1 To iUB)
   
   If iOpt = 0 Then
      PLeft(1) = ciStartX
      PColName(1) = "收文日"
      PLeft(2) = PLeft(1) + 1100
      PColName(2) = "承辦人"
      PLeft(3) = PLeft(2) + 1100
      PColName(3) = "發文日"
      PLeft(4) = PLeft(3) + 1100
      PColName(4) = "代理人編號"
      PLeft(5) = PLeft(4) + 1300
      PColName(5) = "代理人國籍"
      PLeft(6) = PLeft(5) + 1300
      PColName(6) = "管制人"
      PLeft(7) = PLeft(6) + 1100
      PColName(7) = "本所案號"
      PLeft(8) = PLeft(7) + 1800
      PColName(8) = "多申請人(Y)"
   Else
      PLeft(1) = ciStartX
      PColName(1) = "發文日"
      PLeft(2) = PLeft(1) + 1100
      PColName(2) = "代理人編號"
      PLeft(3) = PLeft(2) + 1300
      PColName(3) = "代理人國籍"
      PLeft(4) = PLeft(3) + 1300
      PColName(4) = "承辦人"
      PLeft(5) = PLeft(4) + 1100
      PColName(5) = "管制人"
      PLeft(6) = PLeft(5) + 1100
      PColName(6) = "本所案號"
      PLeft(7) = PLeft(6) + 1800
      PColName(7) = "多申請人(Y)"
      PLeft(8) = PLeft(7) + 1800
      PColName(8) = "本所期限"
      PLeft(9) = PLeft(8) + 1200
      PColName(9) = "法定期限"
      PLeft(10) = PLeft(9) + 1200
      PColName(10) = "IPO所限"
      PLeft(11) = PLeft(10) + 1200
      PColName(11) = "IPO法限"
   End If
End Sub

Private Sub DrawLine()
   Printer.DrawStyle = vbSolid
   Printer.DrawWidth = 4
   Printer.Line (ciStartX, iPrint)-(lngPageWidth - 500, iPrint)
   iPrint = iPrint + 100
End Sub

Private Sub PrintNewLine(Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      DrawLine
      Printer.NewPage
      PrintPageHeader
      PrintPageHeader1
      iPrint = iPrint + lngLineHeight
   End If
End Sub

Private Sub PrintHead()

      Dim strTitle As String
      
      strTitle = "外專 " & ChangeTStringToTDateString(strSrvDate(2)) & " 重新委任送件清單"
      Printer.Print
      Printer.FontSize = 16
      Printer.CurrentX = 5000
      Printer.Print strTitle
      Printer.FontSize = 12
      Printer.CurrentX = 13000
      Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
      Printer.CurrentX = 13000
      Printer.Print "列印時間：" & Format(ServerTime, "##:##:##")
      Printer.CurrentX = 13000
      Printer.Print "列印人員：" & strUserName
      Printer.Print
      Printer.Print "本所案號        規費      規費小計 申請案號     案件性質　　 申請人               案件名稱                                "
      Printer.Print "--------------- --------- -------- ------------ ------------ -------------------- -----------------------------------------------------"
      
End Sub

Private Sub PrintTail(iPage As Integer, Optional p_lngTot As Long, Optional p_iRecs As Integer, Optional p_iCaseCnt As Integer)
   Dim stData As String
   Printer.FontSize = 12
   Printer.Print String(135, "-")
   If p_iRecs > 0 Then
      stData = "總計" & Space(4) & Right(Space(8) & "規費 " & Format(p_lngTot, "###,###,##0") & " 元", 16) & Space(4) & Right(Space(2) & "明細 " & Format(p_iRecs) & " 筆", 9) & Space(4) & Right(Space(2) & "案號 " & Format(p_iCaseCnt) & " 筆", 9)
      Printer.Print stData
   End If
   Printer.CurrentY = 10700
   Printer.CurrentX = 7000
   Printer.Print "第 " & Format(iPage) & " 頁"
End Sub

Private Sub PrintTail1(p_iCount)
   PrintNewLine
   If iPrint + 600 > lngPageHeight Then
      DrawLine
      PrintPageHeader
      PrintPageHeader1
   End If
   DrawLine
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "共 " & p_iCount & " 筆"
End Sub

Private Sub DoPrint1()
   Dim strTemp() As String, iOrientation As Integer
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With rsData
      GetPleft 1
      ReDim strTemp(1 To .Fields.Count)
      iPage = 0
      PrintPageHeader 1
      PrintPageHeader1
      Do While Not .EOF
         For intI = 1 To .Fields.Count
            strTemp(intI) = "" & .Fields(intI - 1)
         Next
         PrintDetail strTemp
         .MoveNext
      Loop
      PrintTail1 .RecordCount
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
End Sub
