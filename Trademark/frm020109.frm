VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020109 
   BorderStyle     =   1  '單線固定
   Caption         =   "TC陸代申請書輸入"
   ClientHeight    =   4665
   ClientLeft      =   -600
   ClientTop       =   2985
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7815
   Begin VB.TextBox txtDate 
      Height          =   330
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   4
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   3
      Top             =   636
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   3
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   2
      Top             =   636
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   2
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   1
      Top             =   636
      Width           =   855
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   636
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "著作權資料(&F)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   60
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔及E-mail (&O)"
      Height          =   400
      Index           =   0
      Left            =   4965
      TabIndex        =   6
      Top             =   60
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6585
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   28
      Top             =   3600
      Width           =   915
      VariousPropertyBits=   27
      Caption         =   "lblFM2(11)"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   27
      Top             =   3120
      Width           =   3495
      VariousPropertyBits=   27
      Caption         =   "lblFM2(10)"
      Size            =   "6165;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   9
      Left            =   2070
      TabIndex        =   26
      Top             =   2712
      Width           =   3495
      VariousPropertyBits=   27
      Caption         =   "lblFM2(9)"
      Size            =   "6165;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   8
      Left            =   2070
      TabIndex        =   25
      Top             =   2304
      Width           =   3495
      VariousPropertyBits=   27
      Caption         =   "lblFM2(8)"
      Size            =   "6165;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   7
      Left            =   2070
      TabIndex        =   24
      Top             =   1896
      Width           =   3495
      VariousPropertyBits=   27
      Caption         =   "lblFM2(7)"
      Size            =   "6165;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   6
      Left            =   2070
      TabIndex        =   23
      Top             =   1488
      Width           =   3495
      VariousPropertyBits=   27
      Caption         =   "lblFM2(6)"
      Size            =   "6165;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   22
      Top             =   3120
      Width           =   915
      VariousPropertyBits=   27
      Caption         =   "lblFM2(5)"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   21
      Top             =   2712
      Width           =   915
      VariousPropertyBits=   27
      Caption         =   "lblFM2(4)"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   20
      Top             =   2304
      Width           =   915
      VariousPropertyBits=   27
      Caption         =   "lblFM2(3)"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   19
      Top             =   1488
      Width           =   915
      VariousPropertyBits=   27
      Caption         =   "lblFM2(1)"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   18
      Top             =   1080
      Width           =   5415
      VariousPropertyBits=   27
      Caption         =   "lblFM2(0)"
      Size            =   "9551;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   17
      Top             =   1896
      Width           =   915
      VariousPropertyBits=   27
      Caption         =   "lblFM2(2)"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "收到陸代申請書日："
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   4125
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "陸代申請書期限："
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "申請人5："
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "申請人4："
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2712
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "申請人3："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   2304
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "申請人2："
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1896
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "申請人1："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1488
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   972
   End
End
Attribute VB_Name = "frm020109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; lblData(index) =>lblFM2
'Create by Lydia 2016/09/07 TC案陸代申請書輸入
Option Explicit
Dim m_CaseNo As String  '帶出的本所案號
Dim m_Nation As String    '申請國家
Dim m_NP01 As String
Dim m_NP22 As String
Dim now_CP13 As String  '回覆委任代理人的業務
Private Sub cmdok_Click(Index As Integer)
Dim objOutLook As Object
Dim objMail As Object

   Select Case Index
      Case 0 '存檔及email
         If TxtValidate = True Then
            If FormSave Then
                '呼叫新郵件
                Set objOutLook = CreateObject("Outlook.Application")
                Set objMail = objOutLook.CreateItem(0)
                strExc(0) = txtCode(1) & "-" & Val(txtCode(2)) & IIf(txtCode(3) & txtCode(4) = "000", "", "-" & txtCode(3) & "-" & txtCode(4))
                strExc(0) = strExc(0) & " － 大陸著作權文件用印"
                objMail.Subject = strExc(0)
                objMail.To = now_CP13
                objMail.Display
                Call ClearForm("A")
                txtCode(2).SetFocus
                txtCode_GotFocus 2
            End If
         End If
         
      Case 1 '結束
         Unload Me
      Case 2 '著作權資料
         If Len(txtCode(2)) <> 6 Then
            MsgBox "請輸入本所案號!"
            txtCode(2).SetFocus
            txtCode_GotFocus 2
            Exit Sub
         End If
         
         txtCode(3) = Right("0" & txtCode(3), 1)
         txtCode(4) = Right("00" & txtCode(4), 2)
         
         If QueryData = False Then
            txtCode(2).SetFocus
            txtCode_GotFocus 2
            Exit Sub
         End If
         txtDate.SetFocus
         txtDate_GotFocus
         
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim strCP09 As String, strCP10 As String, strCP13 As String, strCP06 As String
   Dim strCP12 As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
   
        strSql = "Update Nextprogress set np06='Y' where np01='" & m_NP01 & "' and np22 = '" & m_NP22 & "' AND NP06 IS NULL AND NP07='994'"
        cnnConnection.Execute strSql, intI
       
       '新增回覆委任代理人
        strCP10 = "732"
        strCP13 = PUB_GetAKindSalesNo(txtCode(1), txtCode(2), txtCode(3), txtCode(4))
        strCP12 = PUB_GetStaffST15(strCP13, "1")
        strCP06 = CompDate(2, 15, txtDate)
        strCP09 = AutoNo("B", 6)
        strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP43) " & _
                 "VALUES ('" & txtCode(1) & "','" & txtCode(2) & "','" & txtCode(3) & "','" & txtCode(4) & "','" & txtDate & "','" & strCP06 & "','" & strCP06 & "'" & _
                 ",'" & strCP09 & "','" & strCP10 & "','" & strCP12 & "','" & strCP13 & "','" & strUserNum & "','" & m_NP01 & "')"
    
        cnnConnection.Execute strSql, intI
         
   cnnConnection.CommitTrans
     
   now_CP13 = strCP13
   FormSave = True
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
   
End Function

Private Function QueryData() As Boolean
Dim rsA As New ADODB.Recordset
Dim sqlA As String

   m_CaseNo = ""
   m_NP01 = ""
   m_NP22 = ""
   Call ClearForm
   
   QueryData = False

   sqlA = "select sp01,sp02,sp03,sp04,sp05,sp09,sp08,sp58,sp59,sp65,sp66 from servicepractice" & _
          " where sp01='" & txtCode(1) & "' and sp02='" & txtCode(2) & "' and sp03='" & txtCode(3) & "' and sp04='" & txtCode(4) & "'"
   intI = 0
   Set rsA = ClsLawReadRstMsg(intI, sqlA)
   If intI = 1 Then
      m_Nation = "" & rsA.Fields("sp09")
      If m_Nation <> "020" Then
         MsgBox "本案非大陸案!"
         Exit Function
      Else
        lblFM2(0) = "" & rsA.Fields("sp05")
        lblFM2(1) = "" & rsA.Fields("sp08")
        lblFM2(2) = "" & rsA.Fields("sp58")
        lblFM2(3) = "" & rsA.Fields("sp59")
        lblFM2(4) = "" & rsA.Fields("sp65")
        lblFM2(5) = "" & rsA.Fields("sp66")
        'Added by Lydia 2021/10/07 申請人名稱
        For intI = 1 To 5
            If Trim(lblFM2(intI).Caption) <> "" Then
                 sqlA = GetCustomerName(lblFM2(intI).Caption)
                 lblFM2(intI + 5).Caption = sqlA
            End If
        Next intI
        'end 2021/10/07
         sqlA = "select * from nextprogress where np02='" & txtCode(1) & "' and np03='" & txtCode(2) & "' and np04='" & txtCode(3) & "' and np05='" & txtCode(4) & "'" & _
                "  and np07='994' and nvl(np06,'0')='0' order by np08 "
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, sqlA)
         If intI = 0 Then
            MsgBox "本案的下一程序沒有陸代申請書!"
            Exit Function
         Else
            m_NP01 = "" & rsA.Fields("NP01")
            m_NP22 = "" & rsA.Fields("NP22")

            lblFM2(11) = ChangeWStringToWDateString("" & rsA.Fields("NP08"))
            txtDate = strSrvDate(1)
            m_CaseNo = txtCode(1) & txtCode(2) & txtCode(3) & txtCode(4)
            QueryData = True
         End If
      End If
   End If
   
End Function

Private Function TxtValidate() As Boolean
Dim tmpBol As Boolean

   TxtValidate = False
   
   If txtCode(1) & txtCode(2) & txtCode(3) & txtCode(4) = "TC" Then
      MsgBox "請輸入本所案號!"
      txtCode(2).SetFocus
      txtCode_GotFocus 2
      Exit Function
   End If
   
   If m_CaseNo <> txtCode(1) & txtCode(2) & txtCode(3) & txtCode(4) Then
        If MsgBox("本所案號與著作權資料不一致,是否查詢著作權資料?", vbYesNo + vbDefaultButton1) = vbYes Then
            If QueryData = False Then
                txtCode(2).SetFocus
                txtCode_GotFocus 2
                Exit Function
            Else
                txtDate.SetFocus
                txtDate_GotFocus
                Exit Function
            End If
        Else
            Exit Function
        End If
   End If
   
   If lblFM2(11).Caption = "" Then
      MsgBox "陸代申請書無期限!"
      txtCode(2).SetFocus
      txtCode_GotFocus 2
      Exit Function
   End If
   
   If txtDate = "" Then
      MsgBox "收到陸代申請書日不可空白!"
      Exit Function
   Else
      txtDate_Validate tmpBol
      If tmpBol Then
         txtDate.SetFocus
         txtDate_GotFocus
         Exit Function
      End If
      
      If txtDate > strSrvDate(1) Then
         MsgBox "收到陸代申請書日不可大於系統日!"
         txtDate.SetFocus
         txtDate_GotFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
   Exit Function

JumpExit:
    txtCode(2).SetFocus
    txtCode_GotFocus 2
End Function

Private Sub Form_Activate()
   If txtCode(1) = "" Then
      txtCode(1) = "TC"
      txtCode(2).SetFocus
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Call ClearForm 'Added by Lydia 2021/10/07
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020109 = Nothing
End Sub

'Remove by Lydia 2021/10/07 改在QueryData
'Private Sub lblFM2_Change(Index As Integer)
'Dim strTmp As String
'
'   If Index > 0 And Index < 6 Then
'      If lblFM2(Index).Caption <> "" Then
'         If ClsPDGetCustomer(lblFM2(Index).Caption, strTmp) Then
'            lblFM2(Index + 5).Caption = strTmp
'            Debug.Print Index
'         End If
'      End If
'   End If
'End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
   CloseIme
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ClearForm(Optional ByVal cTyp As String)
Dim oLbl As Control

   If cTyp = "A" Then
      txtCode(2) = ""
      txtCode(3) = ""
      txtCode(4) = ""
   End If
   
   txtDate = ""
   For Each oLbl In lblFM2
       oLbl.Caption = ""
   Next
   
End Sub

Private Sub txtDate_GotFocus()
    TextInverse txtDate
    CloseIme
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If ChkDate(txtDate) Then
       txtDate = DBDATE(txtDate)
    Else
       Cancel = True
       Exit Sub
    End If
End Sub
