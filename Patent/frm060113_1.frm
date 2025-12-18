VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060113_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "FMP案完稿日/核稿完成日輸入"
   ClientHeight    =   5190
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6930
   Begin VB.TextBox txtCP114 
      Height          =   270
      Left            =   4815
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2940
      Width           =   600
   End
   Begin VB.TextBox txtEP04 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4815
      TabIndex        =   4
      Top             =   2625
      Width           =   870
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   1485
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2940
      Width           =   600
   End
   Begin VB.TextBox txtCP14 
      Height          =   270
      Left            =   1485
      TabIndex        =   2
      Top             =   2625
      Width           =   870
   End
   Begin VB.TextBox txtEP33 
      Height          =   270
      Left            =   4815
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3555
      Width           =   915
   End
   Begin VB.TextBox txtEP08 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "txtEP08"
      Top             =   3570
      Width           =   915
   End
   Begin VB.TextBox txtCP48 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "txtCP48"
      Top             =   3255
      Width           =   915
   End
   Begin VB.TextBox txtEP09 
      Height          =   270
      Left            =   4815
      MaxLength       =   8
      TabIndex        =   0
      Top             =   3255
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5850
      TabIndex        =   9
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4605
      TabIndex        =   8
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3780
      TabIndex        =   7
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   1
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   510
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   510
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   2820
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   510
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   510
      Width           =   375
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   1125
      Left            =   1485
      TabIndex        =   6
      Top             =   3870
      Width           =   4800
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "8467;1984"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核稿時數:"
      Height          =   180
      Index           =   6
      Left            =   4005
      TabIndex        =   48
      Top             =   2985
      Width           =   765
   End
   Begin VB.Label LblCP07 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5400
      TabIndex        =   47
      Top             =   2370
      Width           =   915
   End
   Begin VB.Label LblCP06 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   3480
      TabIndex        =   46
      Top             =   2370
      Width           =   915
   End
   Begin VB.Label LblNP23 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1560
      TabIndex        =   45
      Top             =   2370
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   15
      Left            =   4560
      TabIndex        =   44
      Top             =   2370
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   14
      Left            =   2685
      TabIndex        =   43
      Top             =   2355
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "約定期限:"
      Height          =   180
      Index           =   13
      Left            =   645
      TabIndex        =   42
      Top             =   2355
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   5
      Left            =   645
      TabIndex        =   41
      Top             =   2985
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "核稿人:"
      Height          =   180
      Index           =   7
      Left            =   4185
      TabIndex        =   40
      Top             =   2670
      Width           =   585
   End
   Begin MSForms.Label lblEP04C 
      Height          =   285
      Left            =   5760
      TabIndex        =   39
      Top             =   2670
      Width           =   1035
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblEP33 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "核稿完成日:"
      Height          =   180
      Left            =   3795
      TabIndex        =   38
      Top             =   3585
      Width           =   945
   End
   Begin VB.Label lblEP08 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "核搞期限:"
      Height          =   180
      Left            =   645
      TabIndex        =   37
      Top             =   3585
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "承辦期限:"
      Height          =   180
      Index           =   12
      Left            =   645
      TabIndex        =   35
      Top             =   3255
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   10
      Left            =   645
      TabIndex        =   34
      Top             =   3870
      Width           =   765
   End
   Begin VB.Label lblPA08T 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4860
      TabIndex        =   33
      Top             =   2070
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Index           =   9
      Left            =   4005
      TabIndex        =   32
      Top             =   2070
      Width           =   765
   End
   Begin VB.Label lblCP09 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4860
      TabIndex        =   31
      Top             =   1830
      Width           =   1305
   End
   Begin MSForms.Label lblCP14C 
      Height          =   285
      Left            =   2430
      TabIndex        =   30
      Top             =   2670
      Width           =   1035
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP10C 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1500
      TabIndex        =   29
      Top             =   2070
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "總收文號:"
      Height          =   180
      Index           =   8
      Left            =   4005
      TabIndex        =   28
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Index           =   4
      Left            =   825
      TabIndex        =   27
      Top             =   2670
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   3
      Left            =   645
      TabIndex        =   26
      Top             =   2070
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "收文日:"
      Height          =   180
      Index           =   2
      Left            =   825
      TabIndex        =   25
      Top             =   1830
      Width           =   585
   End
   Begin VB.Label lblCP05T 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1500
      TabIndex        =   24
      Top             =   1830
      Width           =   1665
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   3
      Left            =   1500
      TabIndex        =   23
      Top             =   1470
      Width           =   5205
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "9181;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   2
      Left            =   1500
      TabIndex        =   22
      Top             =   1140
      Width           =   5205
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "9181;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   21
      Top             =   810
      Width           =   5205
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "9181;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "完稿日:"
      Height          =   180
      Index           =   1
      Left            =   4185
      TabIndex        =   20
      Top             =   3255
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   270
      TabIndex        =   19
      Top             =   810
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1065
      TabIndex        =   18
      Top             =   810
      Width           =   345
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1065
      TabIndex        =   17
      Top             =   1140
      Width           =   345
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1065
      TabIndex        =   16
      Top             =   1470
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   645
      TabIndex        =   15
      Top             =   510
      Width           =   765
   End
End
Attribute VB_Name = "frm060113_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/25 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Create by Morgan 2012/5/21
Option Explicit

Dim bolActive As Boolean
Dim m_CP10 As String
Dim bolSetData As Boolean


Private Sub cmdBack_Click()
   Call frm060113.SetGrid(False)
   frm060113.Show
   Unload Me
End Sub

Public Sub SetData(ByRef rstGrid As ADODB.Recordset, ByVal iRow As Integer)
    
    Dim ii As Integer
    
    With frm060113
      For ii = 1 To 4
          txtCaseNo(ii) = .txtCaseNo(ii)
      Next ii
      For ii = 1 To 3
          lblCaseName(ii) = .lblCaseName(ii)
      Next ii
    End With
    With rstGrid
      .Move iRow - 1, adBookmarkFirst
      lblCP05T = "" & .Fields("CP05T")
      lblCP09 = "" & .Fields("CP09")
      lblCP10C = "" & .Fields("CP10C")
      lblPA08T = "" & .Fields("PA08T")
      m_CP10 = "" & .Fields("cp10")
      
      txtCP14 = "" & .Fields("CP14")
      lblCP14C = "" & .Fields("CP14C")
      
      'Add by Amy 2015/01/14
      '約定期限
      lblNP23 = GetNP23("" & .Fields("CP43"), m_CP10)
      '本所期限
      lblCP06 = "" & .Fields("CP06")
      '法定期限
      lblCP07 = "" & .Fields("CP07")
      'end 2015/01/14
      
      '承辦期限
      txtCP48 = TransDate("" & .Fields("CP48"), 1)
      '完稿日
      txtEP09 = TransDate("" & .Fields("EP09"), 1)
      
      '進度備註
      txtCP64 = "" & .Fields("CP64")
      '工作時數
      txtCP113 = "" & .Fields("CP113")
      
      'Added by Morgan 2015/9/18
      '核稿時數
      txtCP114 = "" & .Fields("CP114")
      'end 2015/9/18
      
      If m_CP10 = "201" Then
         txtEP04 = "" & .Fields("EP04")
         txtEP04.Tag = txtEP04
         lblEP04C = "" & .Fields("EP04C")
         txtEP08 = TransDate("" & .Fields("EP08"), 1)
         'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
         If strSrvDate(1) >= FCP核完日改用EP39 Then
            txtEP33 = TransDate("" & .Fields("EP39"), 1)
         Else
         '2023/10/30 END
            txtEP33 = TransDate("" & .Fields("EP33"), 1)
         End If
         
         txtEP09.Enabled = False
         txtCP14.Enabled = False
         txtCP113.Enabled = False
      Else
         Label1(7).Visible = False
         txtEP04.Visible = False
         lblEP04C.Visible = False
         lblEP08.Visible = False
         txtEP08.Visible = False
         lblEP33.Visible = False
         txtEP33.Visible = False
         'Added by Morgan 2015/9/18
         Label1(6).Visible = False
         txtCP114.Visible = False
         'end 2015/9/18
      End If
      
      
      txtEP09.Tag = txtEP09.Text
      txtEP33.Tag = txtEP33.Text
      txtCP14.Tag = txtCP14.Text
      txtCP113.Tag = txtCP113.Text
      txtCP114.Tag = txtCP114.Text
      txtCP64.Tag = txtCP64.Text
      
    End With
    bolSetData = True
End Sub

Private Function FormSave() As Boolean
   Dim strUpdate As String
   
On Error GoTo flgError

cnnConnection.BeginTrans
   
   If txtEP09 <> txtEP09.Tag Then
      strSql = " Update engineerPROGRESS Set ep09=" & CNULL(DBDATE(txtEP09), True) & " Where ep02='" & lblCP09 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   If txtEP33 <> txtEP33.Tag Then
      'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
      If strSrvDate(1) >= FCP核完日改用EP39 Then
         strSql = " Update engineerPROGRESS Set ep39=" & CNULL(DBDATE(txtEP33), True) & " Where ep02='" & lblCP09 & "'"
      Else
      '2023/10/30 END
         strSql = " Update engineerPROGRESS Set ep33=" & CNULL(DBDATE(txtEP33), True) & " Where ep02='" & lblCP09 & "'"
      End If
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   
   strUpdate = ""
   If txtCP14 <> txtCP14.Tag Then
      strUpdate = strUpdate & ",cp14='" & txtCP14 & "'"
   End If
   
   If txtCP113 <> txtCP113.Tag Then
      strUpdate = strUpdate & ",cp113=" & CNULL(txtCP113, True)
   End If
   
   'Added by Morgan 2015/9/18
   If txtCP114 <> txtCP114.Tag Then
      strUpdate = strUpdate & ",cp114=" & CNULL(txtCP114, True)
   End If
   'end 2015/9/18
   
   If txtCP64 <> txtCP64.Tag Then
      strUpdate = strUpdate & ",cp64='" & ChgSQL(txtCP64) & "'"
   End If
   
   If strUpdate <> "" Then
      strUpdate = Mid(strUpdate, 2)
      strSql = " Update caseprogress Set " & strUpdate & " Where cp09='" & lblCP09 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   
   cnnConnection.CommitTrans
   FormSave = True

flgError:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If

End Function

Private Sub cmdExit_Click()
   Unload frm060113
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If TxtValidate = True Then
      'Add by Sindy 2021/11/25 檢查畫面上的物件是否含有Unicode文字
      If PUB_ChkUniText(Me, True, True) = False Then
         Exit Sub
      End If
      
      If FormSave() = True Then
         cmdBack_Click
      Else
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If txtCP113.Enabled = True Then
      txtCP113_Validate bCancel
      If bCancel = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
      End If
   End If
   
   If txtCP14.Enabled = True Then
      txtCP14_Validate bCancel
      If bCancel = True Then
         txtCP14.SetFocus
         txtCP14_GotFocus
         Exit Function
      End If
   End If
   
   If txtEP09.Enabled = True Then
      txtEP09_Validate bCancel
      If bCancel = True Then
         txtEP09.SetFocus
         txtEP09_GotFocus
         Exit Function
      End If
   End If
   
   If txtEP33.Visible And txtEP33.Enabled Then
      txtEP33_Validate bCancel
      If bCancel = True Then
         txtEP33.SetFocus
         txtEP33_GotFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Private Sub Form_Activate()
   If bolActive = False Then
      bolActive = True
      If txtEP33.Visible And txtEP33.Enabled Then
         If txtEP33 = "" Then txtEP33 = strSrvDate(2)  'Added by Morgan 2015/9/18 預設系統日
         'Modified by Morgan 2015/9/18
         'txtEP33.SetFocus
         'txtEP33_GotFocus
         txtCP114.SetFocus
         txtCP114_GotFocus
         'end 2015/9/18
      ElseIf txtEP09.Enabled = True Then
         If txtEP09 = "" Then txtEP09 = strSrvDate(2)  'Added by Morgan 2014/6/12 預設系統日--Susan
         'Modified by Morgan 2015/9/18
         'txtEP09.SetFocus
         'txtEP09_GotFocus
         txtCP113.SetFocus
         txtCP113_GotFocus
         'end 2015/9/18
      End If
      
   End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm060113_1 = Nothing
End Sub

Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
   CloseIme
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
   Cancel = Not PUB_CheckCP113(txtCP113, txtCaseNo(1), m_CP10, txtCP14)
End Sub

Private Sub txtCP114_GotFocus()
   TextInverse txtCP114
   CloseIme
End Sub

Private Sub txtCP114_Validate(Cancel As Boolean)
   Static strCP114 As String
   If txtCP114 <> "" And txtCP114.Tag <> txtCP114 Then
      If Not IsNumeric(txtCP114) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP114.SetFocus
         txtCP114_GotFocus
         Cancel = True
         Exit Sub
      ElseIf Val(txtCP114) > 25 Then
         If strCP114 <> txtCP114 Then
            If MsgBox("核稿時數超過25小時，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               Cancel = True
               Exit Sub
            End If
            strCP114 = txtCP114
         End If
      End If
   End If
End Sub

Private Sub txtCP14_Change()
   If lblCP14C <> "" Then lblCP14C = ""
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
   CloseIme
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP14_Validate(Cancel As Boolean)
   If txtCP14 = "" Then
      MsgBox "承辦人不可空白 !", vbCritical
      Cancel = True
   Else
      lblCP14C = GetStaffName(txtCP14, True)
   End If
End Sub

Private Sub txtEP09_GotFocus()
   TextInverse txtEP09
   CloseIme
End Sub

Private Sub txtEP09_Validate(Cancel As Boolean)
   If txtEP09 <> "" Then
      If Not ChkDate(txtEP09) Then
         txtEP09_GotFocus
         Cancel = True
      ElseIf Val(DBDATE(txtEP09)) > Val(strSrvDate(1)) Then
         MsgBox "完稿日不可大於系統日！"
         txtEP09_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub txtEP33_GotFocus()
   TextInverse txtEP33
   CloseIme
End Sub

Private Sub txtEP33_Validate(Cancel As Boolean)
   If txtEP33 <> "" Then
      If Not ChkDate(txtEP33) Then
         txtEP33_GotFocus
         Cancel = True
      ElseIf Val(DBDATE(txtEP33)) > Val(strSrvDate(1)) Then
         MsgBox "核稿完成日不可大於系統日！"
         txtEP33_GotFocus
         Cancel = True
      End If
   End If
End Sub

'Add by Amy 2015/01/14 抓取下一程序約定期限
Private Function GetNP23(pNP01 As String, pNP07 As String) As String
    Dim adoquery As New ADODB.Recordset
    Dim strQuery As String, intQ As Integer
    
    strQuery = "Select sqldatet(NP23) NP23 From NextProgress Where NP01='" & pNP01 & "' And NP07=" & pNP07 & " "
    intQ = 1
    Set adoquery = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        GetNP23 = "" & adoquery.Fields("NP23")
    End If
    adoquery.Close
End Function
