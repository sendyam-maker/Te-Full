VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050110_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "外翻人員給案維護-輸入"
   ClientHeight    =   4272
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8676
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4272
   ScaleWidth      =   8676
   Begin VB.TextBox txtTF32 
      Height          =   324
      Left            =   1248
      MaxLength       =   7
      TabIndex        =   3
      Top             =   3840
      Width           =   1200
   End
   Begin VB.TextBox txtTF23 
      Alignment       =   1  '靠右對齊
      Height          =   324
      Left            =   1248
      MaxLength       =   6
      TabIndex        =   1
      Top             =   3048
      Width           =   1200
   End
   Begin VB.TextBox txtTF04 
      Alignment       =   1  '靠右對齊
      Height          =   324
      Left            =   1248
      MaxLength       =   6
      TabIndex        =   2
      Top             =   3432
      Width           =   1200
   End
   Begin VB.TextBox txtTF26 
      Height          =   324
      Left            =   1248
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2664
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&Q)"
      CausesValidation=   0   'False
      Height          =   492
      Index           =   1
      Left            =   5712
      TabIndex        =   6
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   492
      Index           =   0
      Left            =   4272
      TabIndex        =   5
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   492
      Index           =   2
      Left            =   7152
      TabIndex        =   7
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "收文日："
      Height          =   252
      Left            =   216
      TabIndex        =   42
      Top             =   1704
      Width           =   972
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   9
      Left            =   7320
      TabIndex        =   41
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "發文日："
      Height          =   252
      Index           =   9
      Left            =   6264
      TabIndex        =   40
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   2
      Left            =   1248
      TabIndex        =   39
      Top             =   2016
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "完稿日："
      Height          =   252
      Index           =   8
      Left            =   216
      TabIndex        =   38
      Top             =   2016
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "收達日："
      Height          =   252
      Index           =   7
      Left            =   216
      TabIndex        =   37
      Top             =   3840
      Width           =   972
   End
   Begin MSForms.TextBox txtTF36 
      Height          =   1164
      Left            =   4248
      TabIndex        =   4
      Top             =   2976
      Width           =   4284
      VariousPropertyBits=   679495707
      MaxLength       =   200
      Size            =   "7556;2053"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "備註："
      Height          =   252
      Index           =   6
      Left            =   4272
      TabIndex        =   36
      Top             =   2712
      Width           =   972
   End
   Begin VB.Label lblTransType 
      Caption         =   "(中翻日)"
      Height          =   252
      Left            =   2496
      TabIndex        =   35
      Top             =   3120
      Width           =   1164
   End
   Begin MSForms.Label lblCP14N 
      Height          =   252
      Left            =   1848
      TabIndex        =   34
      Top             =   2352
      Width           =   1860
      VariousPropertyBits=   27
      Size            =   "3281;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblWordCount 
      Alignment       =   1  '靠右對齊
      Caption         =   "中/原文字數："
      Height          =   252
      Left            =   24
      TabIndex        =   33
      Top             =   3072
      Width           =   1164
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "分配點數："
      Height          =   252
      Index           =   5
      Left            =   216
      TabIndex        =   32
      Top             =   3456
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "完稿期限："
      Height          =   252
      Index           =   4
      Left            =   216
      TabIndex        =   31
      Top             =   2688
      Width           =   972
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請國家："
      Height          =   252
      Index           =   2
      Left            =   4224
      TabIndex        =   30
      Top             =   1404
      Width           =   972
   End
   Begin VB.Label lblPA09 
      Height          =   252
      Left            =   5304
      TabIndex        =   29
      Top             =   1404
      Width           =   492
   End
   Begin MSForms.Label lblNation 
      Height          =   252
      Left            =   5856
      TabIndex        =   28
      Top             =   1404
      Width           =   1212
      VariousPropertyBits=   27
      Size            =   "2138;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "給案日："
      Height          =   252
      Index           =   3
      Left            =   3264
      TabIndex        =   27
      Top             =   1704
      Width           =   972
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   8
      Left            =   4320
      TabIndex        =   26
      Top             =   1704
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦人："
      Height          =   252
      Index           =   1
      Left            =   216
      TabIndex        =   25
      Top             =   2352
      Width           =   972
   End
   Begin VB.Label lblCaseField 
      Caption         =   "F5527"
      Height          =   252
      Index           =   7
      Left            =   1248
      TabIndex        =   24
      Top             =   2352
      Width           =   576
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請案號："
      Height          =   252
      Left            =   4224
      TabIndex        =   23
      Top             =   756
      Width           =   972
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "收文號："
      Height          =   252
      Index           =   0
      Left            =   456
      TabIndex        =   22
      Top             =   432
      Width           =   732
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   252
      Index           =   1
      Left            =   216
      TabIndex        =   21
      Top             =   1404
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所期限："
      Height          =   252
      Index           =   0
      Left            =   6240
      TabIndex        =   20
      Top             =   1704
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "會稿日："
      Height          =   252
      Index           =   2
      Left            =   3264
      TabIndex        =   19
      Top             =   2016
      Width           =   972
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱："
      Height          =   252
      Index           =   0
      Left            =   216
      TabIndex        =   18
      Top             =   1056
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   252
      Left            =   216
      TabIndex        =   17
      Top             =   756
      Width           =   972
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   0
      Left            =   1248
      TabIndex        =   16
      Top             =   756
      Width           =   1812
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   1
      Left            =   5304
      TabIndex        =   15
      Top             =   744
      Width           =   2316
   End
   Begin VB.Label lblCP09 
      Height          =   252
      Left            =   1248
      TabIndex        =   14
      Top             =   432
      Width           =   1200
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   3
      Left            =   1248
      TabIndex        =   13
      Top             =   1704
      Width           =   1200
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   4
      Left            =   1248
      TabIndex        =   12
      Top             =   1404
      Width           =   492
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   5
      Left            =   7296
      TabIndex        =   11
      Top             =   1704
      Width           =   1200
   End
   Begin VB.Label lblCaseField 
      Height          =   252
      Index           =   6
      Left            =   4320
      TabIndex        =   10
      Top             =   2016
      Width           =   1200
   End
   Begin MSForms.Label lblCasePropertyName 
      Height          =   252
      Left            =   1776
      TabIndex        =   9
      Top             =   1404
      Width           =   1716
      VariousPropertyBits=   27
      Size            =   "3027;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboPatentName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1248
      TabIndex        =   8
      Top             =   1056
      Width           =   7272
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12827;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm050110_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2025/11/6
Option Explicit

Public m_CP09 As String
Dim m_TF01 As String, m_TF27 As String, m_TF28 As String

Private Sub cmdOK_Click(Index As Integer)
   Dim varSaveCursor As Integer
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0 '確定
         If TxtValidate Then
            If FormSave Then
               frm050110.UpdatRecord
               frm050110.Show
               Unload Me
            End If
         End If
         
      Case 1 '回前畫面
         frm050110.Show
         Unload Me
         
      Case 2 '結束
         Unload frm050110
         Unload Me
   End Select
   
ExitPort:
   Screen.MousePointer = varSaveCursor
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm050110_1 = Nothing
End Sub

Public Sub ReadAllData(pCP09)
   lblCP09 = pCP09
   strExc(0) = "select * from caseprogress,patent,engineerprogress,transfee,nation,Staff_IdMap,staff where cp09='" & lblCP09 & "'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and ep02(+)=cp09 and tf01(+)=cp09 and na01(+)=pa09 and sim02(+)=cp14 and st01(+)=sim01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      m_TF01 = "" & .Fields("TF01")
      '本所案號
      lblCaseField(0) = .Fields("cp01") & "-" & .Fields("cp02") & IIf(.Fields("cp03") & .Fields("cp04") = "000", "", "-" & .Fields("cp03") & "-" & .Fields("cp04"))
      '申請案號
      lblCaseField(1) = "" & .Fields("pa11")
      '案件名稱
      SetNameToCombo cboPatentName, "" & .Fields("pa05"), "" & .Fields("pa06"), "" & .Fields("pa07")
      '案件性質
      lblCaseField(4) = .Fields("cp10")
      If ClsPDGetCaseProperty(.Fields("cp01"), .Fields("cp10"), strExc(1)) Then
         lblCasePropertyName = strExc(1)
      End If
      '申請國家
      lblPA09 = .Fields("pa09")
      lblNation = .Fields("na03")
      
      '收文日
      lblCaseField(3) = TransDate(.Fields("cp05"), 1)
      '給案日
      lblCaseField(8) = TransDate(.Fields("cp157"), 1)
      '本所期限
      lblCaseField(5) = TransDate("" & .Fields("cp06"), 1)
      '完稿日
      lblCaseField(2) = TransDate("" & .Fields("ep09"), 1)
      '會稿日
      lblCaseField(6) = TransDate("" & .Fields("ep07"), 1)
      '發文日
      lblCaseField(9) = TransDate("" & .Fields("cp27"), 1)
      
      '承辦人
      lblCaseField(7) = "" & .Fields("cp14")
      If ClsPDGetStaffN(lblCaseField(7), strExc(1)) = True Then
         lblCP14N = strExc(1)
      End If
      
      '完稿期限
      txtTF26 = TransDate("" & .Fields("tf26"), 1)
      '中/原文字數
      txtTF23 = "" & .Fields("tf23")
      '分配點數
      txtTF04 = "" & .Fields("tf04")
      '收達日
      txtTF32 = TransDate("" & .Fields("tf32"), 1)
      
      '日本部輸入分配點數
      If Left("" & .Fields("st93"), 1) = "J" Then
         txtTF23.Enabled = False
         txtTF23.BackColor = Me.BackColor
         If txtTF26 <> "" And txtTF04 = "" Then
            txtTF04.SetFocus
         End If
         If txtTF32 = "" Then txtTF32 = TransDate("" & .Fields("cp157"), 1)
      Else
         txtTF04.Enabled = False
         txtTF04.BackColor = Me.BackColor
         If txtTF26 <> "" And txtTF23 = "" Then
            txtTF23.SetFocus
         End If
      End If
      '備註
      txtTF36 = "" & .Fields("TF36")
      
      '外翻中
      If Left(lblCP09, 1) = "C" Then
         '原文語種
         If lblPA09 = "011" Then
            m_TF27 = "2" '日文
            lblTransType = "(日翻中)"
         ElseIf lblPA09 = "231" Then
            m_TF27 = "3" '德文
            lblTransType = "(德翻中)"
         End If
         '翻譯語種
         m_TF28 = "1" '繁體中文
         lblWordCount = "原文字數"
      '中翻外
      Else
         '原文語種
         m_TF27 = "5" '中文
         '翻譯語種
         If lblPA09 = "011" Then
            m_TF28 = "3" '日文
            lblTransType = "(中翻日)"
         ElseIf lblPA09 = "231" Then
            m_TF28 = "4" '德文
            lblTransType = "(中翻德)"
         End If
         lblWordCount = "中文字數"
      End If
   
      End With
   End If
End Sub

Private Function FormSave() As Boolean
   Dim stTF27 As String, stTF28 As String
   
   cnnConnection.BeginTrans
   
   If m_TF01 = "" Then
      strSql = "insert into transfee(TF01,TF04,TF23,TF26,TF27,TF28,TF32,TF36) values('" & lblCP09 & "'" & _
         "," & CNULL(txtTF04, True) & "," & CNULL(txtTF23, True) & "," & CNULL(DBDATE(txtTF26), True) & _
         ",'" & m_TF27 & "','" & m_TF28 & "'," & CNULL(DBDATE(txtTF32), True) & ",'" & ChgSQL(txtTF36) & "')"
   Else
      strSql = "update transfee set tf04=" & CNULL(txtTF04, True) & ",tf23=" & CNULL(txtTF23, True) & ",tf26=" & CNULL(DBDATE(txtTF26), True) & _
         ",tf27='" & m_TF27 & "',tf28='" & m_TF28 & "',tf32=" & CNULL(DBDATE(txtTF32), True) & ",tf36='" & ChgSQL(txtTF36) & "' where tf01='" & m_TF01 & "'"
      Pub_SeekTbLog strSql
   End If
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
     
End Function

Private Function TxtValidate() As Boolean
   Dim bolCancel As Boolean
   
   '目前限定日德案件，若要增加其他語種，翻譯費設定及相關程式也要改
   If InStr("011,231", lblPA09) = 0 Then
      MsgBox "目前限定日德案件才可輸入！", vbCritical
      Exit Function
   End If
   
   If txtTF26 = "" Then
      MsgBox "完稿期限不可空白！", vbCritical
      txtTF26.SetFocus
      Exit Function
   End If
   
   Call txtTF26_Validate(bolCancel)
   If bolCancel Then
      txtTF26.SetFocus
      Exit Function
   ElseIf Val(txtTF26) > Val(lblCaseField(5)) And Val(lblCaseField(5)) > 0 Then
      If MsgBox("完稿期限不應大於本所期限，是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         txtTF26.SetFocus
         Exit Function
      End If
   End If
   
   Call txtTF32_Validate(bolCancel)
   If bolCancel Then
      txtTF32.SetFocus
      Exit Function
   ElseIf Val(txtTF32) < Val(lblCaseField(8)) And Val(lblCaseField(8)) > 0 Then
      If MsgBox("收達日不應小於給案日，是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         txtTF32.SetFocus
         Exit Function
      End If
   End If
   
   If txtTF23.Enabled Then
      If txtTF23 = "" Then
         MsgBox "中/原文字數不可空白！", vbCritical
         txtTF23.SetFocus
         Exit Function
      End If
   End If
   
   If txtTF04.Enabled Then
      If txtTF04 = "" Then
         MsgBox "分配點數不可空白！", vbCritical
         txtTF04.SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Private Sub txtTF04_GotFocus()
   TextInverse txtTF04
End Sub

Private Sub txtTF23_GotFocus()
   TextInverse txtTF23
End Sub

Private Sub txtTF26_GotFocus()
   TextInverse txtTF26
End Sub

Private Sub txtTF26_Validate(Cancel As Boolean)
   If txtTF26 <> "" Then
      If Not ChkDate(txtTF26) Then
         Cancel = True
         Call txtTF26_GotFocus
      End If
   End If
End Sub

Private Sub txtTF32_GotFocus()
   TextInverse txtTF32
End Sub

Private Sub txtTF32_Validate(Cancel As Boolean)
   If txtTF32 <> "" Then
      If Not ChkDate(txtTF32) Then
         Cancel = True
         Call txtTF32_GotFocus
      End If
   End If
End Sub
Private Sub txtTF36_GotFocus()
   TextInverse txtTF36
End Sub
