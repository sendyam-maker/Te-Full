VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040109_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "一案兩申請案件資料維護"
   ClientHeight    =   4440
   ClientLeft      =   735
   ClientTop       =   2130
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6795
   Begin VB.Frame fraIn 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      TabIndex        =   18
      Top             =   1680
      Width           =   2412
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   0
         MaxLength       =   3
         TabIndex        =   5
         Top             =   0
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   7
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   8
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   6
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   7
         Top             =   0
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   5
         Left            =   480
         MaxLength       =   6
         TabIndex        =   6
         Top             =   0
         Width           =   852
      End
   End
   Begin VB.Frame fraOut 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      TabIndex        =   17
      Top             =   720
      Width           =   2412
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   480
         MaxLength       =   6
         TabIndex        =   1
         Top             =   0
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Top             =   0
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   3
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   3
         TabIndex        =   0
         Top             =   0
         Width           =   492
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   5772
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   3720
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   4548
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   300
      Left            =   1080
      TabIndex        =   27
      Top             =   3360
      Width           =   5535
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9763;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboIn 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Top             =   2040
      Width           =   5535
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9763;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboOut 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9763;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人:"
      Height          =   180
      Left            =   240
      TabIndex        =   32
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Left            =   240
      TabIndex        =   31
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "處理狀況:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label lblCountry 
      AutoSize        =   -1  'True
      Caption         =   "lblCountry"
      Height          =   180
      Left            =   1080
      TabIndex        =   29
      Top             =   3000
      Width           =   2385
   End
   Begin MSForms.Label lblApplicant 
      Height          =   210
      Left            =   1080
      TabIndex        =   28
      Top             =   2640
      Width           =   5445
      VariousPropertyBits=   27
      Caption         =   "lblApplicant"
      Size            =   "9604;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   6600
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   6600
      Y1              =   2460
      Y2              =   2460
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   3
      Left            =   4560
      TabIndex        =   26
      Top             =   4080
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3704;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   2
      Left            =   1320
      TabIndex        =   25
      Top             =   4080
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3228;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   1
      Left            =   4560
      TabIndex        =   24
      Top             =   3840
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3704;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   0
      Left            =   1320
      TabIndex        =   23
      Top             =   3840
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3228;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Time:"
      Height          =   180
      Index           =   4
      Left            =   3480
      TabIndex        =   22
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Name:"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   21
      Top             =   4080
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Time:"
      Height          =   180
      Index           =   2
      Left            =   3480
      TabIndex        =   20
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Name:"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請案二:"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案一:"
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frm040109_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/21 改成Form2.0 (cboOut,cboIn,Combo3,lblApplicant,Label3)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intLeaveKind As Integer   '0:結束1:回上一畫面
Public m_CM10 As String 'Added by Morgan 2015/9/10
Public strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Public strCode5 As String, strCode6 As String, strCode7 As String, strCode8 As String
Public strCbo As String
Public intChoose As String
Public frmParent As Form

Private Function Process() As Boolean
   Dim strCode() As String, i As Integer, bolSave As Boolean
   
   Select Case intChoose
      Case 1   '新增
         If TxtValidate = False Then Exit Function 'Added by Morgan 2021/12/21
         
         ReDim strCode(8) As String
         For i = 0 To 7
            strCode(i) = txtCode(i)
         Next
         strCode(8) = Combo3.Text
         If CheckLengthIsOK(strCode(8), 20) Then
            'Modified by Morgan 2015/9/10 +擬制喪失新穎性
            If bInsertCaseRelationData(strCode(), IIf(m_CM10 <> "", Val(m_CM10), 3)) Then
               bolSave = True
            End If
         Else
            Combo3.SetFocus
            bolSave = False
         End If
         
      Case 2   '修改
         If TxtValidate = False Then Exit Function 'Added by Morgan 2021/12/21
         
         ReDim strCode(16) As String
         For i = 0 To 7
            strCode(i) = txtCode(i)
         Next
         strCode(8) = strCode1
         strCode(9) = strCode2
         strCode(10) = strCode3
         strCode(11) = strCode4
         strCode(12) = strCode5
         strCode(13) = strCode6
         strCode(14) = strCode7
         strCode(15) = strCode8
         strCode(16) = Combo3.Text
         If CheckLengthIsOK(strCode(16), 20) Then
            If CheckCaseCode Then
               'Modified by Morgan 2015/9/10 +擬制喪失新穎性
               If UpdateCaseRelationData(strCode(), IIf(m_CM10 <> "", Val(m_CM10), 3)) Then
                  bolSave = True
               End If
            End If
         Else
            Combo3.SetFocus
            bolSave = False
         End If
      Case 4   '刪除
         ReDim strCode(7) As String
         For i = 0 To 7
            strCode(i) = txtCode(i)
         Next
         If PUB_ChkExist(strCode(), 3) Then
            If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
               'Modified by Morgan 2017/10/13 +第3參數傳False
               If PUB_DeleteCaseRelation(strCode(), 3, False) Then
                  bolSave = True
               End If
            End If
         End If
   End Select
   Process = bolSave
   
End Function

Private Sub cmdOK_Click(Index As Integer)
   
   Select Case Index
      Case 0
         If Process Then
            intLeaveKind = 1
            Unload Me
         End If
      Case 1
         intLeaveKind = 1
         Unload Me
      Case 2
         intLeaveKind = 0
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()

   Dim bolGoOn As Boolean, Lbl As Object
   Dim strTxt(1 To 17) As String, i As Integer
   
   fraIn.Enabled = False
   fraOut.Enabled = False

   txtCode(0) = strCode1
   txtCode(1) = strCode2
   txtCode(2) = strCode3
   txtCode(3) = strCode4
   txtCode(4) = strCode5
   txtCode(5) = strCode6
   txtCode(6) = strCode7
   txtCode(7) = strCode8
   Combo3.Text = ""
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   lblApplicant = "": lblCountry = ""
   For i = 1 To 8
      strTxt(i) = txtCode(i - 1)
   Next
   
   
   'Added by Morgan 2015/9/10
   If m_CM10 = "6" Then
      strTxt(10) = "6"
   Else
      strTxt(10) = "3"
   End If
   'end 2015/9/10
   
   Select Case intChoose
      Case 1
         Me.Caption = Me.Caption & "(新增)"
         Combo3.SetFocus
         bolGoOn = True
      Case 2
         Me.Caption = Me.Caption & "(修改)"
      Case 4
         Me.Caption = Me.Caption & "(刪除)"
      Case 5
         Me.Caption = Me.Caption & "(查詢)"
   End Select
   
   If intChoose <> 1 Then
      'edit by nickc 2007/02/05 不用 dll 了
      'If obj003.ReadIdTime(strTxt) Then
      If Cls003ReadIdTime(strTxt) Then
         Label3(0) = strTxt(12)
         Label3(2) = strTxt(15)
         Label3(1) = strTxt(13) & "  " & strTxt(14)
         Label3(3) = strTxt(16) & "  " & strTxt(17)
         Combo3.Text = strTxt(9)
      End If
      'edit by nickc 2007/02/05 不用 dll 了
      'If obj003.ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 3) Then
      'Modified by Morgan 2015/9/10 +擬制喪失新穎性
      If Cls003ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, IIf(m_CM10 <> "", Val(m_CM10), 3)) Then
         If intChoose = 5 Then
            cmdOK(0).Visible = False
         Else
            Combo3.SetFocus
         End If
         bolGoOn = True
      End If
   End If
   
   If bolGoOn Then
      If CheckCaseCode = False Then
         bolGoOn = False
      End If
   End If
   If bolGoOn = False Then
      intLeaveKind = 1
      Unload Me
   End If
End Sub

Private Function CheckCaseCode() As Boolean
   Dim strCodeName1 As String, strCodeName2 As String, strCodeName3 As String
   Dim strCustomer As String, strNation As String, strNationName As String
   Dim varSaveCursor
   
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtCode(0), txtCode(1), _
      IIf(txtCode(2) = "", "0", txtCode(2)), IIf(txtCode(3) = "", "00", txtCode(3)), strCodeName1, strCodeName2, strCodeName3, strCustomer, strNation) Then
   If ClsPDCheckCaseCodeIsExist(txtCode(0), txtCode(1), _
      IIf(txtCode(2) = "", "0", txtCode(2)), IIf(txtCode(3) = "", "00", txtCode(3)), strCodeName1, strCodeName2, strCodeName3, strCustomer, strNation) Then
      SetNameToCombo cboOut, strCodeName1, strCodeName2, strCodeName3
      lblApplicant = strCustomer
      ClsPDGetNation strNation, strNationName
      lblCountry = strNationName
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(txtCode(4), txtCode(5), _
         IIf(txtCode(6) = "", "0", txtCode(6)), IIf(txtCode(7) = "", "00", txtCode(7)), strCodeName1, strCodeName2, strCodeName3) Then
      If ClsPDCheckCaseCodeIsExist(txtCode(4), txtCode(5), _
         IIf(txtCode(6) = "", "0", txtCode(6)), IIf(txtCode(7) = "", "00", txtCode(7)), strCodeName1, strCodeName2, strCodeName3) Then
         SetNameToCombo cboIn, strCodeName1, strCodeName2, strCodeName3
         CheckCaseCode = True
      End If
   End If
   Screen.MousePointer = varSaveCursor
   
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If intLeaveKind = 1 Then
      Me.frmParent.Show
   Else
      Unload Me.frmParent
   End If
   Set frm040109_2 = Nothing
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = False Then
      '本所案號錯誤時,讓Cursor繼續往下跳
      If Index <> 3 And Index <> 7 Then
         Cancel = True
         txtCode_GotFocus Index
      End If
   End If
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Boolean
 Dim intCaseKind As Integer, intWhere As Integer, strTemp As String
   Select Case intIndex
      Case 0, 4
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetSystemKind(txtCode(intIndex), intCaseKind, , intWhere) Then
         If ClsPDGetSystemKind(txtCode(intIndex), intCaseKind, , intWhere) Then
            If intCaseKind = 專利 And (intWhere = 國內 Or intWhere = 國外_CF) Then
               CheckKeyIn = True
            Else
               ShowMsg MsgText(1056)
            End If
         End If
      Case 3, 7
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.CheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
              IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
         If ClsPDCheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
              IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
            If CheckCaseCode Then
               CheckKeyIn = True
            End If
         End If
      Case Else
         CheckKeyIn = True
   End Select
End Function

'新增國內外案件關聯表
Private Function bInsertCaseRelationData(ByRef strCode() As String, ByVal iSitu As Integer) As Boolean
'Added by Lydia 2018/06/27
Dim rsWD As ADODB.Recordset
Dim strCaseNo1 As String, strCaseNo2 As String

   cnnConnection.BeginTrans
   
On Error GoTo ErrHand

   strCode(2) = IIf(strCode(2) = "", "0", strCode(2))
   strCode(3) = IIf(strCode(3) = "", "00", strCode(3))
   strCode(6) = IIf(strCode(6) = "", "0", strCode(6))
   strCode(7) = IIf(strCode(7) = "", "00", strCode(7))
   strSql = "insert into casemap (cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,cm10,cm09) values (" + CNULL(strCode(0)) + "," + CNULL(strCode(1)) + "," + CNULL(strCode(2)) + "," + CNULL(strCode(3)) + "," + CNULL(strCode(4)) + "," + CNULL(strCode(5)) + "," + CNULL(strCode(6)) + "," + CNULL(strCode(7)) + ",'" & iSitu & "'," + CNULL(ChgSQL(strCode(8))) + ")"
   
   cnnConnection.Execute strSql, intI
   
   If strCode(0) <> "FCP" And iSitu = 3 Then 'Added by Morgan 2013/7/5
        'Add by Morgan 2010/1/28 一案兩請的新型案承辦人加乘註記要減0.5(未更改過且>=1的才要)
        'Modify by Morgan 2010/5/11 改*0.5
        strExc(0) = "select cp09,cp98 from (select pa01,pa02,pa03,pa04 from patent where pa01='" & strCode(0) & "' and pa02='" & strCode(1) & "'" & _
           " and pa03='" & strCode(2) & "' and pa04='" & strCode(3) & "' and pa08='2'" & _
           " union select pa01,pa02,pa03,pa04 from patent where pa01='" & strCode(4) & "' and pa02='" & strCode(5) & "'" & _
           " and pa03='" & strCode(6) & "' and pa04='" & strCode(7) & "' and pa08='2') X,caseprogress" & _
           " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='102' and cp98>=1" & _
           " and not exists(select * from flagstory where fs01=cp09 and fs04='1')"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           'Modify by Morgan 2010/5/11 改*0.5
           'strExc(1) = Val("" & RsTemp("cp98")) - 0.5
           strExc(1) = Round(Val("" & RsTemp("cp98")) * 0.5, 1)
           strSql = "update caseprogress set cp98=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
           cnnConnection.Execute strSql, intI
           'Modified by Morgan 2021/8/24 +加理由(@開頭為系統自動修改,目前分案會用來判斷有相同承辦的國內案是否要更新加乘註記)
           Call PUB_InsFlagStory(RsTemp("cp09"), "1", RsTemp("cp98"), strExc(1), "@一案兩請(原加乘註記:" & RsTemp("cp98") & ")")
        End If
        'end 2010/1/28
        
        'Add by Morgan 2010/2/5
        '一案兩請的新型案草圖加乘註記要減0.5(未更改過且>=1的才要)
        'Modify by Morgan 2010/5/11 改*0.5
        strExc(0) = "select cp09,cp101 from (select pa01,pa02,pa03,pa04 from patent where pa01='" & strCode(0) & "' and pa02='" & strCode(1) & "'" & _
           " and pa03='" & strCode(2) & "' and pa04='" & strCode(3) & "' and pa08='2'" & _
           " union select pa01,pa02,pa03,pa04 from patent where pa01='" & strCode(4) & "' and pa02='" & strCode(5) & "'" & _
           " and pa03='" & strCode(6) & "' and pa04='" & strCode(7) & "' and pa08='2') X,caseprogress" & _
           " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='102' and cp101>=1" & _
           " and not exists(select * from flagstory where fs01=cp09 and fs04='2')"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           'Modify by Morgan 2010/5/11 改*0.5
           'strExc(1) = Val("" & RsTemp("cp101")) - 0.5
           strExc(1) = Round(Val("" & RsTemp("cp101")) * 0.5, 1)
           strSql = "update caseprogress set cp101=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
           cnnConnection.Execute strSql, intI
           'Modified by Morgan 2021/8/24 +加理由
           Call PUB_InsFlagStory(RsTemp("cp09"), "2", RsTemp("cp101"), strExc(1), "@一案兩請(原加乘註記:" & RsTemp("cp101") & ")")
        End If
        '一案兩請的新型案墨圖加乘註記要減0.5(未更改過且>=1的才要)
        'Modify by Morgan 2010/5/11 改*0.5
        strExc(0) = "select cp09,cp104 from (select pa01,pa02,pa03,pa04 from patent where pa01='" & strCode(0) & "' and pa02='" & strCode(1) & "'" & _
           " and pa03='" & strCode(2) & "' and pa04='" & strCode(3) & "' and pa08='2'" & _
           " union select pa01,pa02,pa03,pa04 from patent where pa01='" & strCode(4) & "' and pa02='" & strCode(5) & "'" & _
           " and pa03='" & strCode(6) & "' and pa04='" & strCode(7) & "' and pa08='2') X,caseprogress" & _
           " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='102' and cp104>=1" & _
           " and not exists(select * from flagstory where fs01=cp09 and fs04='3')"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           'Modify by Morgan 2010/5/11 改*0.5
           'strExc(1) = Val("" & RsTemp("cp104")) - 0.5
           strExc(1) = Round(Val("" & RsTemp("cp104")) * 0.5, 1)
           strSql = "update caseprogress set cp104=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
           cnnConnection.Execute strSql, intI
           'Modified by Morgan 2021/8/24 +加理由
           Call PUB_InsFlagStory(RsTemp("cp09"), "3", RsTemp("cp104"), strExc(1), "@一案兩請(原加乘註記:" & RsTemp("cp104") & ")")
        End If
        'end 2010/2/5
        
        'Added by Morgan 2013/6/14 更新一案兩請繪圖人員(已輸入繪圖人員後才建關聯的情形)
        strExc(0) = "select cp107,cp29,cp10,1 c1  from caseprogress where cp01='" & strCode(0) & "' and cp02='" & strCode(1) & "'" & _
           " and cp03='" & strCode(2) & "' and cp04='" & strCode(3) & "' and cp10 in ('101','102') and cp29 is not null" & _
           " union select cp107,cp29,cp10,2 c1 from caseprogress where cp01='" & strCode(4) & "' and cp02='" & strCode(5) & "'" & _
           " and cp03='" & strCode(6) & "' and cp04='" & strCode(7) & "' and cp10 in ('101','102') and cp29 is not null order by cp107,cp10"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           If RsTemp("c1") = 1 Then
              'Modified by Morgan 2014/1/6 +案件性質為101或102否則其他的也會被設定(Ex.P-107193 的 416)
              strSql = "update caseprogress set cp29='" & RsTemp("cp29") & "' where cp01='" & strCode(4) & "' and cp02='" & strCode(5) & "'" & _
                 " and cp03='" & strCode(6) & "' and cp04='" & strCode(7) & "' and cp27 is null and cp10 in ('101','102')"
           Else
              'Modified by Morgan 2014/1/6 +案件性質為101或102否則其他的也會被設定(Ex.P-107193 的 416)
              strSql = "update caseprogress set cp29='" & RsTemp("cp29") & "' where cp01='" & strCode(0) & "' and cp02='" & strCode(1) & "'" & _
                 " and cp03='" & strCode(2) & "' and cp04='" & strCode(3) & "' and cp27 is null and cp10 in ('101','102')"
           End If
           cnnConnection.Execute strSql, intI
        End If
        'end 2013/6/14
   End If 'Added by Morgan 2013/7/5
   
   'Added by Lydia 2018/06/27 FCP一案兩請建立關聯後新型自動帶發明案之發明人、代表人和優先權資料
   If iSitu = 3 And strCode(0) = "FCP" And strCode(4) = "FCP" Then
        strExc(0) = "select pa01,pa02,pa03,pa04,pa79,pa80,pa81,pa82,pa83,pa84,pa109,pa110,pa111,pa112,pa113,pa114,pa115,pa116,pa117,pa118,pa119,pa120,pa121,pa122,pa123,pa124,pa125,pa126,pa127,pa128,pa129,pa130,pa131,pa132 " & _
                          "from patent where pa57 is null and pa08='1' and ((pa01='" & strCode(0) & "' and pa02='" & strCode(1) & "' and pa03='" & strCode(2) & "' and pa04='" & strCode(3) & "') or (pa01='" & strCode(4) & "' and pa02='" & strCode(5) & "' and pa03='" & strCode(6) & "' and pa04='" & strCode(7) & "')) "
        intI = 1
        Set rsWD = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            If rsWD.RecordCount = 1 Then
                   If "" & rsWD.Fields("pa01") & rsWD.Fields("pa02") & rsWD.Fields("pa03") & rsWD.Fields("pa04") = strCode(0) & strCode(1) & strCode(2) & strCode(3) Then
                         strCaseNo1 = strCode(0) & strCode(1) & strCode(2) & strCode(3)
                         strCaseNo2 = strCode(4) & strCode(5) & strCode(6) & strCode(7)
                   Else
                         strCaseNo1 = strCode(4) & strCode(5) & strCode(6) & strCode(7)
                         strCaseNo2 = strCode(0) & strCode(1) & strCode(2) & strCode(3)
                   End If
                   '更新新型案-個案代表人
                   strExc(2) = ""
                   For intI = 79 To 84
                        strExc(2) = strExc(2) & ", pa" & Format(intI, "00") & "=" & CNULL(ChgSQL("" & rsWD.Fields("pa" & Format(intI, "00"))))
                   Next intI
                   For intI = 109 To 132
                        strExc(2) = strExc(2) & ", pa" & Format(intI, "00") & "=" & CNULL(ChgSQL("" & rsWD.Fields("pa" & Format(intI, "00"))))
                   Next intI

                   strSql = " update patent set " & Mid(strExc(2), 2) & " where " & ChgPatent(strCaseNo2) & " and pa08='2' and pa57 is null"
                   cnnConnection.Execute strSql, intI
                   If intI > 0 Then
                       Pub_SeekTbLog strSql '維護log
                       Call ChgCaseNo(strCaseNo2, strExc)
                       
                       '更新新型案-優先權
                       strSql = "select * from pridate where " & Replace(ChgPatent(strCaseNo1), "PA", "PD")
                       intI = 1
                       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                       If intI = 1 Then
                           strSql = "delete from pridate where pd01='" & strExc(1) & "' and pd02='" & strExc(2) & "' and pd03='" & strExc(3) & "' and pd04='" & strExc(4) & "' "
                           cnnConnection.Execute strSql, intI
                           RsTemp.MoveFirst
                           Do While Not RsTemp.EOF
                                 strExc(5) = ", PD01, PD02, PD03, PD04"
                                 strExc(6) = ", " & CNULL(strExc(1)) & ", " & CNULL(strExc(2)) & ", " & CNULL(strExc(3)) & ", " & CNULL(strExc(4))
                                 For intI = 5 To RsTemp.Fields.Count
                                        strExc(5) = strExc(5) & ", PD" & Format(intI, "00")
                                        strExc(6) = strExc(6) & ", " & CNULL("" & RsTemp.Fields(intI - 1))
                                 Next
                                 strSql = "insert into PriDate (" & Mid(strExc(5), 2) & ") values (" & Mid(strExc(6), 2) & ") "
                                 cnnConnection.Execute strSql, intI
                                 RsTemp.MoveNext
                           Loop
                       End If

                       '更新專利案-發明人
                        strSql = "delete from PatentInventor where PI01='" & strExc(1) & "' and PI02='" & strExc(2) & "' and PI03='" & strExc(3) & "' and PI04='" & strExc(4) & "' "
                        cnnConnection.Execute strSql, intI
                        strSql = "insert into PatentInventor (PI01,PI02,PI03,PI04,PI05,PI06) select '" & strExc(1) & "', '" & strExc(2) & "', '" & strExc(3) & "', '" & strExc(4) & "', pi05,pi06 from PatentInventor where " & Replace(ChgPatent(strCaseNo1), "PA", "PI")
                        cnnConnection.Execute strSql, intI
                   End If
            End If
        End If
        Set rsWD = Nothing
   End If
   'end 2018/06/27
   
   'Added by Morgan 2019/9/10
   '一案兩請檢視中說設定不請款--敏莉
   If iSitu = 3 And (strCode(0) = "FCP" Or strCode(0) = "P") Then
      strSql = "update caseprogress set cp20='N' where cp01='" & strCode(0) & "' and cp02='" & strCode(1) & "' and cp03='" & strCode(2) & "' and cp04='" & strCode(3) & "' and cp10='209' and cp20||cp60 is null and cp12 like 'F%'" & _
         " and exists(select * from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and pa08='2')"
      cnnConnection.Execute strSql, intI
      
      strSql = "update caseprogress set cp20='N' where cp01='" & strCode(4) & "' and cp02='" & strCode(5) & "' and cp03='" & strCode(6) & "' and cp04='" & strCode(7) & "' and cp10='209' and cp20||cp60 is null and cp12 like 'F%'" & _
         " and exists(select * from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and pa08='2')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2019/9/10
   
   cnnConnection.CommitTrans
   bInsertCaseRelationData = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

'修改國內外案件關聯表
Private Function UpdateCaseRelationData(ByRef strCode() As String, ByVal iSitu As Integer) As Boolean
   
On Error GoTo ErrHand
   strCode(2) = IIf(strCode(2) = "", "0", strCode(2))
   strCode(3) = IIf(strCode(3) = "", "00", strCode(3))
   strCode(6) = IIf(strCode(6) = "", "0", strCode(6))
   strCode(7) = IIf(strCode(7) = "", "00", strCode(7))
   strSql = "begin user_data.user_enabled:=1; Update casemap set cm09=" & CNULL(ChgSQL(strCode(16))) & " where cm01=" + CNULL(strCode(0)) + " and cm02=" + CNULL(strCode(1)) + " and cm03=" + CNULL(strCode(2)) + " and cm04=" + CNULL(strCode(3)) + " and cm05=" + CNULL(strCode(4)) + " and cm06=" + CNULL(strCode(5)) + " and cm07=" + CNULL(strCode(6)) + " and cm08=" + CNULL(strCode(7)) + " and cm10='" & iSitu & "'; end;"
   
   cnnConnection.Execute strSql, intI
   UpdateCaseRelationData = True
   
ErrHand:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical

End Function

'Added by Morgan 2021/12/21 檢查畫面輸入欄位是否含有Unicode文字
Private Function TxtValidate() As Boolean
   TxtValidate = False
   If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
       Exit Function
   End If
   TxtValidate = True
End Function

