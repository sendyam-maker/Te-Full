VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040108_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸發明案件資料維護"
   ClientHeight    =   4440
   ClientLeft      =   735
   ClientTop       =   2130
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   3720
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   4548
      TabIndex        =   12
      Top             =   70
      Width           =   1200
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
   Begin MSForms.ComboBox Combo3 
      Height          =   300
      Left            =   1080
      TabIndex        =   10
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   6600
      Y1              =   2460
      Y2              =   2460
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
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   3
      Left            =   4560
      TabIndex        =   32
      Top             =   4080
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3651;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   2
      Left            =   1320
      TabIndex        =   31
      Top             =   4080
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3492;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   1
      Left            =   4560
      TabIndex        =   30
      Top             =   3840
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3651;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   0
      Left            =   1320
      TabIndex        =   29
      Top             =   3840
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3492;317"
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
      TabIndex        =   28
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Name:"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   27
      Top             =   4080
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Time:"
      Height          =   180
      Index           =   2
      Left            =   3480
      TabIndex        =   26
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Name:"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   3840
      Width           =   945
   End
   Begin MSForms.Label lblPromoterIn 
      Height          =   210
      Left            =   1080
      TabIndex        =   24
      Top             =   2640
      Width           =   1515
      VariousPropertyBits=   27
      Caption         =   "lblPromoterIn"
      Size            =   "2672;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSendDay 
      AutoSize        =   -1  'True
      Caption         =   "lblSendDay"
      Height          =   180
      Left            =   1080
      TabIndex        =   23
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "處理狀況:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "准駁日期:"
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CF案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "大陸案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frm040108_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/21 改成Form2.0 (cboOut,cboIn,Combo3,lblPromoterIn,Label3)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
'0從FRM040108_1來,1從FRM040108_3來

Public intWhereToGo As Integer
Public strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Public strCode5 As String, strCode6 As String, strCode7 As String, strCode8 As String
Public strCbo As String
Public intChoose As String

Private Sub cmdOK_Click(Index As Integer)
 Dim strCode() As String, i As Integer, bolSave As Boolean
   Select Case Index
      Case 0
         For i = 0 To 7
            If CheckKeyIn(i) = False Then
               '本所案號錯誤時,讓Cursor繼續往下跳
               If i <> 3 And i <> 7 Then
                  Me.txtCode(i).SetFocus
                  txtCode_GotFocus i
                  Exit Sub
               End If
            End If
         Next i
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         Select Case intChoose
            Case 1
               ReDim strCode(8) As String
               For i = 0 To 7
                  strCode(i) = txtCode(i)
               Next
               strCode(8) = Combo3.Text
               If CheckLengthIsOK(strCode(8), 20) Then
                    'Add By Cheng 2002/11/06
                    cnnConnection.BeginTrans
                    'edit by nickc 2007/02/05 不用 dll 了
                  'If obj003.InsertCaseRelationData(strCode(), 2) Then
                  If Cls003InsertCaseRelationData(strCode(), 2) Then
                         bolSave = True
                        'Add By Cheng 2002/11/06
                        cnnConnection.CommitTrans
                    Else
                        'Add By Cheng 2002/11/06
                        cnnConnection.RollbackTrans
                  End If
               Else
                  Combo3.SetFocus
                  bolSave = False
               End If
            Case 2
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
                    'Add By Cheng 2002/11/06
                    cnnConnection.BeginTrans
                     '910910 nick tigger
                     '***** start
                     'If obj003.UpdateCaseRelationData(strCode(), 2) Then
                     'edit by nickc 2007/02/05 不用 dll 了
                     'If obj003.UpdateCaseRelationData(strCode(), 2, True) Then
                     If Cls003UpdateCaseRelationData(strCode(), 2, True) Then
                     '***** end
                        bolSave = True
                        'Add By Cheng 2002/11/06
                        cnnConnection.CommitTrans
                     Else
                        'Add By Cheng 2002/11/06
                        cnnConnection.RollbackTrans
                     End If
                  End If
               Else
                  Combo3.SetFocus
                  bolSave = False
               End If
            Case 4
               ReDim strCode(7) As String
               For i = 0 To 7
                  strCode(i) = txtCode(i)
               Next
               'edit by nickc 2007/02/05 不用 dll 了
               'If obj003.ChkExist(strCode(), 2) Then
               If Cls003ChkExist(strCode(), 2) Then
                  If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                    'Add By Cheng 2002/11/06
                    cnnConnection.BeginTrans
                    'edit by nickc 2007/02/05 不用 dll 了
                     'If obj003.DeleteCaseRelation(strCode(), 2) Then
                     If Cls003DeleteCaseRelation(strCode(), 2) Then
                        bolSave = True
                        'Add By Cheng 2002/11/06
                        cnnConnection.CommitTrans
                    Else
                        'Add By Cheng 2002/11/06
                        cnnConnection.RollbackTrans
                     End If
                  End If
               End If
         End Select
         If bolSave Then
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
   For i = 1 To 8
      strTxt(i) = txtCode(i - 1)
   Next
   strTxt(10) = "2"
   'edit by nickc 2007/02/05 不用 dll 了
   'If obj003.ReadIdTime(strTxt) Then
   If Cls003ReadIdTime(strTxt) Then
      Label3(0) = strTxt(12)
      Label3(2) = strTxt(15)
      Label3(1) = strTxt(13) & "  " & strTxt(14)
      Label3(3) = strTxt(16) & "  " & strTxt(17)
      Combo3.Text = strTxt(9)
   End If
   
   If intChoose = 2 Or intChoose = 5 Then
      'edit by nickc 2007/02/05 不用 dll 了
      'If obj003.ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 2) Then
      If Cls003ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 2) Then
         If intChoose = 5 Then
            cmdOK(0).Visible = False
         Else
            Combo3.SetFocus
         End If
         bolGoOn = True
      End If
   Else
      bolGoOn = True
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
Dim varSaveCursor

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.CheckCaseCodeIsExist(txtCode(0), txtCode(1), _
      IIf(txtCode(2) = "", "0", txtCode(2)), IIf(txtCode(3) = "", "00", txtCode(3)), strCodeName1, strCodeName2, strCodeName3) Then
If ClsPDCheckCaseCodeIsExist(txtCode(0), txtCode(1), _
      IIf(txtCode(2) = "", "0", txtCode(2)), IIf(txtCode(3) = "", "00", txtCode(3)), strCodeName1, strCodeName2, strCodeName3) Then
   SetNameToCombo cboOut, strCodeName1, strCodeName2, strCodeName3
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtCode(4), txtCode(5), _
        IIf(txtCode(6) = "", "0", txtCode(6)), IIf(txtCode(7) = "", "00", txtCode(7)), strCodeName1, strCodeName2, strCodeName3) Then
   If ClsPDCheckCaseCodeIsExist(txtCode(4), txtCode(5), _
        IIf(txtCode(6) = "", "0", txtCode(6)), IIf(txtCode(7) = "", "00", txtCode(7)), strCodeName1, strCodeName2, strCodeName3) Then
      SetNameToCombo cboIn, strCodeName1, strCodeName2, strCodeName3
      'edit by nickc 2007/02/05 不用 dll 了
      'If obj003.GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, 2) Then
      '   If obj003.GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, 2) Then
      If Cls003GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, 2) Then
         If Cls003GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, 2) Then
            lblPromoterIn = strCodeName1
            lblSendDay = ChangeWStringToWDateString(strCodeName2)
            CheckCaseCode = True
         End If
      End If
   End If
End If
Screen.MousePointer = varSaveCursor
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If intLeaveKind = 1 Then
   If intWhereToGo = 0 Then
      frm040108_1.Show
   Else
      frm040108_3.Show
   End If
Else
  If intWhereToGo = 0 Then
     Unload frm040108_1
  Else
     Unload frm040108_3
  End If
End If
intLeaveKind = 0
'Add By Cheng 2002/07/18
Set frm040108_2 = Nothing
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
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

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/21 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
       Exit Function
   End If
   'end 2021/12/21
   
For Each objTxt In Me.txtCode
   If objTxt.Enabled = True Then
      Cancel = False
      txtCode_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.txtCode(objTxt.Index).SetFocus
         txtCode_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

