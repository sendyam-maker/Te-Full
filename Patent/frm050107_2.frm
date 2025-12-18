VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050107_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "美國IDS資料對照維護"
   ClientHeight    =   3900
   ClientLeft      =   -2325
   ClientTop       =   2430
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   6720
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4668
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5496
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.Frame fraIn 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   1920
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
      Height          =   375
      Left            =   1440
      TabIndex        =   19
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
   Begin MSForms.ComboBox cboIn 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   6135
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "10821;529"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   6135
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "10821;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   32
      Top             =   3600
      Width           =   2670
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "4710;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   31
      Top             =   3600
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3598;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   30
      Top             =   3300
      Width           =   2670
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "4710;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   29
      Top             =   3300
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3598;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Time:"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   28
      Top             =   3600
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Name:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Time:"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   26
      Top             =   3300
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   3300
      Width           =   945
   End
   Begin VB.Label Label4 
      Caption         =   "承辦人："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   855
   End
   Begin MSForms.Label lblPromoterIn 
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   2640
      Width           =   6495
      VariousPropertyBits=   27
      Size            =   "11456;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSendDay 
      Height          =   255
      Left            =   1290
      TabIndex        =   22
      Top             =   2970
      Width           =   1815
   End
   Begin MSForms.Label lblPromoterOut 
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   1440
      Width           =   6495
      VariousPropertyBits=   27
      Size            =   "11456;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "美國案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "承辦人："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "EPC,英國案號："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "IDS收文日："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2970
      Width           =   1095
   End
End
Attribute VB_Name = "frm050107_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (cboOut,cboIn,lblPromoterOut,lblPromoterIn,Label3)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
'0從frm050107_1來,1從frm050107_3來
Public intWhereToGo As Integer
Public strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Public strCode5 As String, strCode6 As String, strCode7 As String, strCode8 As String
Public intChoose As String

Private Sub cmdOK_Click(Index As Integer)
 Dim strCode() As String, i As Integer, bolSave As Boolean
   Select Case Index
      Case 0
         Select Case intChoose
             Case 1
                  ReDim strCode(7) As String
                  For i = 0 To 7
                         strCode(i) = txtCode(i)
                  Next
                  'Add By Cheng 2002/05/22
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Sub
                  
                  '911105 nick transation
                  cnnConnection.BeginTrans
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If obj003.InsertCaseRelationData(strCode(), 1) Then
                  If Cls003InsertCaseRelationData(strCode(), 1) Then
                     '911105 nick transation
                     cnnConnection.CommitTrans
                     
                     bolSave = True
                     intWhereToGo = 0
                  Else
                  '911105 nick transation
                     cnnConnection.RollbackTrans
                  End If
             Case 2
                  ReDim strCode(15) As String
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
                  If CheckCaseCode Then
                     'Add By Cheng 2002/05/22
                     '重新檢查欄位有效性
                     If TxtValidate = False Then Exit Sub
                     '910910 nick tigger
                     '***** start
                     'If obj003.UpdateCaseRelationData(strCode(), 1) Then
                     '911105 nick  transation
                     cnnConnection.BeginTrans
                     'edit by nickc 2007/02/05 不用 dll 了
                     'If obj003.UpdateCaseRelationData(strCode(), 1, True) Then
                     If Cls003UpdateCaseRelationData(strCode(), 1, True) Then
                        '911105 nick transation
                        cnnConnection.CommitTrans
                     '***** end
                        bolSave = True
                     Else
                        '911105 nick transation
                        cnnConnection.RollbackTrans
                     End If
                  End If
            Case 4
               ReDim strCode(7) As String
               For i = 0 To 7
                  strCode(i) = txtCode(i)
               Next
               'edit by nickc 2007/02/05 不用 dll 了
               'If obj003.ChkExist(strCode(), 1) Then
               If Cls003ChkExist(strCode(), 1) Then
                  If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                     '911105 nick transation
                     cnnConnection.BeginTrans
                     'edit by nickc 2007/02/05 不用 dll 了
                     'If obj003.DeleteCaseRelation(strCode(), 1) Then
                     If Cls003DeleteCaseRelation(strCode(), 1) Then
                        '911105 nick transation
                        cnnConnection.CommitTrans
                        bolSave = True
                     Else
                        '911105 nick
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

   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   For i = 1 To 8
      strTxt(i) = txtCode(i - 1)
   Next
   strTxt(10) = "1"
   'edit by nickc 2007/02/05 不用 dll 了
   'If obj003.ReadIdTime(strTxt) Then
   If Cls003ReadIdTime(strTxt) Then
      Label3(0) = strTxt(12)
      Label3(2) = strTxt(15)
      Label3(1) = strTxt(13) & "  " & strTxt(14)
      Label3(3) = strTxt(16) & "  " & strTxt(17)
   End If

   If intChoose = 2 Or intChoose = 5 Then
      'edit by nickc 2007/02/05 不用 dll 了
      'If obj003.ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 1) Then
      If Cls003ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 1) Then
         If intChoose = 5 Then
            cmdOK(0).Visible = False
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
Dim bOldShowMsg As Boolean

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
      'Modify by Morgan 2004/10/27 不在用dll的
      'If obj003.GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, 1) Then
      If GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, 1) Then
         lblPromoterOut = strCodeName1
         ' 90.07.03 modify by louis
         'edit by nickc 2007/02/05 不用 dll 了
         'bOldShowMsg = obj003.EnableShowMessage(False)
         'If obj003.GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, 1) Then
         bOldShowMsg = Cls003EnableShowMessage(False)
         If Cls003GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, 1) Then
            lblPromoterIn = strCodeName1
            lblSendDay = ChangeWStringToWDateString(strCodeName2)
            CheckCaseCode = True
         ' 90.07.03 不管結果, 直接可進入此畫面
         Else
            lblPromoterIn = strCodeName1
            lblSendDay = ChangeWStringToWDateString(strCodeName2)
            CheckCaseCode = True
         End If
         'edit by nickc 2007/02/05 不用 dll 了
         'obj003.EnableShowMessage (bOldShowMsg)
         Cls003EnableShowMessage (bOldShowMsg)
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
      frm050107_1.Show
   Else
      frm050107_3.Show
   End If
Else
  If intWhereToGo = 0 Then
     Unload frm050107_1
  Else
     Unload frm050107_3
  End If
End If
intLeaveKind = 0
'Add By Cheng 2002/07/18
Set frm050107_2 = Nothing
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
                           If intCaseKind = 專利 And intWhere = 國外_CF Then
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
For Each objTxt In Me.txtCode
   If objTxt.Enabled = True Then
      Cancel = False
      txtCode_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

