VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050106_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內外案件資料維護"
   ClientHeight    =   4065
   ClientLeft      =   435
   ClientTop       =   1635
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7695
   Begin VB.Frame fraOut 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   33
      Top             =   552
      Width           =   2412
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   3
         TabIndex        =   37
         Top             =   0
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   36
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   35
         Top             =   0
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   600
         MaxLength       =   6
         TabIndex        =   34
         Top             =   0
         Width           =   852
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   6756
      TabIndex        =   9
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4704
      TabIndex        =   7
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5532
      TabIndex        =   8
      Top             =   45
      Width           =   1200
   End
   Begin VB.Frame fraIn 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   1875
      Width           =   2412
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   0
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   7
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   4
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   6
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   3
         Top             =   0
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   5
         Left            =   600
         MaxLength       =   6
         TabIndex        =   2
         Top             =   0
         Width           =   852
      End
   End
   Begin MSForms.TextBox Text1 
      Height          =   330
      Left            =   1080
      TabIndex        =   6
      Top             =   3150
      Width           =   6495
      VariousPropertyBits=   -1467987941
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "11456;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboIn 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   2235
      Width           =   6495
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11456;529"
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
      TabIndex        =   0
      Top             =   915
      Width           =   6495
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11456;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Index           =   8
      Left            =   90
      TabIndex        =   44
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label lblSendDay1 
      Height          =   225
      Left            =   1050
      TabIndex        =   43
      Top             =   1530
      Width           =   1815
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      Caption         =   "lblPoint"
      Height          =   180
      Left            =   6615
      TabIndex        =   42
      Top             =   555
      Width           =   810
   End
   Begin VB.Label lblCountry 
      AutoSize        =   -1  'True
      Caption         =   "lblCountry"
      Height          =   180
      Left            =   4275
      TabIndex        =   41
      Top             =   555
      Width           =   1665
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   180
      Index           =   7
      Left            =   6030
      TabIndex        =   40
      Top             =   555
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "國家："
      Height          =   180
      Index           =   6
      Left            =   3645
      TabIndex        =   39
      Top             =   555
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國外案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   555
      Width           =   900
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Index           =   1
      Left            =   4515
      TabIndex        =   32
      Top             =   2595
      Width           =   1665
      VariousPropertyBits=   27
      Caption         =   "Label5"
      Size            =   "2937;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   5
      Left            =   3600
      TabIndex        =   31
      Top             =   2595
      Width           =   900
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Index           =   0
      Left            =   4515
      TabIndex        =   30
      Top             =   1275
      Width           =   1665
      VariousPropertyBits=   27
      Caption         =   "Label5"
      Size            =   "2937;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   4
      Left            =   3600
      TabIndex        =   29
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "記錄："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   3210
      Width           =   540
   End
   Begin MSForms.Label Label3 
      Height          =   210
      Index           =   3
      Left            =   4440
      TabIndex        =   27
      Top             =   3810
      Width           =   2160
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3810;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   210
      Index           =   2
      Left            =   1200
      TabIndex        =   26
      Top             =   3810
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3334;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   210
      Index           =   1
      Left            =   4440
      TabIndex        =   25
      Top             =   3570
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3757;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   210
      Index           =   0
      Left            =   1200
      TabIndex        =   24
      Top             =   3570
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3334;370"
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
      Left            =   3360
      TabIndex        =   23
      Top             =   3810
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Name:"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   3810
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Time:"
      Height          =   180
      Index           =   2
      Left            =   3360
      TabIndex        =   21
      Top             =   3570
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Name:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   3570
      Width           =   945
   End
   Begin VB.Label lblSendDay 
      Height          =   225
      Left            =   1080
      TabIndex        =   19
      Top             =   2850
      Width           =   1815
   End
   Begin VB.Label lblPromoterIn 
      AutoSize        =   -1  'True
      Caption         =   "lblPromoterIn"
      Height          =   210
      Left            =   1080
      TabIndex        =   18
      Top             =   2595
      Width           =   1800
   End
   Begin VB.Label lblPromoterOut 
      AutoSize        =   -1  'True
      Caption         =   "lblPromoterOut"
      Height          =   210
      Left            =   1080
      TabIndex        =   17
      Top             =   1275
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   915
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1275
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國內案號："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1875
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2235
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2595
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2850
      Width           =   720
   End
End
Attribute VB_Name = "frm050106_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Text1,cboOut,cboIn,Label5,Label3,...)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
'0從frm050106_1來,1從frm050106_3來
Public intWhereToGo As Integer
Public strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Public strCode5 As String, strCode6 As String, strCode7 As String, strCode8 As String
Public strCode18 As String
Public intChoose As String
Dim m_CP09 As String, m_CP10 As String, m_PA09 As String 'Add by Morgan 2005/3/15

Private Sub cmdOK_Click(Index As Integer)
 Dim strCode() As String, i As Integer, bolSave As Boolean
   Select Case Index
      Case 0
         'Add By Cheng 2002/05/17
         '檢查資料輸入的完整性
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
               ReDim strCode(7) As String
               For i = 0 To 7
                      strCode(i) = txtCode(i)
               Next
               'Add by Morgan 2005/5/3 檢查一件國內案不可關聯兩件相同國家之國外案
               If PUB_CheckDaulCaseMap(strCode(0) & strCode(1) & strCode(2) & strCode(3), strCode(4) & strCode(5) & strCode(6) & strCode(7)) = True Then
                  Exit Sub
               End If
               '2005/5/3 end
               '911105 nick transation
               cnnConnection.BeginTrans
               'Modify by Morgan 2004/3/15
               '加 113, 114, 307案件性質
               'If obj003.InsertCaseRelationData(strCode(), 0) Then
               If Cls003InsertCaseRelationData(strCode(), 0) Then
                  '911105 nick transation
                  'Modify by Morgan 2005/3/10 更新進度檔計件值
                  'cnnConnection.CommitTrans
                  'bolSave = True

                  
'Add by Morgan 2005/3/15 若國外案為CFP或大陸設計且無繪圖人員時帶國內案繪圖人員
On Error GoTo ErrHnd:
                  '國內案號
                  strExc(1) = txtCode(4)
                  strExc(2) = txtCode(5)
                  strExc(3) = txtCode(6)
                  strExc(4) = txtCode(7)
                  If (txtCode(0) = "CFP") Or (m_PA09 = "020" And m_CP10 = "103") Then
                     Call PUB_UpdateEP13(m_CP09, strExc())
                  End If
'2005/3/15 end
                  '國外案號
                  strExc(1) = txtCode(0)
                  strExc(2) = txtCode(1)
                  strExc(3) = txtCode(2)
                  strExc(4) = txtCode(3)
                  'Remove by Morgan 2005/4/13 改由 trigger 更新
                  'If PUB_UpdateCaseValueA(strExc()) = False Then
                  '   cnnConnection.RollbackTrans
                  'Else
                     cnnConnection.CommitTrans
                     bolSave = True
                  'End If
                  '2005/3/10 end
                  
'Add by Morgan 2005/3/15
ErrHnd:
                  If Err.NUMBER <> 0 Then
                     cnnConnection.RollbackTrans
                     MsgBox Err.Description, vbCritical
                  End If
'2005/3/15 end
               Else
                  '911105 nick transation
                  cnnConnection.RollbackTrans
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
               strCode(16) = Text1
               If CheckCaseCode Then
                  '910910 nick tigger
                  '***** start
                  'If obj003.UpdateCaseRelationData(strCode(), 0) Then
                  '911105 nick transation
                  cnnConnection.BeginTrans
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If obj003.UpdateCaseRelationData(strCode(), 0, True) Then
                  If Cls003UpdateCaseRelationData(strCode(), 0, True) Then
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
               'If obj003.ChkExist(strCode(), 0) Then
               If Cls003ChkExist(strCode(), 0) Then
                  If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                     '911105 nick transation
                     cnnConnection.BeginTrans
                     'edit by nickc 2007/02/05 不用 dll 了
                     'If obj003.DeleteCaseRelation(strCode(), 0) Then
                     If Cls003DeleteCaseRelation(strCode(), 0) Then
                        '911105 nick transation
                        'Modify by Morgan 2005/3/10 更新進度檔計件值
                        'cnnConnection.CommitTrans
                        'bolSave = True
                        strExc(1) = txtCode(0)
                        strExc(2) = txtCode(1)
                        strExc(3) = txtCode(2)
                        strExc(4) = txtCode(3)
                        'Remove by Morgan 2005/4/13 改由 trigger 更新
                        'If PUB_UpdateCaseValueA(strExc()) = False Then
                        '   cnnConnection.RollbackTrans
                        'Else
                           cnnConnection.CommitTrans
                           bolSave = True
                        'End If
                        '2005/3/10 end
                     Else
                        '911105 nick transation
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
   Text1 = ""
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   For i = 1 To 8
      strTxt(i) = txtCode(i - 1)
   Next
   strTxt(10) = "0"
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
      'If obj003.ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 0) Then
      If Cls003ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 0) Then
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
      'Add By Cheng 2002/12/09
      Exit Sub
   End If
   
   '910802  Sieg 503
   strExc(0) = "select cm18 from casemap where cm01='" & txtCode(0) & "' and cm02='" & txtCode(1) & "' and cm03='" & txtCode(2) & "' and cm04='" & txtCode(3) & "' and cm05='" & txtCode(4) & "' and cm06='" & txtCode(5) & "' and cm07='" & txtCode(6) & "' and cm08='" & txtCode(7) & "' and cm10='0'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 And Not IsNull(RsTemp.Fields(0)) Then
         Text1 = RsTemp.Fields(0)
   End If
   
   If intChoose = 1 Or intChoose = 2 Then
      Text1.Locked = False
   Else
      Text1.Locked = True
   End If
End Sub

Private Function CheckCaseCode() As Boolean
Dim strCodeName1 As String, strCodeName2 As String, strCodeName3 As String
Dim varSaveCursor
Dim stCountry As String, stPoint As String 'Add by Morgan 2004/10/29

lblPromoterOut.Caption = ""
Label5(0).Caption = ""
Label5(1).Caption = ""
lblPromoterIn.Caption = ""

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
        
      'Modify by Morgan 2004/3/15
      '加 113, 114, 307案件性質
      'If obj003.GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, 0, strExc(0)) Then
      'Modify by Morgan 2004/12/27 加發文日
      'Modify by Morgan 2005/3/15 加收文號，案件性質，申請國家代碼
      m_CP09 = "": m_CP10 = "": m_PA09 = "" 'Add by Morgan 2005/3/15
      If GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, 0, strExc(0), stCountry, stPoint, strCodeName2, m_CP09, m_CP10, m_PA09) Then
         lblPromoterOut = strCodeName1
         Label5(0) = strExc(0)
         lblCountry = stCountry
         lblPoint = stPoint
         
         'Add by Morgan 2004/12/27
         lblSendDay1 = ChangeWStringToWDateString(strCodeName2)
         strCodeName2 = ""
         '2004/12/27
         
         'Modify by Morgan 2004/3/15
         '加 113, 114, 307案件性質
         'If obj003.GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, 0, strExc(0)) Then
         If GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, 0, strExc(0)) Then
            lblPromoterIn = strCodeName1
            lblSendDay = ChangeWStringToWDateString(strCodeName2)
            Label5(1) = strExc(0)
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
      frm050106_1.Show
   Else
      frm050106_3.Show
   End If
Else
  If intWhereToGo = 0 Then
     Unload frm050106_1
  Else
     Unload frm050106_3
  End If
End If
intLeaveKind = 0
Set frm050106_2 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   'edit by nickc 2007/06/06 切換輸入法改用API
   'Text1.IMEMode = 1
   OpenIme
End Sub

Private Sub Text1_LostFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'Text1.IMEMode = 2
   CloseIme
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
            'Modify by Morgan 2005/12/6 加國內案可輸FCP
            'If intCaseKind = 專利 And (intWhere = 國內 Or intWhere = 國外_CF) Then
            If intCaseKind = 專利 And (intWhere = 國內 Or intWhere = 國外_CF Or intWhere = 國外_FC) Then
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

   'Added by Morgan 2021/12/20 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/20
   
For Each objTxt In txtCode
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

