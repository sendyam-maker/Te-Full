VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050109_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸香港案件資料維護"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7650
   Begin VB.Frame fraIn 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   315
      Left            =   1065
      TabIndex        =   15
      Top             =   1962
      Width           =   2412
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   5
         Left            =   600
         MaxLength       =   6
         TabIndex        =   2
         Top             =   0
         Width           =   852
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
         Index           =   7
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   4
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   0
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Width           =   492
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5520
      TabIndex        =   8
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4695
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6735
      TabIndex        =   9
      Top             =   60
      Width           =   800
   End
   Begin VB.Frame fraOut 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   315
      Left            =   1065
      TabIndex        =   10
      Top             =   570
      Width           =   2412
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   600
         MaxLength       =   6
         TabIndex        =   14
         Top             =   0
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   13
         Top             =   0
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   12
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   3
         TabIndex        =   11
         Top             =   0
         Width           =   492
      End
   End
   Begin MSForms.TextBox Text1 
      Height          =   510
      Left            =   1050
      TabIndex        =   6
      Top             =   3330
      Width           =   6495
      VariousPropertyBits=   -1467989989
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "11456;900"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboIn 
      Height          =   330
      Left            =   1050
      TabIndex        =   5
      Top             =   2310
      Width           =   6495
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "11456;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboOut 
      Height          =   330
      Left            =   1050
      TabIndex        =   0
      Top             =   918
      Width           =   6495
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "11456;582"
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
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   44
      Top             =   3006
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   255
      Index           =   2
      Left            =   75
      TabIndex        =   43
      Top             =   2685
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   42
      Top             =   2310
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "大陸案號："
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   41
      Top             =   1962
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   40
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   39
      Top             =   918
      Width           =   900
   End
   Begin MSForms.Label lblPromoterOut 
      Height          =   255
      Left            =   1050
      TabIndex        =   38
      Top             =   1290
      Width           =   1260
      VariousPropertyBits=   27
      Caption         =   "lblPromoterOut"
      Size            =   "2222;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPromoterIn 
      Height          =   255
      Left            =   1050
      TabIndex        =   37
      Top             =   2685
      Width           =   1260
      VariousPropertyBits=   27
      Caption         =   "lblPromoterIn"
      Size            =   "2222;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSendDay 
      Height          =   255
      Left            =   1050
      TabIndex        =   36
      Top             =   3006
      Width           =   1815
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Name:"
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   35
      Top             =   3915
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Time:"
      Height          =   180
      Index           =   2
      Left            =   3525
      TabIndex        =   34
      Top             =   3915
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Name:"
      Height          =   255
      Index           =   3
      Left            =   75
      TabIndex        =   33
      Top             =   4185
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Time:"
      Height          =   180
      Index           =   4
      Left            =   3525
      TabIndex        =   32
      Top             =   4185
      Width           =   945
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   31
      Top             =   3915
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   4515
      TabIndex        =   30
      Top             =   3915
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   29
      Top             =   4185
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   3
      Left            =   4515
      TabIndex        =   28
      Top             =   4185
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "記錄："
      Height          =   255
      Index           =   3
      Left            =   75
      TabIndex        =   27
      Top             =   3360
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   255
      Index           =   4
      Left            =   3585
      TabIndex        =   26
      Top             =   1290
      Width           =   900
   End
   Begin MSForms.Label Label5 
      Height          =   255
      Index           =   0
      Left            =   4515
      TabIndex        =   25
      Top             =   1290
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label5"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   255
      Index           =   5
      Left            =   3585
      TabIndex        =   24
      Top             =   2685
      Width           =   900
   End
   Begin MSForms.Label Label5 
      Height          =   255
      Index           =   1
      Left            =   4515
      TabIndex        =   23
      Top             =   2685
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label5"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "香港案號："
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   22
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "國家："
      Height          =   255
      Index           =   6
      Left            =   3630
      TabIndex        =   21
      Top             =   570
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   255
      Index           =   7
      Left            =   6015
      TabIndex        =   20
      Top             =   570
      Width           =   540
   End
   Begin VB.Label lblCountry 
      AutoSize        =   -1  'True
      Caption         =   "lblCountry"
      Height          =   255
      Left            =   4260
      TabIndex        =   19
      Top             =   570
      Width           =   765
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      Caption         =   "lblPoint"
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   570
      Width           =   810
   End
   Begin VB.Label lblSendDay1 
      Height          =   255
      Left            =   1050
      TabIndex        =   17
      Top             =   1614
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   255
      Index           =   8
      Left            =   75
      TabIndex        =   16
      Top             =   1614
      Width           =   720
   End
End
Attribute VB_Name = "frm050109_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; cboOut、cboIn、Label3(index)、Label5(index)、Text1、lblPromoterOut、lblPromoterIn
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
'0從frm050109_1來,1從frm050109_3來
Public intWhereToGo As Integer
Public strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Public strCode5 As String, strCode6 As String, strCode7 As String, strCode8 As String
Public strCode18 As String
Public intChoose As String
Dim m_CP09 As String, m_CP10 As String, m_PA09 As String 'Add by Morgan 2005/3/15
'Added by Lydia 2015/07/27 +大陸澳門案(共用表單frm050109_1,frm050109_2,frm050109_3)
Public iK_CM10 As Integer  '判斷案件類別
Dim m_NA03 As String  '案件-國別
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
               cnnConnection.BeginTrans
               'Modified by Lydia 2015/07/27
               'If InsertCaseRelationData(strCode(), 4) Then
               If InsertCaseRelationData(strCode(), iK_CM10) Then
                     cnnConnection.CommitTrans
                     bolSave = True
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
                  '911105 nick transation
                  cnnConnection.BeginTrans
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If obj003.UpdateCaseRelationData(strCode(), 4, True) Then
                  'Modified by Lydia 2015/07/27
                  'If Cls003UpdateCaseRelationData(strCode(), 4, True) Then
                  If Cls003UpdateCaseRelationData(strCode(), iK_CM10, True) Then
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
               'If obj003.ChkExist(strCode(), 4) Then
               'Modified by Lydia 2015/07/27
               'If Cls003ChkExist(strCode(), 4) Then
               If Cls003ChkExist(strCode(), iK_CM10) Then
                  If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                     cnnConnection.BeginTrans
                     'edit by nickc 2007/02/05 不用 dll 了
                     'If obj003.DeleteCaseRelation(strCode(), 4) Then
                     'Modified by Lydia 2015/07/27
                     'If Cls003DeleteCaseRelation(strCode(), 4) Then
                     If Cls003DeleteCaseRelation(strCode(), iK_CM10) Then
                        strExc(1) = txtCode(0)
                        strExc(2) = txtCode(1)
                        strExc(3) = txtCode(2)
                        strExc(4) = txtCode(3)
                        cnnConnection.CommitTrans
                        bolSave = True
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
    'Added by Lydia 2015/07/27 +澳門大陸案
    Select Case iK_CM10
        Case 4:  m_NA03 = "香港"
        Case 5:  m_NA03 = "澳門"
    End Select
    Me.Caption = "大陸" & m_NA03 & "案件資料維護"
    Label1(0).Caption = m_NA03 & "案號："
    'end 2015/07/27

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
      'If obj003.ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, 4) Then
      If Cls003ReadCaseRelationData(strCode1, strCode2, strCode3, strCode4, strCode5, strCode6, strCode7, strCode8, iK_CM10) Then
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
   'Modified by Lydia 2015/07/27
   'strExc(0) = "select cm18 from casemap where cm01='" & txtCode(0) & "' and cm02='" & txtCode(1) & "' and cm03='" & txtCode(2) & "' and cm04='" & txtCode(3) & "' and cm05='" & txtCode(4) & "' and cm06='" & txtCode(5) & "' and cm07='" & txtCode(6) & "' and cm08='" & txtCode(7) & "' and cm10='4'"
   strExc(0) = "select cm18 from casemap where cm01='" & txtCode(0) & "' and cm02='" & txtCode(1) & "' and cm03='" & txtCode(2) & "' and cm04='" & txtCode(3) & "' and cm05='" & txtCode(4) & "' and cm06='" & txtCode(5) & "' and cm07='" & txtCode(6) & "' and cm08='" & txtCode(7) & "' and cm10='" & iK_CM10 & "'"
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
      'Modified by Lydia 2015/07/27
      'If GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, 4, strExc(0), stCountry, stPoint, strCodeName2, m_CP09, m_CP10, m_PA09) Then
      If GetCaseRelationDataOut(txtCode(0), txtCode(1), txtCode(2), txtCode(3), strCodeName1, iK_CM10, strExc(0), stCountry, stPoint, strCodeName2, m_CP09, m_CP10, m_PA09) Then
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
         'Modified by Lydia 2015/07/27
         'If GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, 4, strExc(0)) Then
         If GetCaseRelationDataIn(txtCode(4), txtCode(5), txtCode(6), txtCode(7), strCodeName1, strCodeName2, iK_CM10, strExc(0)) Then
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
      frm050109_1.Show
   Else
      frm050109_3.iK_CM10 = iK_CM10 'Added by Lydia 2015/07/27
      frm050109_3.Show
   End If
Else
  If intWhereToGo = 0 Then
     Unload frm050109_1
  Else
     Unload frm050109_3
  End If
End If
intLeaveKind = 0
Set frm050109_2 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text1.IMEMode = 1
   OpenIme
End Sub

Private Sub Text1_LostFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
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

'Added by Lydia 2022/02/18 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function



