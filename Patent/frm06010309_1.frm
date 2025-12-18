VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010309_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-詢進度"
   ClientHeight    =   3790
   ClientLeft      =   -290
   ClientTop       =   2720
   ClientWidth     =   7940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3790
   ScaleWidth      =   7940
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   285
      Left            =   1080
      TabIndex        =   33
      Top             =   1470
      Visible         =   0   'False
      Width           =   1875
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   900
         Style           =   1  '圖片外觀
         TabIndex        =   34
         Top             =   -30
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   35
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.TextBox txtCP84 
      Height          =   300
      Left            =   4050
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2550
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6900
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5748
      TabIndex        =   5
      Top             =   70
      Width           =   1110
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2940
      Width           =   375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2580
      Width           =   1032
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010309_1.frx":0000
      Left            =   1080
      List            =   "frm06010309_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   1110
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   10
      Top             =   480
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1632
      MaxLength       =   6
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2472
      MaxLength       =   1
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2712
      MaxLength       =   2
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3300
      Width           =   300
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額:"
      Height          =   180
      Left            =   3150
      TabIndex        =   32
      Top             =   2610
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   7680
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   7680
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請書類別:              (1.爭議案審查確定 2.年費繳交狀況 3.改變原處分)"
      Height          =   180
      Left            =   240
      TabIndex        =   31
      Top             =   3000
      Width           =   5475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   2610
      Width           =   945
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3240
      TabIndex        =   29
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3240
      TabIndex        =   28
      Top             =   2130
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   2130
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   26
      Top             =   540
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3237
      TabIndex        =   25
      Top             =   1800
      Width           =   768
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   240
      TabIndex        =   24
      Top             =   1800
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   23
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   22
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3240
      TabIndex        =   21
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   20
      Top             =   1140
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   19
      Top             =   810
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   18
      Top             =   810
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1740
      TabIndex        =   17
      Top             =   1140
      Width           =   6090
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "10742;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   16
      Top             =   1800
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4080
      TabIndex        =   15
      Top             =   1800
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1230
      TabIndex        =   14
      Top             =   2130
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   13
      Top             =   2130
      Width           =   3480
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6138;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容            (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   3345
      Width           =   2985
   End
End
Attribute VB_Name = "frm06010309_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/8 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Public strReceiveNo As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, CP10 As String
Dim pa() As String, cp() As String, CP10 As String
Dim intWhere As Integer
Public m_CP118isY As Boolean 'Add By Sindy 2019/1/2 是否為電子送件申請書:True.是
Dim m_CaseNo As String 'Add By Sindy 2019/1/2


Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim strFolder As String, strFileName As String 'Add By Sindy 2019/1/2
   
   Select Case Index
      Case 0
         'Add by Sindy 2019/1/2
         'If TxtValidate = False Then Exit Sub
         
         'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
         If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
             MsgBox MsgText(1111), vbInformation
             If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                  Exit Sub
             End If
         End If
         'end 2020/02/17
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Add by Sindy 2019/1/2 電子送件申請書
         If m_CP118isY = True Then
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            'If Pub_StrUserSt03 = "M51" Then
            If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
               strFolder = PUB_Getdesktop
            Else
               strFolder = FCP電子送件檔案存放路徑
            End If
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            '1.基本資料
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
            '2.申請書
            If StartLetter2("01", "00") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "00", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & "查詢案件進度申請書"
            Call PUB_MakeDoc(strExc(9), strFileName)
         Else
         '2019/1/2 END
            If Text7 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            Select Case Text6.Text
               Case "1"
                  If CP10 = 舉發 Or CP10 = 異議_專 Then
                     '舉發成立 01
                     strTmp = "01"
                  ElseIf CP10 = 被舉發理由 Or CP10 = 被異議理由 Then
                     '被舉發不成立 02
                     strTmp = "02"
                  Else
                     MsgBox "案件性質只有異議,舉發或被異議理由,被舉發理由時才有申請書 !", vbCritical
                     Exit Sub
                  End If
               Case "2"
                  '年費繳交狀況 03
                  strTmp = "03"
               Case "3"
                  '改變原處份 04
                  strTmp = "04"
            End Select
            strLetterDate = Text5.Text
            NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, 0
         End If
         
         frm060103_1.Show
         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
         frm060103_1.ClearForm
      Case 1
         frm060103_1.Show
      Case 2
         Unload frm060103_1
   End Select
   Unload Me
End Sub

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP07Add2M As String
Dim strPA10Add4M As String, strPA10Add6M As String
Dim strCP27Add4M As String, strCP27Add6M As String
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())

   '出名代理人
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
'   '辦理依據
'   If Text10 <> "" Then
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(Text10) & "')"
'
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & Text11 & "')"
'
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & Text12 & "')"
'   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label12(3) = pa(5)
      Case "英"
         Label12(3) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label12(3) = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim cp(TF_CP) 'Add By Sindy 2019/1/2
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   ReadPatent
   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010309_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
   For Each Lbl In Label12
      Lbl = ""
   Next
   
   'Add By Sindy 2019/1/2
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If Val(cp(84)) > 0 Then
         txtCP84 = cp(84) '發文規費
      ElseIf Val(cp(17)) > 0 Then
         txtCP84 = cp(17) '規費
      Else
         txtCP84 = 0
      End If
   End If
   '2019/1/2 END
   
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Text5 = pa(10)
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      Label12(3) = pa(5)
   End If
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,CP10 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      If Not IsNull(.Fields(0)) Then Label12(0) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2)
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields(1)
         End If
      End If
      'Add By Cheng 2002/07/17
      CP10 = ""
      If Not IsNull(.Fields(4)) Then CP10 = .Fields(4)
   End If
   End With
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
End Sub

'Add By Sindy 2019/1/2
Private Function FormSave() As Boolean
Dim strCon As String
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   If m_CP118isY = True Then
      cp(118) = "A"
   Else
      cp(118) = ""
   End If
   strCon = strCon & ",cp84=" & Val(txtCP84) '發文規費
   strSql = " UPDATE CASEPROGRESS SET cp118=" & CNULL(cp(118)) & strCon & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
   End If
End Function

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 49 Or KeyAscii > 51) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
