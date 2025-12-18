VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010310_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-其他案件性質"
   ClientHeight    =   4680
   ClientLeft      =   130
   ClientTop       =   2500
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7950
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4848
      MaxLength       =   7
      TabIndex        =   52
      Top             =   3225
      Width           =   456
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   945
      Left            =   3690
      TabIndex        =   49
      Top             =   3570
      Width           =   2385
      Begin VB.CheckBox chk412 
         Enabled         =   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   45
         Value           =   1  '核取
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtCP71 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   11
         Top             =   270
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblCP71 
         AutoSize        =   -1  'True
         Caption         =   "延緩公告："
         Height          =   180
         Left            =   270
         TabIndex        =   51
         Top             =   45
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "延緩月數/日期"
         Height          =   180
         Left            =   30
         TabIndex        =   50
         Top             =   315
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.TextBox TextPA178 
      Height          =   270
      Left            =   1290
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3225
      Width           =   375
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2910
      Width           =   990
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   945
      Left            =   240
      TabIndex        =   40
      Top             =   3570
      Width           =   3105
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   1
         Left            =   1830
         MaxLength       =   2
         TabIndex        =   7
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   990
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "1"
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1410
         MaxLength       =   1
         TabIndex        =   8
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   1290
         MaxLength       =   1
         TabIndex        =   9
         Top             =   630
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "年 年費"
         Height          =   180
         Left            =   2430
         TabIndex        =   46
         Top             =   75
         Width           =   585
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   1590
         TabIndex        =   45
         Top             =   75
         Width           =   180
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "繳納第:"
         Height          =   180
         Left            =   30
         TabIndex        =   44
         Top             =   75
         Width           =   585
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "(Y:雙倍)"
         Height          =   180
         Left            =   1890
         TabIndex        =   43
         Top             =   375
         Width           =   645
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "費用是否要雙倍:"
         Height          =   180
         Left            =   30
         TabIndex        =   42
         Top             =   375
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "是否逾期補繳:              (Y:是)"
         Height          =   180
         Left            =   30
         TabIndex        =   41
         Top             =   675
         Visible         =   0   'False
         Width           =   2220
      End
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5115
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2580
      Width           =   555
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   6948
      TabIndex        =   14
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4896
      TabIndex        =   12
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5724
      TabIndex        =   13
      Top             =   90
      Width           =   1200
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2580
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   18
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   17
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   16
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   15
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   4140
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "Y"
      Top             =   2910
      Width           =   300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "專利範圍的項數 :               項"
      Height          =   180
      Left            =   3360
      TabIndex        =   53
      Top             =   3270
      Width           =   2256
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "證書形式:   　      (1:電子 2:紙本)"
      Height          =   180
      Left            =   420
      TabIndex        =   48
      Top             =   3270
      Width           =   2505
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   1572
      Left            =   6396
      TabIndex        =   5
      Top             =   2940
      Width           =   1500
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;2773"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   1440
      Width           =   6645
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11721;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額:"
      Height          =   180
      Left            =   480
      TabIndex        =   47
      Top             =   2940
      Width           =   765
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5445
      TabIndex        =   39
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "優先權證明文件份數 :               (份)"
      Height          =   180
      Left            =   3345
      TabIndex        =   38
      Top             =   2610
      Width           =   2685
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   7740
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   270
      X2              =   7830
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期 :"
      Height          =   180
      Left            =   240
      TabIndex        =   37
      Top             =   2610
      Width           =   1020
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3345
      TabIndex        =   36
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3345
      TabIndex        =   35
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   240
      TabIndex        =   34
      Top             =   2160
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   0
      Left            =   4185
      TabIndex        =   33
      Top             =   720
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3334;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3345
      TabIndex        =   32
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   240
      TabIndex        =   31
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   29
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3345
      TabIndex        =   28
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   1440
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   26
      Top             =   1080
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3440;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   2
      Left            =   4185
      TabIndex        =   25
      Top             =   1080
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3387;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   4
      Left            =   1080
      TabIndex        =   24
      Top             =   1800
      Width           =   1680
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2963;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   5
      Left            =   4185
      TabIndex        =   23
      Top             =   1800
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3387;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   6
      Left            =   1320
      TabIndex        =   22
      Top             =   2160
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2487;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   7
      Left            =   4185
      TabIndex        =   21
      Top             =   2160
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "6297;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   2460
      TabIndex        =   20
      Top             =   2940
      Width           =   2880
   End
End
Attribute VB_Name = "frm04010310_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (Combo1,lstNameAgent,Label12)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

Public strReceiveNo As String
'Modify by Morgan 2005/8/1 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, cp() As String
Dim m_CP110 As String
Dim m_CP22 As String

Dim CP10 As String
Dim intWhere As Integer
'Add By Cheng 2002/07/12
Dim m_strSitu As String '處理狀況
Public m_strCP10 As String '案件性質
Public iFrom As Integer '0=內專,1=承辦人 Add by Morgan 2011/9/22
Public oParentForm As Form '呼叫的Form Add by Morgan 2011/9/22
Dim m_CaseNo As String 'Add By Sindy 2020/3/26
Dim m_DiscType As String, m_lngDisc As Long '減免
Dim m_lngDisc1Year As Long '第一年減免金額 Add By Sindy 2020/4/8
Dim m_CP81 As String
Dim strCaseFee(1 To 2) As String 'strCaseFee(1) 國家檔中繳費年度，strCaseFee(2) 國家檔中起算日
Dim m_strOfficalFee  As String
Dim m_strServiceFee  As String
Public m_CP118isY As Boolean '是否為電子送件申請書:True.是
Dim m_bol412 As Boolean '是否有收延緩公告
Dim m_str412CP09 As String '延緩公告收文號
Dim m_str412CP71 As String '延緩月數


'Add by Morgan 2005/8/1
Private Function FormSave() As Boolean
Dim stUpdate As String
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   'Add By Sindy 2020/3/27
   cp(84) = Val(txtCP84)
   stUpdate = stUpdate & ",cp84=" & cp(84)
   If m_CP118isY = True Then
      cp(118) = "A": cp(160) = ""
      If cp(10) = 領證及繳年費 Or cp(10) = 年費 Then
         cp(160) = strSrvDate(1)
      End If
      stUpdate = stUpdate & ",cp118=" & CNULL(cp(118)) & ",cp160=" & CNULL(cp(160), True)
   End If
   
   If Frame2.Visible = True Then
      '當601領證及605年費key繳費年度而產生電子送件申請書時，將key的年度自動帶到發文作業的繳費年度
      If Val(Text7(0)) > 0 And Val(Text7(1)) > 0 Then
         stUpdate = stUpdate & ",cp53=" & Val(Text7(0)) & ",cp54=" & Val(Text7(1))
      End If
      If m_bol412 = True Then
         stUpdate = stUpdate & ",cp71=" & IIf(CheckIsDate(DBDATE(txtCP71), False) = True, DBDATE(txtCP71), Val(txtCP71))
      End If
   End If
   If lstNameAgent.Visible = True Then
      cp(110) = m_CP110 'Add By Sindy 2020/3/27
      stUpdate = stUpdate & ",cp22=" & CNULL(m_CP22) & ",cp110=" & CNULL(m_CP110)
   End If
   If stUpdate <> "" Then
      If Left(stUpdate, 1) = "," Then stUpdate = Mid(stUpdate, 2)
      strSql = "UPDATE CASEPROGRESS SET " & stUpdate & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql, intI
   End If
   
   'Add by Amy P台灣案電子化
   'Modified by Morgan 2015/1/12 工程師不要轉pdf
   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And iFrom = 0 Then
      If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
         '新增申請書轉檔記錄
         PUB_AddAppForm strReceiveNo
      End If
   End If
   
   'Add By Sindy 2022/12/28
   If TextPA178.Visible = True And TextPA178.Tag <> TextPA178.Text Then
      strSql = " UPDATE patent SET pa178='" & TextPA178 & "' WHERE PA01='" & pa(1) & "' and PA02='" & pa(2) & "' and PA03='" & pa(3) & "' and PA04='" & pa(4) & "'"
      cnnConnection.Execute strSql, intI
   End If
   '2022/12/28 END
   
   cnnConnection.CommitTrans

   FormSave = True
   
ErrorHandler:
   If Err.NUMBER <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         If lstNameAgent.Enabled = True Then
            lstNameAgent.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If Frame2.Visible = True Then
      Text7_Validate 0, Cancel
      If Cancel = True Then Exit Function
      Text7_Validate 1, Cancel
      If Cancel = True Then Exit Function
   End If
   
   Call Text9_Validate(False)
   
   'Added by Morgan 2023/8/22
   If (cp(10) = "421" Or cp(10) = "807") And Text11.Visible = True And Text11 = "" Then
      MsgBox "請輸入專利範圍的項數!!!", vbExclamation
      Text11.SetFocus
      Exit Function
   End If
   
   strExc(0) = "select cp17 from caseprogress where cp09='" & cp(9) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Val("" & RsTemp(0)) <> Val(txtCP84) Then
         If MsgBox("繳費金額" & "與收文規費(" & RsTemp(0) & ")不同，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
   End If
   'end 2023/8/22
   
   TxtValidate = True
End Function

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
'Add By Cheng 2002/07/12
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim Cancel As Boolean
Dim strFolder As String, strFileName As String 'Add By Sindy 2020/3/26
Dim bolRun As Boolean

   Select Case Index
      Case 0 '確定
         If m_strCP10 = 申請優先權證明 Then
            If IsEmptyText(Me.Text6.Text) Then
               MsgBox "優先權證明文件份數欄不可空白!!!", vbExclamation + vbOKOnly
               Me.Text6.SetFocus
               Text6_GotFocus
               Exit Sub
            ElseIf IsNumeric(Me.Text6.Text) = False Then
               MsgBox "優先權證明文件份數欄輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Text6.SetFocus
               Text6_GotFocus
               Exit Sub
            End If
            '92.1.23 CANCEL BY SONIA不可更新
            'strSQL = "Update CaseProgress Set CP17=" & Val(GetCF08(pa(1), pa(9), m_strCP10)) * Val(Me.Text6.Text) & " Where CP09='" & strReceiveNo & "'"
            'cnnConnection.Execute strSQL
            '92.1.23 END
         End If
         'Add By Sindy 2022/12/28
         If TextPA178.Visible = True Then
            If TextPA178 = "" Then
               MsgBox "證書形式不可空白！", vbExclamation + vbOKOnly
               Me.TextPA178.SetFocus
               Exit Sub
            End If
         End If
         '2022/12/28 END
         
         If TxtValidate = False Then Exit Sub
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         '2005/8/1---
         
         'Modify By Sindy 2020/3/20
         If m_CP118isY = True Then '電子送件
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            'Modify By Sindy 2022/12/28
            If cp(10) = 領證及繳年費 Or cp(10) = 年費 Then
            '2022/12/28 END
               strFolder = PUB_Getdesktop & "\領證及繳年費"
               If Dir(strFolder, vbDirectory) = "" Then
                  MkDir strFolder
               End If
               strFolder = strFolder & "\" & m_CaseNo
               If Dir(strFolder, vbDirectory) = "" Then
                  MkDir strFolder
               End If
               
               If m_bol412 = True Then
                  If Dir(PUB_Getdesktop & "\領證及繳年費\" & m_CaseNo & ".1", vbDirectory) = "" Then
                     MkDir PUB_Getdesktop & "\領證及繳年費\" & m_CaseNo & ".1"
                  End If
               End If
            Else
               strFolder = PUB_Getdesktop & "\" & m_CaseNo
               If Dir(strFolder, vbDirectory) = "" Then
                  MkDir strFolder
               End If
            End If
            
            bolRun = True
            '2.申請書
            'Add By Sindy 2020/3/20 領證及繳年費
            If cp(10) = 領證及繳年費 Then
               If StartLetter("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               'strFileName = strFolder & "\" & "申領專利證書及申請延緩公告申請書"
               strFileName = strFolder & "\" & m_CaseNo & ".data"
               'Call PUB_MakeDoc(strExc(9), strFileName)
               
               'Add By Sindy 2022/12/28
               If m_bol412 = True Then
                  '1.基本資料
                  StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False, True
                  NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(10)
                  'strFileName = strFolder & "\" & m_CaseNo & ".contact"
                  'Call PUB_MakeDoc(strExc(9), strFileName)
                  Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
                  
                  '延緩公告
                  strReceiveNo = m_str412CP09
                  If StartLetter("01", "01") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
                  strFileName = PUB_Getdesktop & "\領證及繳年費\" & m_CaseNo & ".1\" & m_CaseNo & ".data"
                  Call PUB_MakeDoc(strExc(9), strFileName, False)
                  bolRun = False
               End If
               '2022/12/28 END
               
            ElseIf cp(10) = 年費 Then
               If StartLetter("01", "06") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "06", False, strUserNum, , , True, strExc(9)
               'strFileName = strFolder & "\" & "專利年費" & IIf(m_DiscType <> "" Or m_lngDisc > 0, "減收", "") & "繳納申請書"
               strFileName = strFolder & "\" & m_CaseNo & ".data"
               'Call PUB_MakeDoc(strExc(9), strFileName)
            
            'Add By Sindy 2022/12/28
            ElseIf cp(10) = 延緩公告 Then
               If StartLetter("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & m_CaseNo & ".data"
               Call PUB_MakeDoc(strExc(9), strFileName, False)
               bolRun = False
               
            'Add By Sindy 2022/12/28
            ElseIf cp(10) = "443" Then '申請證書副本
               If StartLetter("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & m_CaseNo & ".data"
            
            'Add By Sindy 2022/12/28
            ElseIf cp(10) = 補換發證書 Then
               If StartLetter("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & m_CaseNo & ".data"
            
            'Added by Morgan 2023/8/21
            ElseIf cp(10) = "405" Or cp(10) = "436" Or cp(10) = "437" Then
               If StartLetter("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & m_CaseNo & ".data"
               Call PUB_MakeDoc(strExc(9), strFileName, False)
               bolRun = False
            ElseIf cp(10) = "421" Or cp(10) = "807" Then
               If StartLetter("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & m_CaseNo & ".data"
               If cp(10) = "421" Then
                  Call PUB_MakeDoc(strExc(9), strFileName, False)
                  bolRun = False
               End If
            End If
            
            '1.基本資料
            If bolRun = True Then
               StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False, True
               NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(10)
               'strFileName = strFolder & "\" & m_CaseNo & ".contact"
               'Call PUB_MakeDoc(strExc(9), strFileName)
               Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
            End If
            
         Else
         '2020/3/20 END
         
            If Text8 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            strLetterDate = Text5.Text
            
            'Modify by Amy 2014/08/14 P台灣案電子化 申請書修改 改開frm1105_1
            'Add by Morgan 2011/9/22
            If iFrom = 1 Then
               NowPrint strReceiveNo, "01", "16", bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
            Else
            'end 2011/9/22
            
               ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
               InsExpField m_strCP10
               'Modify By Cheng 2002/07/12
      '         NowPrint strReceiveNo, "01", "00", bolChk, strUserNum, 0
               'Mark by Amy 2014/08/14
               'NowPrint strReceiveNo, "01", IIf(m_strSitu = "", "00", m_strSitu), bolChk, strUserNum, 0
               'Add By Cheng 2003/01/19
               '申請書附件
               'Modified by Morgan 2015/6/2
               'NowPrint strReceiveNo, "01", "03", True, strUserNum, 0, , , , , , , , , , , , strReceiveNo
               'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
               '   'Modified by Morgan 2015/1/7
               '   '工程師仍然開啟Word(doc要放歷程)
               '   If iFrom = 0 Then
               NowPrint strReceiveNo, "01", "03", bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
               If bolChk = True Then
               'end 2015/6/2
                     frm1105_1.m_RecNo = strReceiveNo
                     'Modify By Sindy 2022/5/11 流水號要足6碼
                     frm1105_1.m_PdfName = Text1 & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4) & "." & CP10 & ".DATA.PDF"
                     frm1105_1.Show
                  'End If 'Removed by Morgan 2015/6/2
               End If
            End If 'Add by Morgan 2011/9/22
            'end 2014/08/14
         End If
         
         'Modify by Morgan 2011/9/22 配合承辦人系統也要用
         'frm040103_1.Show
         'frm040103_1.ClearForm
         oParentForm.Show
         oParentForm.ClearForm
         'end 2011/9/22
      Case 1
         'Modify by Morgan 2011/9/22 配合承辦人系統也要用
         'frm040103_1.Show
         oParentForm.Show
         
      Case 2 '結束
         'Modify by Morgan 2011/9/22 配合承辦人系統也要用
         'Unload frm040103_1
         Unload oParentForm
   End Select
   'Add By Sindy 2020/3/26
   If Frame2.Visible = True Then
      Unload frm040104_7
   End If
   '2020/3/26 END
   Unload Me
End Sub

'申請書
Private Function StartLetter(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP07Add2M As String
Dim strAD15 As String, strAD16 As String
Dim strOa02 As String
Dim strNote As String '備註內容
Dim strNote2 As String '申請內容

   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())

   '出名代理人
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         'Add By Sindy 2019/12/13
'         strOa02 = IIf(strOa02 <> "", "、", "") & PUB_ConvertNameFormat("" & .Fields("st02"))
'         '2019/12/13 END
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
   'Modify By Sindy 2020/4/8 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, cp(110), ET01, strReceiveNo, ET03, ii, strTxt())
   
   If cp(10) = 領證及繳年費 Or cp(10) = 年費 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳納起年','" & Text7(0) & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳納迄年','" & Text7(1) & "')"
      
      If cp(10) = 領證及繳年費 Then
         '一般資格繳費項目
'         'Modify By Sindy 2019/10/24 無減免才要顯示
'         If Not (m_DiscType <> "" Or m_lngDisc > 0) Then
'         '2019/10/24 END
         'Modify By Sindy 2020/3/13
         If m_CP81 <> "Y" Then   '無減免
         '2020/3/13 END
            strTmp = ""
            If Text7(0) = 1 Then
               'Modified by Morgan 2023/1/18 第1年年費加倍時也要考慮 Ex:P-124857
               'strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & Format(frm040104_7.m_strOfficalFee1) & "元及第1年年費" & Format(frm040104_7.m_lngOfficalFee1Year) & "元(共計" & Format(frm040104_7.m_lngFee1) & "元)"
               strExc(1) = Format(frm040104_7.m_strOfficalFee1)
               strExc(2) = Format(IIf(Text9.Text = "Y", 2, 1) * frm040104_7.m_lngOfficalFee1Year)
               strExc(3) = Format(Val(strExc(1)) + Val(strExc(2)))
               strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & strExc(1) & "元及第1年年費" & strExc(2) & "元(共計" & strExc(3) & "元)"
               'end 2023/1/18
            End If
            If Text7(1) <> 1 Then
               If strTmp <> "" Then strTmp = strTmp & "，及"
               'Modified by Morgan 2023/1/18 第1年年費加倍時也要考慮 Ex:P-124857
               'strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & Format(frm040104_7.m_lngFee2) & "元，合計" & Format(frm040104_7.m_lngFee1 + frm040104_7.m_lngFee2) & "元。"
               strExc(4) = Format(frm040104_7.m_lngFee2)
               strExc(5) = Format(Val(strExc(3)) + Val(strExc(4)))
               strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & strExc(4) & "元，合計" & strExc(5) & "元。"
               'end 2023/1/18
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一般資格繳費項目','" & strTmp & "')"
         End If
         
         'Add By Sindy 2022/12/27 + 申請證書形式
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請證書形式','" & IIf(TextPA178 = "1", "電子證書", IIf(TextPA178 = "2", "紙本證書", "電子證書/紙本證書")) & "')"
         '2022/12/27 END
      Else
         'Modify By Sindy 2019/11/15 年費第7年(含)以後自然人、學校、中小企業皆無減免，故應改成年費繳納申請書
         If Val(Text7(0)) >= 7 Then
            m_DiscType = ""
            m_lngDisc = 0
         End If
         '2019/11/15 END
      End If
      
      'Modify By Sindy 2020/3/13
      If m_CP81 = "Y" Then   '有減免
      '2020/3/13 END
         If m_DiscType <> "" Or m_lngDisc > 0 Then
            '**************************************************************************
            '申請人1~5
            '**************************************************************************
            For jj = 0 To 4
               If pa(26 + jj) <> "" Then
                  'Add By Sindy 2019/11/15
                  Call PUB_GetAD03(pa(26 + jj), pa(9), m_DiscType, , strAD15, strAD16) '申請人減免資料
                  '2019/11/15 END
                  
                  '符合年費減收資格
                  strTmp = ""
                  If m_DiscType = "1" Then
                     strTmp = "自然人"
                  ElseIf m_DiscType = "2" Then
                     strTmp = "學校"
                  ElseIf m_DiscType = "3" Then
                     strTmp = "中小企業"
                  End If
                  If jj = 0 Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','符合年費減收資格','" & strTmp & "')"
                  Else
                     '備註
                     'If strNote <> "" Then strNote = strNote & vbCrLf '加折行
                     strNote = strNote & vbCrLf & "申請人" & jj + 1 & "減免資格身分：" & strTmp
                  End If
                  'Modify By Sindy 2019/1/11
                  strTmp = ""
                  '中小企業符合減收資格依據
                  If m_DiscType = "3" Then '中小企業
                     'Modify By Sindy 2019/11/15
         '               strTmp = vbCrLf & "　製造業、營造業、礦業及土石採取業實收資本額八千萬以下：           元" & vbCrLf & _
         '                        "　前項除外之其他行業前一年營業額一億元以下：           元" & vbCrLf & _
         '                        "　我國製造業、營造業、礦業及土石採取業實收資本額新台幣八千萬以上但經常僱用員工數未滿200人：員工數        人" & vbCrLf & _
         '                        "　我國前項除外之其他行業前一年營業額一億元以上者但經常僱用員工數未滿100人：員工數        人"
                     'Modify By Sindy 2020/7/24 傳入中小企業符合減收資格依據代碼 , 轉換中文
                     strTmp = PUB_AD15ToText(strAD15, strAD16)
                     '2020/7/24 END
'                     If strAD15 = "1" Then
'                        strTmp = "製造業、營造業、礦業及土石採取業實收資本額八千萬以下：" & strAD16 & "元"
'                     ElseIf strAD15 = "2" Then
'                        strTmp = "前項除外之其他行業前一年營業額一億元以下：" & strAD16 & "元"
'                     ElseIf strAD15 = "3" Then
'                        strTmp = "我國製造業、營造業、礦業及土石採取業實收資本額新台幣八千萬以上但經常僱用員工數未滿200人：員工數" & strAD16 & "人"
'                     Else '4
'                        strTmp = "我國前項除外之其他行業前一年營業額一億元以上者但經常僱用員工數未滿100人：員工數" & strAD16 & "人"
'                     End If
                     '2019/11/15 END
                  End If
                  If strTmp <> "" Then
                     If jj = 0 Then
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','中小企業符合減收資格依據','" & strTmp & "')"
                     Else
                        '備註
                        If strNote <> "" Then strNote = strNote & vbCrLf '加折行
                        strNote = strNote & "中小企業符合減收資格依據：" & strTmp
                     End If
                  End If
               End If
            Next jj
            '**************************************************************************
            
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減收','減收')"
            If cp(10) = 領證及繳年費 Then
               '減收資格繳費項目
               strTmp = ""
               If Text7(0) = 1 Then
                  'Modified by Morgan 2023/1/18 第1年年費加倍時也要考慮 Ex:P-124857
                  'strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & Format(frm040104_7.m_strOfficalFee1) & "元及第1年年費" & Format(frm040104_7.m_lngOfficalFee1Year - m_lngDisc1Year) & "元(共計" & Format(frm040104_7.m_lngFee1 - m_lngDisc1Year) & "元)"
                  strExc(1) = Format(frm040104_7.m_strOfficalFee1)
                  strExc(2) = Format(IIf(Text9.Text = "Y", 2, 1) * (frm040104_7.m_lngOfficalFee1Year - m_lngDisc1Year))
                  strExc(3) = Format(Val(strExc(1)) + Val(strExc(2)))
                  strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & strExc(1) & "元及第1年年費" & strExc(2) & "元(共計" & strExc(3) & "元)"
                  'end 2023/1/18
               End If
               If Text7(1) <> 1 Then
                  If strTmp <> "" Then strTmp = strTmp & "，及"
                  'Modified by Morgan 2023/1/18 第1年年費加倍時也要考慮 Ex:P-124857
                  'strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & Format(frm040104_7.m_lngFee2) - (m_lngDisc - m_lngDisc1Year) & "元，合計" & Format(frm040104_7.m_lngFee1 - m_lngDisc + frm040104_7.m_lngFee2) & "元。"
                  strExc(4) = Format(frm040104_7.m_lngFee2 - m_lngDisc + IIf(Text9.Text = "Y", 2, 1) * m_lngDisc1Year)
                  strExc(5) = Format(Val(strExc(3)) + Val(strExc(4)))
                  strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & strExc(4) & "元，合計" & strExc(5) & "元。"
                  'end 2023/1/18
               End If
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減收資格繳費項目','" & strTmp & "')"
            End If
         End If
      End If
'      'Modify By Sindy 2020/5/13
'      '申領專利證書及申請延緩公告申請書:
'      '有備註
'      If strNote <> "" And cp(10) = 領證及繳年費 Then
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strNote & "')"
'      End If
      If Me.Text10.Text = "Y" Then '逾期補繳
         '逾期補繳=>繳費金額含逾期費用。
         strNote2 = "繳費金額含逾期費用。"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','繳費金額含逾期費用。')"
      ElseIf cp(10) = 年費 And pa(143) = "N" Then '年費申請人是否出名為"N"
         strNote2 = "代理人" & strOa02 & "僅辦理年費繳納之事宜，本案後續相關程序之進行，均維持原代理人。"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','代理人" & strOa02 & "僅辦理年費繳納之事宜，本案後續相關程序之進行，均維持原代理人。')"
      End If
      '年費繳納申請書:
      '有申請內容
      If strNote <> "" Then
         strNote2 = strNote2 & strNote
      End If
      If strNote2 <> "" And cp(10) = 年費 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','" & strNote2 & "')"
      End If
      '2020/5/13 END
      
      '有收414回復原狀才需要顯示此段內容
      If PUB_ChkCPExist(cp, "414", 1) Then '1=未發文
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請回復領證','♀')"
      End If
   End If
   'If chk412.Value = 1 And chk412.Visible = True Then
   If m_bol412 = True Then
      If CheckIsDate(DBDATE(txtCP71), False) = True Then
         '日期
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','申請延緩公告至民國 " & PUB_DBYEAR(txtCP71) - 1911 & " 年 " & PUB_DBMONTH(txtCP71) & " 月 " & PUB_DBDAY(txtCP71) & "日')"
         '備註
         If strNote <> "" Then strNote = strNote & vbCrLf '加折行
         strNote = strNote & "申請延緩公告至民國 " & PUB_DBYEAR(txtCP71) - 1911 & " 年 " & PUB_DBMONTH(txtCP71) & " 月 " & PUB_DBDAY(txtCP71) & "日"
      Else
         '月份
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延緩月數','" & txtCP71 & "')"
      End If
   End If
   '2024/10/17 END
   'Modify By Sindy 2020/5/13
   '申領專利證書及申請延緩公告申請書:
   '有備註
   If strNote <> "" And (cp(10) = 領證及繳年費 Or m_bol412 = True) Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strNote & "')"
   End If
   
   'Added by Morgan 2023/8/22
   If cp(10) = 申請優先權證明 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','份數','" & Val(Text6) & "')"
   
      If PUB_ChkCPExist(cp(), "437", "1") Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','電子交換','♀')"
      End If
   End If
   If cp(10) = "421" Or cp(10) = "807" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','項數','" & Val(Text11) & "')"
      
      If cp(10) = "421" And pa(14) = "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','未公告','♀')"
      End If
   End If
   'end 2023/8/22
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   
   'Modify By Sindy 2020/4/9 有繳費金額就要帶出收據抬頭
   If Val(txtCP84) > 0 Then
      Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 3, , ET01, strReceiveNo, ET03, ii, strTxt())
   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter = True
   End If
End Function

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField(ByVal strCP10 As String)
Dim strSql As String
Dim strTemp As String
'Add By Cheng 2003/02/13
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
' 案件性質
Select Case strCP10
Case 申請優先權證明
    ' 清除定稿例外欄位檔原有資料
    EndLetter "01", strReceiveNo, "00", strUserNum
    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('01','" & strReceiveNo & "','00','" & strUserNum & _
             "','附件幾份','" & Me.Text6.Text & "')"
    cnnConnection.Execute strSql
    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('01','" & strReceiveNo & "','00','" & strUserNum & _
             "','規費','" & Val(GetCF08(pa(1), pa(9), m_strCP10)) * Val(Me.Text6.Text) & "')"
    cnnConnection.Execute strSql
'Add By Cheng 2003/02/13
'案件性質為授權時, 加例外欄位的處理
Case 授權
    '附件
    ' 清除定稿例外欄位檔原有資料
    EndLetter "01", strReceiveNo, "03", strUserNum
    StrSQLa = "Select CP72,CU04,CU11,CU23,CP50 From Caseprogress,Customer Where substr(CP72,1,8)=CU01(+) And substr(CP72,9,1)=CU02(+) And CP09='" & strReceiveNo & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('01','" & strReceiveNo & "','03','" & strUserNum & _
                 "','被授權人名稱','" & IIf("" & rsA("CU04").Value <> "", "" & rsA.Fields("CU04").Value, "" & rsA("CP50").Value) & "')"
        cnnConnection.Execute strSql
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('01','" & strReceiveNo & "','03','" & strUserNum & _
                 "','被授權人IDNO','" & IIf("" & rsA("CU11").Value <> "", "" & rsA.Fields("CU11").Value, "") & "')"
        cnnConnection.Execute strSql
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('01','" & strReceiveNo & "','03','" & strUserNum & _
                 "','被授權人地址','" & IIf("" & rsA("CU23").Value <> "", "" & rsA.Fields("CU23").Value, "") & "')"
        cnnConnection.Execute strSql
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
'Add By Cheng 2003/03/27
'案件性質為申請英文證明時, 加例外欄位的處理
Case 申請英文證明
    '英譯本
    ' 清除定稿例外欄位檔原有資料
    EndLetter "01", strReceiveNo, "03", strUserNum
    StrSQLa = "Select PA08, NA07,NA09,NA11 From Caseprogress, Patent, Nation Where CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND PA09=NA01 And CP09='" & strReceiveNo & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('01','" & strReceiveNo & "','03','" & strUserNum & _
                 "','專用年度','" & IIf("" & rsA.Fields(0).Value = "1", "" & rsA.Fields(1).Value, IIf("" & rsA.Fields(0).Value = "2", "" & rsA.Fields(2).Value, "" & rsA.Fields(3).Value)) & "')"
        cnnConnection.Execute strSql
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
Case Else
 '無動作
End Select
End Sub
'Add by Morgan 2011/9/22
Private Sub Form_Activate()
   'Added by Morgan 2023/8/22
   If Me.m_strCP10 <> "421" And Me.m_strCP10 <> "807" Then
      Label14.Visible = False
      Text11.Visible = False
   End If
   'end 2023/8/22
   
   If Me.m_strCP10 <> 申請優先權證明 Then
      Label4.Visible = False
      Text6.Visible = False
   End If
   
   If iFrom = 1 Then
      lstNameAgent.Enabled = False
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   'Modify by Morgan 2011/9/22 配合承辦人系統也要用
   'With frm040103_1
   With oParentForm
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
      'Add By Cheng 2002/07/12
      m_strSitu = ""
      
      If iFrom = 0 Then
         If .Text6 = "2" Then
            Me.Text8.Text = "Y"
            m_strSitu = "08"
         End If
      End If
      
   End With
   'Add by Morgan 2005/7/29
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
   ReadPatent
   
   If Me.m_strCP10 = 申請優先權證明 Then
      Me.Text6.Text = "1"
      Text6_Validate False 'Added by Morgan 2023/8/22
   'Added by Morgan 2015/6/2
   ElseIf Me.m_strCP10 = "232" Then
      Text8.Text = ""
   'end 2015/6/2
   End If

   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   'Add by Morgan 2005/7/14
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   If pa(9) = "000" Then
      '要傳入案件性質:年費會排除桂所長
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10), True  'Modified by Morgan 2021/12/10 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 END
   
   Combo1.ListIndex = 0
   Text5 = strSrvDate(2)
   
   'Added by Morgan 2023/8/22
   If m_CP118isY Then
      Label18(1).Visible = False
      Text8.Visible = False
   End If
   'end 2023/8/22
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010310_1 = Nothing
End Sub

Private Sub ReadPatent()
Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
Dim strTmp1(0 To 5) As String
Dim i As Integer
Dim m_strNP09_1 As String
   
   For Each Lbl In Label12
      Lbl.Caption = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   
   If pa(1) = "P" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         Label12(1) = pa(11)
         Label12(2) = pa(22)
         AddCboName Combo1, pa(5), pa(6), pa(7)
      End If
   ElseIf pa(1) = "PS" Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         Label12(1) = pa(11)
         AddCboName Combo1, pa(5), pa(6), pa(7)
      End If
   End If
   
   'Add By Sindy 2020/3/26
   cp(9) = strReceiveNo
   Call PUB_ReadCaseProgressDatabase(cp(), intWhere)
   '2020/3/26 END
   
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,cp10,CP110 from caseprogress,casepropertymap,staff,staff staff1 where " & _
      "cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
      "cp14=staff.st01(+) and cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
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
      If Not IsNull(.Fields(4)) Then CP10 = .Fields(4)
   End If
   End With
   
   'Add By Sindy 2020/3/26
   If CP10 = 領證及繳年費 Or CP10 = 年費 Then
      m_CP81 = "" '設定案件是否可減免
      If PUB_GetCaseDiscStat(pa(1) & pa(2) & pa(3) & pa(4), m_DiscType) = "Y" Then
         m_CP81 = "Y"
      Else
         m_CP81 = "N"
      End If
      
      '讀取繳年費記錄
      strTmp1(0) = strReceiveNo
      For i = 1 To 4
         strTmp1(i) = pa(i)
      Next
      If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
      End If
   End If
   
   Frame2.Visible = False: Frame1.Visible = False
   If CP10 = 延緩公告 Then
      Frame1.Visible = True
      
   ElseIf CP10 = 領證及繳年費 Then
      'Added by Morgan 2021/12/14
      '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
      If PUB_CheckFormExist("frm040104_7") = False Then
         Set frm040104_7 = Nothing
      End If
      'end 2021/12/14
   
      Frame2.Visible = True: Frame1.Visible = True
      frm040104_7.SetParent Me
      'Modified by Morgan 2022/3/17
      'frm040104_7.Hide
      frm040104_7.Show
      '觸發是否規費雙倍檢查
      Call frm040104_7.TxtValidate
      frm040104_7.Hide
      'end 2022/3/17
      
      Me.Text7(0).Text = frm040104_7.Text7(0).Text '繳納起始年度
      If Val(Me.Text7(0).Text) = 0 Then Me.Text7(0).Text = "1"
      Me.Text7(1).Text = frm040104_7.Text7(1).Text '繳納終止年度
      
      If Val(cp(53)) > 0 And Val(cp(54)) > 0 Then
         Text7(0) = cp(53): Text7(1) = cp(54)
      End If
      
      Me.Text9.Text = frm040104_7.Text6.Text '費用是否要雙倍
      
      Call Text9_Validate(False) '計算繳費金額
      
   ElseIf CP10 = 年費 Then
      Frame2.Visible = True
      Text7(0).Text = ""
      Text7(0).Enabled = True
      Label6.Visible = True
      Text10.Visible = True
      
      If Val(cp(53)) > 0 And Val(cp(54)) > 0 Then
         Text7(0) = cp(53): Text7(1) = cp(54)
      End If
      
      Call Text9_Validate(False) '計算繳費金額
      
      m_strNP09_1 = PUB_GetNextFeeDate(pa)
      '若法定期限為假日時, 抓大於法定期限最近的工作天
      If m_strNP09_1 <> "" Then
         m_strNP09_1 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09_1)))
      End If
      If strSrvDate(1) > m_strNP09_1 Then
         '費用雙倍
         Me.Text9.Text = "Y"
         Me.Text10.Text = "Y" '逾期補繳
      Else
         '取消費用雙倍
         '2005/3/1 已設雙倍不可清除
         'Me.Text9.Text = ""
      End If
   End If
   '2020/3/26 END
   If CP10 = 延緩公告 Or CP10 = 領證及繳年費 Then
      'Add by Morgan 2004/6/24 檢查是否有延緩公告未發文
      If Val(pa(14)) = 0 Or Val(pa(14)) >= 930701 Then
         m_bol412 = PUB_Get412Data(pa, m_str412CP09, m_str412CP71)
         If m_bol412 = True Then
            Me.chk412.Visible = True
            Me.chk412.Enabled = False
            Me.chk412.Value = 1
            Me.lblCP71.Visible = True
            Me.Label8.Visible = True
            Me.txtCP71.Visible = True
            Me.txtCP71.Enabled = True
            'Add By Sindy 2020/11/12 若有收文"412延緩公告"直接帶入延緩的月數
            If Val(m_str412CP71) > 0 Then
               txtCP71 = m_str412CP71
            End If
            '2020/11/12 END
         End If
      End If
   End If
   
   'Add By Sindy 2022/12/28
   If (cp(10) = 領證及繳年費 Or cp(10) = 補換發證書) And strSrvDate(1) >= "20230101" Then
      Label10.Visible = True
      TextPA178.Visible = True
      TextPA178.Text = pa(178)
      TextPA178.Tag = pa(178)
   Else
      Label10.Visible = False
      TextPA178.Visible = False
   End If
   '2022/12/28 END
End Sub

Private Sub Text10_GotFocus()
   InverseTextBox Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub


Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

'Added by Morgan 2023/8/22
Private Sub Text11_Validate(Cancel As Boolean)
   If Not IsEmptyText(Me.Text11.Text) Then
      If IsNumeric(Me.Text11.Text) = False Then
         MsgBox "專利範圍的項數輸入錯誤!!!", vbExclamation + vbOKOnly
         Cancel = True
      
      ElseIf Text11.Tag <> Text11.Text Then
         txtCP84 = Val(GetCF08(pa(1), pa(9), cp(10))) + IIf(Val(Me.Text11.Text) > 10, 600 * (Val(Me.Text11.Text) - 10), 0)
         Text11.Tag = Text11.Text
      'end 2023/8/22
      End If
   End If
   If Cancel = True Then TextInverse Text11
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
TextInverse Me.Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
'Modified by Morgan 2023/8/22
'If IsEmptyText(Me.Text6.Text) Then
If Not IsEmptyText(Me.Text6.Text) Then
'end 2023/8/22
   If IsNumeric(Me.Text6.Text) = False Then
      MsgBox "優先權證明文件份數輸入錯誤!!!", vbExclamation + vbOKOnly
      Cancel = True
   'Added by Morgan 2023/8/22
   ElseIf Text6.Tag <> Text6.Text Then
      txtCP84 = Val(GetCF08(pa(1), pa(9), cp(10))) * Val(Me.Text6.Text)
      Text6.Tag = Text6.Text
   'end 2023/8/22
   End If
End If
If Cancel = True Then TextInverse Text6
End Sub

Private Sub Text7_GotFocus(Index As Integer)
  TextInverse Text7(Index)
End Sub
Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
Dim i As Integer, bolChk As Boolean, varTmp As Variant
Dim varTmpNICK As Variant, TMPnick060104 As Integer
Dim strNextFeeDate As String '下次繳費日
   
   If Text7(Index) <> "" Then
      If Index = 1 Then
         If ChkRange(Text7(0), Text7(1), "繳費年度") = True Then
            For i = Text7(0) To Text7(1)
               If InStr(pa(72), Format(i)) > 0 Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = True Then
               MsgBox "繳費年度錯誤，請查明後再輸入 !", vbCritical
               Cancel = True
               Exit Sub
            '92.7.7 ADD BY SONIA
            Else
               varTmp = Split(strCaseFee(2), ",")
               If Text7(1) > UBound(varTmp) + 1 Then
                  MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
                  Cancel = True
                  Exit Sub
                  
               'Add by Morgan 2011/7/1
               Else
                  If cp(81) = "Y" And pa(8) = "3" And Val(Text7(1)) < 3 And Val(Text7(1)) <> UBound(varTmp) + 1 Then
                     If UBound(varTmp) + 1 < 3 Then
                        strExc(1) = UBound(varTmp) + 1
                     Else
                        strExc(1) = 3
                     End If
                     MsgBox "繳費年度請輸入 " & strExc(1) & " 以上(可減免客戶1~3年免繳年費)!!"
                     Cancel = True
                     Exit Sub
                  End If
               
               End If
            '92.7.7 END
            End If
         Else
            Cancel = True
            Exit Sub
         End If
      End If
   Else
      MsgBox "年度不可空白 !", vbCritical
      TextInverse Text7(Index)
      Cancel = True
   End If
   If Cancel Then
      TextInverse Me.Text7(Index)
   Else
      Text9_Validate False
   End If
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Cheng 2002/07/12
'取得案件收費表的規費
Private Function GetCF08(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF08 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF08 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF08 IS NOT NULL"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetCF08 = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'Add by Morgan 2005/7/29
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/10 Forms2.0 改用模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      m_CP22 = ""
   Else
      m_CP22 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
   If Text7(0) <> "" And Text7(1) <> "" Then
      '計算繳費金額
      If cp(10) = 年費 Then
         'p_strOfficalFee 規費
         'p_strServiceFee  服務費
         'p_lngDisc  減免金額
         'cp(81) => m_CP81
         PUB_GetPatentYearFee pa(9), pa(8), "Y00000001", cp(10), Me.Text7(0).Text, Me.Text7(1).Text, _
            IIf(Me.Text9.Text = "Y", True, False), m_CP81, pa(14), strSrvDate(1), _
            m_strOfficalFee, m_strServiceFee, m_lngDisc
         txtCP84 = m_strOfficalFee
      Else
         Call frm040104_7.ChkPatentYearFee(pa(9), pa(8), "Y00000001", CP10, Me.Text7(0).Text, Me.Text7(1).Text, _
         IIf(Me.Text9.Text = "Y", True, False), False)
         m_lngDisc1Year = frm040104_7.m_lngDisc1Year '第一年減免金額
         m_lngDisc = frm040104_7.m_lngDisc
         txtCP84 = frm040104_7.m_lngFinalFee
      End If
   End If
End Sub

'Add By Sindy 2022/12/28
Private Sub TextPA178_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP71_GotFocus()
   CloseIme
   TextInverse txtCP71
End Sub

Private Sub txtCP71_Validate(Cancel As Boolean)
   If txtCP71 = "" Then
      Exit Sub
   ElseIf Val(txtCP71) <> Val(m_str412CP71) Then
      MsgBox "延緩月數/日期必須與分案時相同！", vbCritical
      Cancel = True
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
