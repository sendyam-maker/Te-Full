VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010313_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-自請撤回(重新委任)"
   ClientHeight    =   3900
   ClientLeft      =   180
   ClientTop       =   1290
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7785
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   6840
      TabIndex        =   26
      Top             =   1170
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   27
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   28
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1152
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2655
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2655
      Width           =   300
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1152
      MaxLength       =   3
      TabIndex        =   12
      Top             =   555
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1704
      MaxLength       =   6
      TabIndex        =   11
      Top             =   555
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2556
      MaxLength       =   1
      TabIndex        =   10
      Top             =   555
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2796
      MaxLength       =   2
      TabIndex        =   9
      Top             =   555
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   630
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3045
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6800
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4848
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5676
      TabIndex        =   5
      Top             =   70
      Width           =   1100
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1152
      TabIndex        =   36
      Top             =   1230
      Width           =   5625
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "9922;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   195
      Index           =   7
      Left            =   4350
      TabIndex        =   35
      Top             =   2070
      Width           =   3345
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5900;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   195
      Index           =   6
      Left            =   1152
      TabIndex        =   34
      Top             =   2064
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   195
      Index           =   5
      Left            =   4350
      TabIndex        =   33
      Top             =   1800
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   195
      Index           =   4
      Left            =   1152
      TabIndex        =   32
      Top             =   1800
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   195
      Index           =   2
      Left            =   4350
      TabIndex        =   31
      Top             =   900
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   195
      Index           =   1
      Left            =   1152
      TabIndex        =   30
      Top             =   900
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   195
      Index           =   0
      Left            =   4350
      TabIndex        =   29
      Top             =   600
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   6270
      TabIndex        =   2
      Top             =   2640
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5355
      TabIndex        =   25
      Top             =   2685
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   50
      X2              =   7720
      Y1              =   2424
      Y2              =   2424
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   50
      X2              =   7820
      Y1              =   2448
      Y2              =   2448
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "2.  申請人無法提供"
      Height          =   180
      Index           =   3
      Left            =   1152
      TabIndex        =   24
      Top             =   3330
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   2280
      TabIndex        =   22
      Top             =   2700
      Width           =   2880
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3540
      TabIndex        =   21
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3540
      TabIndex        =   20
      Top             =   2064
      Width           =   768
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   117
      TabIndex        =   19
      Top             =   2064
      Width           =   948
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3540
      TabIndex        =   18
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   300
      TabIndex        =   17
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   300
      TabIndex        =   16
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   300
      TabIndex        =   15
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3540
      TabIndex        =   14
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   300
      TabIndex        =   13
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "說明:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3060
      Width           =   405
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "1.  IPO要求撤回"
      Height          =   180
      Index           =   0
      Left            =   1152
      TabIndex        =   7
      Top             =   3090
      Width           =   1215
   End
End
Attribute VB_Name = "frm06010313_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/13 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Add by Morgan 2007/9/13
Option Explicit

Dim strReceiveNo As String
Dim pa() As String, m_CP110 As String, m_AgentName As String
Dim intWhere As Integer
Dim m_CP43 As String


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(5) As String
Dim ii As Integer

   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   If Text6 = "1" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','其他理由','依  鈞局來函(來電)辦理')"
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Function FormSave() As Boolean
   Dim stCP64 As String

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(m_CP110) & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
   End If
End Function
Private Sub cmdOK_Click(Index As Integer)

   Dim bolChk As Boolean, strTmp As String
   Select Case Index
      Case 0
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
         If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
             MsgBox MsgText(1111), vbInformation
             If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                Exit Sub
             End If
         End If
         'end 2020/02/17
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If Text7 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         
         StartLetter "01", "03"
         strLetterDate = Text5.Text
         NowPrint strReceiveNo, "01", "03", bolChk, strUserNum, , , , , 2
         
         frm060103_1.Show
         frm060103_1.ClearForm
         Unload Me
      Case 1
         frm060103_1.Show
         Unload Me
      Case 2
         Unload frm060103_1
         Unload Me
   End Select
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
   ReDim pa(TF_PA)
   ReadPatent
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010313_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
   For Each Lbl In Label12
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      AddCboName Combo1, pa(5), pa(6), pa(7)
   End If
   
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,CP110 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      Label12(0) = "" & .Fields(0)
      Label12(4) = "" & .Fields(1)
      Label12(5) = "" & .Fields(2)
      m_CP43 = "" & .Fields("CP43")
      m_CP110 = "" & .Fields("cp110")
      '抓相關總收文號內容
      If m_CP43 <> Empty Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & m_CP43 & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With rsTemp1
               If m_CP43 > "C" Then
                  Label12(6) = TransDate("" & .Fields("CP05"), 1)
                  Label12(7) = "" & .Fields("CP08")
               End If
            End With
         End If
      End If
   End If
   End With
   'Remove by Lydia 2020/02/21
   'strExc(0) = "SELECT Max(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & 年費
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'end 2020/02/21
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
End Sub

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
   If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
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

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   If Text6.Text = "" Then
      MsgBox "說明選項不可空白！", vbExclamation
      Text6.SetFocus
      Exit Function
   End If
   
   If Me.Text5.Enabled = True Then
      Cancel = False
      Text5_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If

TxtValidate = True
End Function

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
End Sub
