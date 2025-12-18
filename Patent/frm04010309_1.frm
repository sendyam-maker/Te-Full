VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010309_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-詢進度"
   ClientHeight    =   3735
   ClientLeft      =   555
   ClientTop       =   1815
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   6924
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4872
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5700
      TabIndex        =   4
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   1
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   2
      Top             =   3360
      Width           =   300
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      TabIndex        =   10
      Top             =   1380
      Width           =   6630
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11695;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   7680
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   7680
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請書類別            (1.爭議案審查確定 2.年費繳交狀況)"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期"
      Height          =   180
      Left            =   180
      TabIndex        =   28
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3180
      TabIndex        =   27
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3180
      TabIndex        =   26
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   25
      Top             =   2160
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   0
      Left            =   4020
      TabIndex        =   24
      Top             =   720
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3598;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3180
      TabIndex        =   23
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   21
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3180
      TabIndex        =   19
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   18
      Top             =   1440
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   17
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
      Index           =   2
      Left            =   4020
      TabIndex        =   16
      Top             =   1080
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   4
      Left            =   1020
      TabIndex        =   15
      Top             =   1800
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3334;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   5
      Left            =   4020
      TabIndex        =   14
      Top             =   1800
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   6
      Left            =   1260
      TabIndex        =   13
      Top             =   2160
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2752;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   7
      Left            =   4020
      TabIndex        =   12
      Top             =   2160
      Width           =   3600
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "6350;317"
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
      Left            =   180
      TabIndex        =   11
      Top             =   3360
      Width           =   2880
   End
End
Attribute VB_Name = "frm04010309_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (Combo1,Label12)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit
Public strReceiveNo As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, CP10 As String
Dim pa() As String, CP10 As String

Dim intWhere As Integer

Private Sub cmdOK_Click(Index As Integer)
 Dim bolChk As Boolean, strTmp As String
   Select Case Index
      Case 0
         'Add by Amy 2014/08/14 解 未輸入申請書類別無法產生定稿
         If Trim(Text6) = "" Then
            MsgBox "申請書類別不可為空", vbCritical
            Me.Text6.SetFocus
            Text6_GotFocus
            Exit Sub
         End If
         'end 2014/08/14
         If Text7 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         Select Case Text6.Text
            Case "1"
               If CP10 = 舉發 Or 異議_專 Then
                  '舉發成立 11
                  strTmp = "11"
               ElseIf CP10 = 被舉發理由 Or 被異議理由 Then
                  '被舉發不成立 12
                  strTmp = "12"
               Else
                  MsgBox "案件性質只有異議,舉發或被異議理由,被舉發理由時才有申請書 !", vbCritical
                  Exit Sub
               End If
            Case "2"
               '年費繳交狀況 13
               strTmp = "13"

         End Select
         
         'Modify by Amy P台灣案電子化
         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
         If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
            '新增申請書轉檔記錄
            PUB_AddAppForm strReceiveNo
         End If
         End If
         
         strLetterDate = Text5.Text
         NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
         If bolChk = True Then
             frm1105_1.m_RecNo = strReceiveNo
             'Modify By Sindy 2022/5/11 流水號要足6碼
             frm1105_1.m_PdfName = Text1 & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4) & "." & CP10 & ".DATA.PDF"
             frm1105_1.Show
         End If
         End If
         'end 2014/08/14
         
         frm040103_1.Show
         ' 90.08.27 modify by louis
         frm040103_1.ClearForm
      Case 1
         frm040103_1.Show
      Case 2
         Unload frm040103_1
   End Select
   Unload Me
End Sub

Private Sub Form_Initialize()
    'add by nickc 2007/02/02
    ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()

    
   MoveFormToCenter Me
   intWhere = 國內
   With frm040103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
      Text7 = "Y"
   End With
   ReadPatent
   Combo1.ListIndex = 0
   Text5 = strSrvDate(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010309_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
   For Each Lbl In Label12
      Lbl.Caption = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      AddCboName Combo1, pa(5), pa(6), pa(7)
   End If
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,cp10 from caseprogress,casepropertymap,staff,staff staff1 where " & _
      "cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
      "cp14=staff.st01(+) and cp13=staff1.st01(+)"
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
