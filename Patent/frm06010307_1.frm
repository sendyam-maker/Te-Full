VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010307_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-更改 "
   ClientHeight    =   4590
   ClientLeft      =   -1140
   ClientTop       =   2560
   ClientWidth     =   7970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7970
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   300
      Left            =   3672
      TabIndex        =   37
      Top             =   3636
      Width           =   2208
      Begin VB.OptionButton Option2 
         Caption         =   "不請款"
         Height          =   280
         Index           =   1
         Left            =   1188
         TabIndex        =   39
         Top             =   0
         Width           =   912
      End
      Begin VB.OptionButton Option2 
         Caption         =   "需請款"
         Height          =   280
         Index           =   0
         Left            =   108
         TabIndex        =   38
         Top             =   0
         Value           =   -1  'True
         Width           =   876
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "因客戶/本所誤繕需更改 (                                                )"
      Height          =   280
      Index           =   1
      Left            =   1404
      TabIndex        =   3
      Top             =   3636
      Width           =   4728
   End
   Begin VB.OptionButton Option1 
      Caption         =   "因智慧局來函誤繕"
      Height          =   280
      Index           =   0
      Left            =   1404
      TabIndex        =   2
      Top             =   3240
      Value           =   -1  'True
      Width           =   2064
   End
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   345
      Left            =   1050
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   930
         Style           =   1  '圖片外觀
         TabIndex        =   35
         Top             =   30
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   60
         TabIndex        =   36
         Top             =   60
         Width           =   765
      End
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   4590
      MaxLength       =   7
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6945
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4890
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5760
      TabIndex        =   6
      Top             =   70
      Width           =   1110
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1248
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2844
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010307_1.frx":0000
      Left            =   1050
      List            =   "frm06010307_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   11
      Top             =   540
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1605
      MaxLength       =   6
      TabIndex        =   10
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2445
      MaxLength       =   1
      TabIndex        =   9
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2685
      MaxLength       =   2
      TabIndex        =   8
      Top             =   540
      Width           =   375
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   312
      Left            =   6336
      TabIndex        =   1
      Top             =   2844
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
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額:"
      Height          =   180
      Left            =   3756
      TabIndex        =   33
      Top             =   3240
      Width           =   768
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5400
      TabIndex        =   32
      Top             =   2880
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   80
      X2              =   7900
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   80
      X2              =   7900
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請書類別:  "
      Height          =   180
      Left            =   240
      TabIndex        =   31
      Top             =   3240
      Width           =   1044
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   2880
      Width           =   948
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3270
      TabIndex        =   29
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3270
      TabIndex        =   28
      Top             =   2310
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   60
      TabIndex        =   27
      Top             =   2310
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   26
      Top             =   570
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3270
      TabIndex        =   25
      Top             =   1980
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   240
      TabIndex        =   24
      Top             =   1980
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   23
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   22
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3270
      TabIndex        =   21
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   20
      Top             =   1260
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1050
      TabIndex        =   19
      Top             =   900
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
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
      Top             =   900
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1710
      TabIndex        =   17
      Top             =   1230
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
      Left            =   1050
      TabIndex        =   16
      Top             =   1980
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
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
      Top             =   1980
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1050
      TabIndex        =   14
      Top             =   2310
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
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
      Top             =   2310
      Width           =   2100
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3704;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm06010307_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/5 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Public strReceiveNo As String
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, cp() As String, m_CP110 As String, m_AgentName As String
Public m_CP118isY As Boolean 'Add By Sindy 2019/7/23 是否為電子送件申請書:True.是
Dim intWhere As Integer
Dim m_CaseNo As String
Dim m_SendDate As String, m_SendWord As String, m_SendNumber As String 'Add By Sindy 2019/7/23


'Add by Morgan 2005/8/8
Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
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

'Add by Morgan 2005/8/8
Private Function FormSave() As Boolean
Dim strCon As String

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   'Add By Sindy 2019/7/23
   If m_CP118isY = True Then
      If m_CP118isY = True Then
         cp(118) = "A"
      Else
         cp(118) = ""
      End If
      strCon = strCon & ",cp118=" & CNULL(cp(118))
   End If
   strCon = strCon & ",cp84=" & Val(txtCP84) '發文規費
   '2019/7/23 END
   
   'Modify By Sindy 2023/6/27
   If Me.Option1(1).Value = True Then '因客戶/本所誤繕需更改
      '需請款:則更改進度檔預設的不請款"N"請刪除
      If Me.Option2(0).Value = True Then
         strCon = strCon & ",cp20=null"
      Else
         '不請款,在變更申請書作業中處理....
      End If
   End If
   '2023/6/27 END
   
'   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(m_CP110) & strCon & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql
'   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

Private Sub cmdok_Click(Index As Integer)
Dim strTmp As String
Dim strFolder As String, strFileName As String
   
   Select Case Index
      Case 0
         'Add by Morgan 2005/8/8
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
         
         'Add by Sindy 2019/7/23 電子送件申請書
'         If m_CP118isY = True Then
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
               strFolder = PUB_Getdesktop
            Else
               strFolder = FCP電子送件檔案存放路徑
            End If
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            'Modify By Sindy 2023/6/27
            If Option1(0).Value = True Then '因智慧局來函誤繕
            '2023/6/27 END
               '1.基本資料
               StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False
               NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & m_CaseNo & ".contact"
               Call PUB_MakeDoc(strExc(9), strFileName)
               
               '2.申請書
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "一般事項申復申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
               frm060103_1.Show
               ' 90.08.27 modify by louis (回到原畫面要清除畫面)
               frm060103_1.ClearForm
            
            'Modify By Sindy 2023/6/27
            Else '因客戶/本所誤繕需更改; 開啟變更申請書作業畫面
               Set frm06010303_1.oParent = frm060103_1
               frm06010303_1.m_CP118isY = "Y" '電子送件申請書
               If Me.Option2(1).Value = True Then
                  frm06010303_1.m_CP20isN = True '不請款
               Else
                  frm06010303_1.m_CP20isN = False
               End If
               frm06010303_1.Caption = "各式申請書-電子送件-變更"
               frm06010303_1.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 61
               Unload Me
               '2023/6/27 END
            End If
'         Else
'         '2019/7/23 END
'
'            Select Case Text6
'               Case "1" '申請更改審定書的申請日 1
'                  strTmp = "06" '"01" Modify By Sindy 2019/7/23
'               Case "2" '申請更改審定書的申請人 2
'                  strTmp = "02"
'               Case "3" '申請更改審定書的專利名稱 3
'                  strTmp = "03"
'               Case "4" '申請更改審定書的期限 4
'                  strTmp = "04"
'               'Modify By Sindy 2022/6/8 Mark
''               Case "5" '申請更改實審函的申請日 5
''                  strTmp = "05"
'            End Select
'            strLetterDate = Text5.Text
'            NowPrint strReceiveNo, "01", strTmp, False, strUserNum
'         End If
'         frm060103_1.Show
'         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
'         frm060103_1.ClearForm
      Case 1
         frm060103_1.Show
      Case 2
         Unload frm060103_1
   End Select
   Unload Me
End Sub

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
   
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Add By Sindy 2019/7/23
   ReadPatent
   'Add by Morgan 2005/8/8
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
   Set frm06010307_1 = Nothing
End Sub

Private Sub ReadPatent()
Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
   
   'Add By Sindy 2019/7/23
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
   '2019/7/23 END
   
   For Each Lbl In Label12
      Lbl = ""
   Next
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
      "cp43,CP110 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
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
   End If
   End With
   
   'Add By Sindy 2019/7/23
   '來函文號:
   strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & cp(9) & "' AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) DESC"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("ED08")) Then
         m_SendDate = RsTemp("ED08") - 19110000
         If Not IsNull(RsTemp("cp08")) Then
            strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
            m_SendWord = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
            strExc(0) = Replace(strExc(0), m_SendWord & "字第", "")
            m_SendNumber = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
         End If
      End If
   End If
   
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

'Private Sub Text6_GotFocus()
'  TextInverse Text6
'End Sub
'
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   If (KeyAscii < 49 Or KeyAscii > 52) And KeyAscii <> 8 Then
'      KeyAscii = 0
'      Beep
'   End If
'End Sub

'Add by Morgan 2005/8/8
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

'Add By Sindy 2019/7/23
'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
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
   
   '辦理依據
   If m_SendDate <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(m_SendDate) & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & m_SendWord & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & m_SendNumber & "')"
   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','因「　　　」原因，申請/聲明「　　　    」。')"
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

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
