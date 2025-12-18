VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010607_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "延期受理輸入"
   ClientHeight    =   5748
   ClientLeft      =   156
   ClientTop       =   960
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8952
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1200
      Width           =   7065
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2040
      Width           =   1152
   End
   Begin VB.TextBox Text16 
      Height          =   270
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2340
      Width           =   1152
   End
   Begin VB.TextBox Text17 
      Height          =   270
      Left            =   5550
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2340
      Width           =   1272
   End
   Begin VB.Frame Frame2 
      Height          =   552
      Left            =   4110
      TabIndex        =   40
      Top             =   1440
      Width           =   4332
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   6
         Top             =   200
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "          月"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   8
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   780
         MaxLength       =   2
         TabIndex        =   4
         Top             =   200
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                        日"
         Height          =   225
         Index           =   2
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到           天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   552
      Left            =   1260
      TabIndex        =   39
      Top             =   1440
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   5550
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2040
      Width           =   1272
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   8250
      TabIndex        =   27
      Top             =   60
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6660
      TabIndex        =   25
      Top             =   60
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7320
      TabIndex        =   26
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   20
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   19
      Top             =   150
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   18
      Top             =   150
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1260
      MaxLength       =   3
      TabIndex        =   17
      Top             =   150
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4860
      TabIndex        =   16
      Top             =   150
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1575
      Left            =   1260
      TabIndex        =   13
      Top             =   2670
      Width           =   7095
      _ExtentX        =   12510
      _ExtentY        =   2773
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.TextBox Text31 
      Height          =   735
      Left            =   1260
      TabIndex        =   15
      Top             =   4965
      Width           =   7095
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "10557;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text29 
      Height          =   660
      Left            =   1260
      TabIndex        =   14
      Top             =   4275
      Width           =   7095
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "10557;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1260
      TabIndex        =   28
      Top             =   450
      Width           =   5175
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "13652;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label46 
      Caption         =   "案件備註:"
      Height          =   180
      Left            =   300
      TabIndex        =   50
      Top             =   5010
      Width           =   855
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   300
      TabIndex        =   49
      Top             =   1260
      Width           =   765
   End
   Begin VB.Label Label22 
      Caption         =   "來函期限:"
      Height          =   180
      Left            =   300
      TabIndex        =   48
      Top             =   1635
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "法定期限:"
      Height          =   255
      Left            =   4470
      TabIndex        =   47
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   300
      TabIndex        =   46
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label Label25 
      Caption         =   "承辦期限:"
      Height          =   255
      Left            =   4470
      TabIndex        =   45
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label26 
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   300
      TabIndex        =   44
      Top             =   2370
      Width           =   915
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   7
      Left            =   2445
      TabIndex        =   43
      Top             =   2370
      Width           =   1170
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label43 
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   300
      TabIndex        =   42
      Top             =   4290
      Width           =   855
   End
   Begin VB.Label Label37 
      Caption         =   "本案期限:"
      Height          =   180
      Left            =   300
      TabIndex        =   41
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y:閉卷)"
      Height          =   180
      Index           =   4
      Left            =   7920
      TabIndex        =   38
      Top             =   960
      Width           =   645
   End
   Begin VB.Label lblPA57 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   7200
      TabIndex        =   37
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷"
      Height          =   180
      Index           =   3
      Left            =   6240
      TabIndex        =   36
      Top             =   960
      Width           =   720
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   6
      Left            =   4860
      TabIndex        =   35
      Top             =   990
      Width           =   1230
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   1260
      TabIndex        =   34
      Top             =   990
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3651;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日"
      Height          =   180
      Index           =   2
      Left            =   3780
      TabIndex        =   33
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收 文 號:"
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   32
      Top             =   990
      Width           =   675
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   4860
      TabIndex        =   31
      Top             =   780
      Width           =   1230
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   30
      Top             =   780
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3651;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   29
      Top             =   780
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   300
      TabIndex        =   24
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3780
      TabIndex        =   23
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   300
      TabIndex        =   22
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   3780
      TabIndex        =   21
      Top             =   780
      Width           =   585
   End
End
Attribute VB_Name = "frm06010607_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/23 Form2.0已修改
'Create By Sindy 2016/8/11 參考frm06010604_3撰寫
Option Explicit

Dim strReceiveNo As String, strTemp As String
Dim pa() As String, cp() As String
Dim intWhere As Integer, intLastRow As Integer
Dim m_strCP09ByCheng As String '總收文號
Dim m_CP16 As String       '預設請款金額
Dim m_blnClosed As Boolean '是否閉卷


Private Sub cmdok_Click(Index As Integer)
Dim ii As Integer
   
   Select Case Index
      Case 0
         If Me.Text14(0).Text <> "" Then
            If Len(Me.Text14(0).Text) = 8 Then
               If Val(Me.Text14(0).Text) < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日期!!!", vbExclamation
                  If Me.Text14(0).Enabled Then
                      Me.Text14(0).SetFocus
                      Me.Text14(0).SelStart = 0
                      Me.Text14(0).SelLength = Len(Me.Text14(0).Text)
                      Exit Sub
                  End If
               End If
            Else
               If Val(Me.Text14(0).Text) + 19110000 < ServerDate Then
                  MsgBox "本所期限不可小於系統日期!!!", vbExclamation
                  If Me.Text14(0).Enabled Then
                      Me.Text14(0).SetFocus
                      Me.Text14(0).SelStart = 0
                      Me.Text14(0).SelLength = Len(Me.Text14(0).Text)
                      Exit Sub
                  End If
               End If
            End If
         End If
         '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
         If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 Then
            If Val(Me.Text14(0).Text) < Val(Me.Text17.Text) Then
               MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
               Exit Sub
            End If
         End If
         '若輸入的來函性質為通知補文件或延期受理時
         For ii = 1 To Me.MSHFlexGrid1.Rows - 1
            '若有勾選本案期限
            If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
               If Me.Text14(0).Text = "" Or Me.Text14(1).Text = "" Then
                  MsgBox "本所期限及法定期限不可空白!!!", vbExclamation + vbOKOnly
                  Exit Sub
               End If
            End If
         Next ii
         
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Add by Sindy 2021/11/23 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         '加漏斗
         Screen.MousePointer = vbHourglass
         
         If FormSave = False Then
            Screen.MousePointer = vbDefault
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         End If
         
         Screen.MousePointer = vbDefault
         
         Unload Me
         frm06010607_2.FormConfirm
      Case 1
         frm06010607_2.Show
         Unload Me
      Case 2
         Unload frm06010607_2
         Unload frm06010607_1
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
Dim i As Integer
Dim Ncp() As String
Dim BlnCheck As Boolean '判斷是否有勾選本案期限
Dim strDate1 As String '本所期限
Dim strDate2 As String '法定期限
   
   FormSave = True
   
On Error GoTo CheckingErr
   
   ReDim Ncp(1 To TF_CP) As String
   cnnConnection.BeginTrans
   
   Ncp(1) = cp(1)
   Ncp(2) = cp(2)
   Ncp(3) = cp(3)
   Ncp(4) = cp(4)
   Ncp(5) = Label3(6)
   Ncp(6) = Text14(0)
   Ncp(7) = Text14(1)
   Ncp(8) = Text9
   Ncp(9) = "C" & CompAutoNumberYear(GetTaiwanThisYear)
   Ncp(10) = "1004" '延期受理
   
   '智權人員存國家檔FCP承辦智權人員
   Ncp(13) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
   Ncp(12) = GetSalesArea(Ncp(13))
   Ncp(14) = Text16
   Ncp(48) = Text17
   
   If Ncp(20) = "" Then
      Ncp(16) = Val(m_CP16)
      Ncp(17) = 0
      Ncp(18) = Val(m_CP16) / 1000
   End If
   
   Ncp(32) = "N"
   Ncp(43) = cp(9)
   Ncp(64) = Text29
   Ncp(27) = strSrvDate(2)
  
   m_strCP09ByCheng = Empty
   If Not ClsPDSaveNewCaseProgressDatabase("C", Ncp, intWhere, m_strCP09ByCheng) Then
      cnnConnection.RollbackTrans
      FormSave = False
      Exit Function
   End If
   
   '無序號,則更新案件進度檔的本所期限及法定期限
   '有序號,則更新下一程序檔的本所期限及法定期限
   BlnCheck = False
   '本所期限
   strDate1 = DBDATE(Me.Text14(0).Text)
   '法定期限
   strDate2 = DBDATE(Me.Text14(1).Text)
   With Me.MSHFlexGrid1
      For i = 1 To .Rows - 1
         If LCase("" & .TextMatrix(i, 0)) = "v" Then
            If Val(.TextMatrix(i, 8)) > 0 Then
               strSql = "UPDATE NEXTPROGRESS SET NP08='" & strDate1 & "'" & _
                        ",NP09 = '" & strDate2 & "'" & _
                        " WHERE NP22=" & .TextMatrix(i, 8) & " and np01='" & .TextMatrix(i, 9) & "'"
               cnnConnection.Execute strSql
            Else
               strSql = "UPDATE CaseProgress SET CP06='" & strDate1 & "'" & _
                        ",CP07='" & strDate2 & "'" & _
                        " WHERE CP09='" & .TextMatrix(i, 9) & "'"
               cnnConnection.Execute strSql
            End If
         End If
      Next i
   End With
   
   cnnConnection.CommitTrans
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
End Function

Private Sub Form_Initialize()
ReDim pa(1 To TF_PA) As String
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
Dim strTmp As String
   
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm06010607_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      ReadPatent
   End With
   Combo1.ListIndex = 0
   
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   Text9.Text = "（" & strTmp & "）智專一（二）字第號"
End Sub

Private Sub ReadPatent()
Dim Lbl As Object
Dim bolTmp As Boolean
Dim str1003CP09 As String 'Add By Sindy 2016/8/17
Dim Cancel As Boolean 'Add By Sindy 2016/8/17
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   Label3(6) = frm06010607_1.Text5
   Label3(5) = strReceiveNo
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   '是否閉卷
   lblPA57 = Empty
   If pa(1) = "FCP" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text31 = pa(91)
         
         If pa(57) = "Y" Then
            m_blnClosed = True
         Else
            m_blnClosed = False
         End If
         lblPA57 = pa(57)
      End If
   Else
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text31 = pa(18)
         
         If pa(15) = "Y" Then
            m_blnClosed = True
         Else
            m_blnClosed = False
         End If
         lblPA57 = pa(15)
      End If
   End If
   '下一程序名稱帶出來,在相關人後加備註欄
   'Modify By Sindy 2016/8/16 + 本案期限：帶出下一程序檔未處理期限，但要剔除備註有優先權的資料，
   '                            同時再帶進度檔未發文未取消收文且有期限的進度，也要剔除備註有優先權的資料
   strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & " NP08," & SQLDate("NP09") & " NP09,NP13," & _
      "NP14,NP15," & SQLDate("NP11") & ",NP22,NP01 FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
      "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND (NP06<>'Y' OR NP06 IS NULL) AND NP02=CPM01(+) AND NP07=CPM02(+) AND (instr(NP15,'優先權')=0 or NP15 is null)"
   strExc(0) = strExc(0) & " union " & _
               "SELECT '',CPM03," & SQLDate("CP06") & "," & SQLDate("CP07") & ",CP08," & _
      "'' NP14,CP64,'',0 NP22,CP09 FROM CASEPROGRESS,CASEPROPERTYMAP " & _
      "WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND CP27 IS NULL AND CP57 IS NULL AND CP07 IS NOT NULL AND CP01=CPM01(+) AND CP10=CPM02(+) AND (instr(CP64,'優先權')=0 or CP64 is null)"
   'strExc(0) = strExc(0) & " order by 2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   cp(9) = strReceiveNo
   If ClsPDReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(10) <> "" Then
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         If ClsPDGetCaseProperty(cp(1), cp(10), strExc(0), bolTmp) Then Label3(1) = strExc(0)
      End If
   End If
   'Add By Sindy 2016/8/17 若為新申請案之第一次通知補文件之延期-補文件，
   '                       則法定期限預設為指定日期且為案件之申請日+6個月
   strExc(0) = "SELECT c3.cp05,c3.cp09 FROM CASEPROGRESS c1,CASEPROGRESS c2,CASEPROGRESS c3" & _
      " WHERE c1.cp09='" & cp(9) & "'" & _
      " and c1.cp43=c2.cp09(+) and c2.cp10='202'" & _
      " and c2.cp43=c3.cp09(+) and c3.cp10='1003'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      str1003CP09 = RsTemp.Fields(1)
      strExc(0) = "SELECT c2.cp09 FROM CASEPROGRESS c1,CASEPROGRESS c2" & _
         " WHERE c1.cp01='" & pa(1) & "'" & _
         " and c1.cp02='" & pa(2) & "'" & _
         " and c1.cp03='" & pa(3) & "'" & _
         " and c1.cp04='" & pa(4) & "'" & _
         " and c1.cp05<=" & RsTemp.Fields(0) & " and c1.cp10='1003' and c1.cp43=c2.cp09(+)" & _
         " and instr('" & NewCasePtyList & "',c2.cp10)>0" & _
         " order by SQLDatet2(c2.CP05) asc,c2.CP66 asc,c2.CP67 asc,c2.CP09 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         If str1003CP09 = RsTemp.Fields(0) Then
            If Val(pa(10)) > 0 Then
               Option4(2).Value = True
               '指定日期且為案件之申請日+6個月
               Text12 = TransDate(CompDate(1, 6, TransDate(pa(10), 2)), 1)
               Cancel = False
               Call Text12_Validate(Cancel)
               If Cancel = True Then
                  Text12.SetFocus
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
   '2016/8/17 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010607_3 = Nothing
End Sub

Private Sub Text10_GotFocus()
   InverseTextBox Text10
End Sub

Private Sub Text11_GotFocus()
   InverseTextBox Text11
End Sub

Private Sub Text12_GotFocus()
   InverseTextBox Text12
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
      MsgBox "來函期限不可空白 !", vbCritical
      Cancel = True
   Else
      If ChkDate(Text12) Then
         If Val(Text12) < Val(strSrvDate(2)) Then
            MsgBox "來函期限不可小於系統日 !", vbCritical
            Cancel = True
         Else
            Text14(1) = Text12
            Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
            If CompDate(1, 1, GetTodayDate) < TransDate(Text12, 2) Then
                Text14(0) = TransDate(CompDate(2, -4, TransDate(Text14(1), 2)), 1)
            End If
            Text14(0) = TransDate(PUB_GetWorkDay1(DBDATE(Text14(0)), True), 1) 'Added by Lydia 2025/11/12 改抓最近工作天
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub GetTime()
Dim i As Integer
   If Option4(0).Value = True Then
      If Val(Text10) > 0 Then
         Text14(1) = TransDate(CompDate(2, Val(Text10), TransDate(Label3(6), 2)), 1)
         If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      End If
   ElseIf Option4(1).Value = True Then
      If Val(Text11) > 0 Then
         Text14(1) = TransDate(CompDate(1, Val(Text11), TransDate(Label3(6), 2)), 1)
         If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
   End If
   If Text14(1) <> "" Then
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      Text14(0) = TransDate(PUB_GetWorkDay1(CompDate(2, i, TransDate(Text14(1), 2)), True), 1)
   End If
End Sub

Private Sub Text14_GotFocus(Index As Integer)
   InverseTextBox Text14(Index)
End Sub

Private Sub Text14_Validate(Index As Integer, Cancel As Boolean)
   If Text14(Index) <> "" Then
      If Not ChkDate(Text14(Index)) Then
         Cancel = True
      Else
         '若有輸入本所期限時, 不可小於系統日
         If Index = 0 Then
            If Len(Me.Text14(0).Text) = 8 Then
               If Val(Me.Text14(0).Text) < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            ElseIf Len(Me.Text14(0).Text) = 7 Or Len(Me.Text14(0).Text) = 6 Then
               If Val(Me.Text14(0).Text) + 19110000 < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            End If
         End If
         
         If Index = 1 Then
            If Not ChkRange(Text14(0), Text14(1), "本所期限、法定期限") Then
               Cancel = True
            Else
               If ClsLawChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                  If Text14(0) <> TransDate(strExc(1), 1) Then
                     If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                        frm06010607_1.Show
                        Unload frm06010607_2
                        Unload Me
                     Else
                        Text14(0) = ""
                        Text14(1) = ""
                     End If
                  ElseIf Text14(1) <> TransDate(strExc(2), 1) Then
                     If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                        frm06010607_1.Show
                        Unload frm06010607_2
                        Unload Me
                     Else
                        Text14(0) = ""
                        Text14(1) = ""
                     End If
                  End If
               Else
                  If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Cancel = True
               End If
            End If
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text14(Index)
End Sub

Private Sub Text16_GotFocus()
   InverseTextBox Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTmp As String
   
   Cancel = False
   Label3(7) = Empty
   If IsEmptyText(Text16) = False Then
      strTemp = Empty
      strTemp = GetStaffName(Text16)
      Label3(7) = strTemp
      If IsEmptyText(strTemp) Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的承辦人"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Text16_GotFocus
      End If
   End If
End Sub

Private Sub Text17_GotFocus()
   InverseTextBox Text17
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   If Text17 <> "" Then
      If ChkWorkDay(TransDate(Text17, 2)) Then
         If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 And Val(Text17) > Val(Text14(0)) Then
            MsgBox "承辦期限不可大於本所期限，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         MsgBox "承辦期限不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text17
End Sub

Private Sub Text31_GotFocus()
   InverseTextBox Text31
End Sub

Private Sub GridHead()
Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1500: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1500: .Text = "備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1500: .Text = "解除期限日期"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 0 'NP22
      .col = 9: .ColWidth(9) = 0 '總收文號
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Integer
   
   GridClick MSHFlexGrid1, intLastRow, 0, 1
End Sub

Private Sub Text9_GotFocus()
Dim intPos As Integer
   OpenIme
   '當來函性質為"1601"或"1604"時, 將游標設定在機關文號欄的"第"的後面, 其餘則放在"專"的後面
   With Me.Text9
      If Len("" & .Text) > 0 Then
         intPos = InStr("" & .Text, "專")
         If intPos > 0 Then
            .SelStart = intPos
            .SelLength = 0
         End If
      End If
   End With
End Sub

Private Sub Text9_LostFocus()
   CloseIme
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim bFind As String

TxtValidate = False

'If (cp(43) = "" Or cp(43) > "C") Then
'   bFind = False
'   For ii = 1 To Me.MSHFlexGrid1.Rows - 1
'      If Me.MSHFlexGrid1.TextMatrix(ii, 0) = "v" Then
'         bFind = True
'         Exit For
'      End If
'   Next ii
'   If bFind = False Then
'      MsgBox "請於延期程序之相關總收文號欄位補輸原下一程序之A類總收文號!!!", vbExclamation + vbOKOnly
'      Cancel = True
'     Exit Function
'   End If
'Add By Sindy 2016/8/17
'Else
   bFind = False
   For ii = 1 To Me.MSHFlexGrid1.Rows - 1
      If Me.MSHFlexGrid1.TextMatrix(ii, 0) = "v" Then
         bFind = True
         Exit For
      End If
   Next ii
   If bFind = False Then
      MsgBox "本案期限至少要勾選一筆期限資料!!!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Function
   End If
'2016/8/17 END
'End If

If Me.Text10.Enabled = True Then
   Cancel = False
   Text10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text11.Enabled = True Then
   Cancel = False
   Text11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text12.Enabled = True Then
   Cancel = False
   Text12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

For Each objTxt In Text14
   Cancel = False
      If Text14(objTxt.Index) <> "" Then
         If Not ChkDate(Text14(objTxt.Index)) Then
            Cancel = True
         Else
            '若有輸入本所期限時, 不可小於系統日
            If objTxt.Index = 0 Then
               If Len(Me.Text14(0).Text) = 8 Then
                  If Val(Me.Text14(0).Text) < strSrvDate(1) Then
                     MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                     Cancel = True
                  End If
               ElseIf Len(Me.Text14(0).Text) = 7 Or Len(Me.Text14(0).Text) = 6 Then
                  If Val(Me.Text14(0).Text) + 19110000 < strSrvDate(1) Then
                     MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                     Cancel = True
                  End If
               End If
            End If
            If objTxt.Index = 1 Then
               If Not ChkRange(Text14(0), Text14(1), "本所期限、法定期限") Then
                  Cancel = True
               Else
                  If ClsLawChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                     If Text14(0) <> TransDate(strExc(1), 1) Then
                        If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                           Cancel = True
                        End If
                     ElseIf Text14(1) <> TransDate(strExc(2), 1) Then
                        If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                           Cancel = True
                        End If
                     End If
                  Else
                     If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Cancel = True
                  End If
               End If
            End If
         End If
      End If
   If Cancel = True Then
      Exit Function
   End If
Next

If Me.Text16.Enabled = True Then
   Cancel = False
   Text16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text17.Enabled = True Then
   Cancel = False
   Text17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'檢查機關文號
If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Text14(1) <> "" Then
   If DBDATE(Text14(1)) > CompDate(1, 6, Label3(6)) Then
      MsgBox "法定期限大於來函收文日6個月!!", vbCritical
      Exit Function
   End If
End If

If Check1004 = False Then Exit Function

TxtValidate = True
End Function

Private Sub Text9_Validate(Cancel As Boolean)
   If CheckLengthIsOK(Text9, Text9.MaxLength) = False Then
      Cancel = True
      Text9.SetFocus
      Text9_GotFocus
   End If
End Sub

'檢查申復延期受理期限
Private Function Check1004() As Boolean
   '205.申復
   strExc(0) = " select np08, np09, 1 srt from caseprogress a, nextprogress where a.cp09='" & cp(9) & "' and a.cp10='404' and np01(+)=a.cp43 and np07='205'" & _
      " union all select b.cp06, b.cp07, 2 from caseprogress a, caseprogress b where a.cp09='" & cp(9) & "' and a.cp10='404' and b.cp09(+)=a.cp43 and b.cp10='205'" & _
      " order by srt"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '法限相同時所限設為原所限
      If DBDATE(Text14(1)) = RsTemp("np09") Then
         strExc(1) = TransDate(RsTemp("np08"), 1)
         If strExc(1) <> Text14(0) Then
            MsgBox "本次來函為申復的延期受理，本所期限將更新為延期發文時的期限 " & strExc(1) & " !!", vbInformation, "申復延期受理函檢查"
            Text14(0) = strExc(1)
         End If
      '超過延期發文法限+10天
      ElseIf DBDATE(Text14(1)) > CompDate(2, 10, RsTemp("np09")) Then
         If MsgBox("法定期限超過申復延期發文時的期限+10天!!是否確定要繼續?", vbYesNo + vbExclamation + vbDefaultButton2, "申復延期受理函檢查") = vbNo Then
            Exit Function
         End If
      End If
   End If
   Check1004 = True
End Function
