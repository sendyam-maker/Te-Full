VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_19 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內往來記錄資料查詢"
   ClientHeight    =   5060
   ClientLeft      =   1440
   ClientTop       =   2310
   ClientWidth     =   9380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5060
   ScaleWidth      =   9380
   Begin VB.ListBox lstAtt 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      ItemData        =   "frm100101_19.frx":0000
      Left            =   1065
      List            =   "frm100101_19.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3210
      Width           =   7440
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   255
      Left            =   8550
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3330
      Width           =   735
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "下一筆(&N)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   7470
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   105
      Width           =   885
   End
   Begin VB.TextBox txtCOR 
      Height          =   276
      Index           =   1
      Left            =   1065
      MaxLength       =   9
      TabIndex        =   0
      Top             =   375
      Width           =   1092
   End
   Begin VB.TextBox txtCOR 
      Height          =   264
      Index           =   2
      Left            =   1065
      MaxLength       =   9
      TabIndex        =   1
      Top             =   720
      Width           =   1125
   End
   Begin VB.TextBox txtCOR 
      Height          =   276
      Index           =   3
      Left            =   1065
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1050
      Width           =   1092
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8400
      TabIndex        =   8
      Top             =   105
      Width           =   800
   End
   Begin VB.TextBox txtCF 
      Height          =   264
      Index           =   6
      Left            =   4620
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3870
      Visible         =   0   'False
      Width           =   4560
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4650
      Width           =   6225
      VariousPropertyBits=   671107103
      Size            =   "10980;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCOR05 
      Height          =   1400
      Left            =   1065
      TabIndex        =   4
      Top             =   1695
      Width           =   7755
      VariousPropertyBits=   -1462747109
      ScrollBars      =   2
      Size            =   "13679;2469"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCOR04 
      Height          =   300
      Left            =   1065
      TabIndex        =   3
      Top             =   1365
      Width           =   7755
      VariousPropertyBits=   675299355
      MaxLength       =   200
      Size            =   "13679;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2190
      TabIndex        =   10
      Top             =   1080
      Width           =   6600
      Size            =   "11642;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "附件："
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "往來對象："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "往來日期："
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   14
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "記錄編號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   435
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "主　　旨："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1395
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "內　　容："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1725
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/22 改成Form2.0(lbl1,textCUID,txtCOR(4)改為txtCOR04,txtCOR(5)改為txtCOR05)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

Public cmdState As Integer
Dim strTmp As String
Dim rsContact As ADODB.Recordset
Dim m_bReadGrid As Boolean
Dim oText As TextBox
Dim idx As Integer
Private Const cTableName As String = "CONTACTFILE" 'Add By Sindy 2020/5/17 指定FTP資料夾名稱
Public m_pub_QL05 As String 'Add By Sindy 2025/8/27 只記錄於此Form


'Add By Sindy 2020/5/17
Private Sub cmdOpenAtt_Click()
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long

   If lstAtt.Text = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      If txtCF(6).Text <> "" Then
         tmpArr = Empty
         tmpArr = Split(txtCF(6).Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, " (") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   textCUID.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_19 = Nothing
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 2
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
   End Select
End Sub

Sub StrMenu()
   Dim strKey  As String
   strKey = Me.Tag
   pub_QL05 = m_pub_QL05 & ";記錄編號：" & Me.Tag & "(國內往來記錄資料)" 'Add By Sindy 2025/8/13
   
   strExc(0) = "select * from contactrecord1 where cor01='" & strKey & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pub_QL04 <> "" Then InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2025/8/13
      'Add By Sindy 2011/01/03 檢查國內外權限
      If CheckSR12(RsTemp.Fields("cor03")) = False Then
         Screen.MousePointer = vbDefault
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
      
      ShowRecord RsTemp
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset)
Dim rsRec As ADODB.Recordset
Dim CUID(1 To 6) As String
'Add By Sindy 2020/5/17
Dim AdoRs As New ADODB.Recordset
Dim strCF02 As String
'2020/5/17 END
Dim strName As String 'Add By Sindy 2020/5/21
   
   ClearField
   SetCtrlReadOnly True
   Set rsRec = p_Rst.Clone
   With rsRec
      If .RecordCount > 0 Then
         '2008/12/9 ADD BY SONIA 加判斷語文權限
         'Modified by Lydai 2019/09/10 +建檔人
         'If GetCustData(Mid(.Fields("COR03"), 1, 8)) = False Then
         'Modify By Sindy 2020/5/21
         'If GetCustData(Mid(.Fields("COR03"), 1, 8), "" & .Fields("COR06")) = False Then
         lbl1 = "" 'Add By Sindy 2020/5/21
         'Modified by Lydia 2020/11/03 建檔人改成必傳(COR06或ADD新增用)
         'If PUB_GetCustData_frm100101_19(.Fields("COR03"), , strName) = False Then
         ''2020/5/21 END
         If PUB_GetCustData_frm100101_19(.Fields("COR03"), .Fields("COR06"), True, strName) = False Then
         
            Screen.MousePointer = vbDefault
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
            Exit Sub
         'Modify By Sindy 2020/5/21
         Else
            lbl1 = strName
         '2020/5/21 END
         End If
         '2008/12/9 END
        
         For Each oText In txtCOR
            idx = oText.Index
            oText.Text = "" & .Fields("COR" & Format(idx, "0#"))
            '2010/8/20 add by sonia 往來日期改顯示民國年月日
            If idx = 2 Then
               oText.Text = ChangeTStringToTDateString(ChangeWStringToTString("" & .Fields("COR" & Format(idx, "0#"))))
            End If
            '2010/8/20 end
         Next
         'add by sonia 2022/1/22 txtCOR(4)改為txtCOR04,txtCOR(5)改為txtCOR05
         txtCOR04.Text = "" & .Fields("COR04")
         txtCOR05.Text = "" & .Fields("COR05")
         'end 2022/1/22
         
'         '往來對象
'         'Modified by Lydia 2019/09/10 +建檔人COR06
'         GetCustData txtCOR(3), "" & .Fields("COR06")
         
         'Add By Sindy 2020/5/17
         strExc(0) = "SELECT cf02,cf06,cf07 FROM ContactFile where CF01='" & txtCOR(1) & "'"
         intI = 1
         Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            AdoRs.MoveFirst
            Do While Not AdoRs.EOF
               strCF02 = strCF02 & "," & AdoRs.Fields("cf02") & IIf("" & AdoRs.Fields("cf07") <> "", " (" & AdoRs.Fields("cf07") & " KB)", "")
               txtCF(6) = txtCF(6) & "," & AdoRs.Fields("cf06")
               AdoRs.MoveNext
            Loop
            strCF02 = Mid(strCF02, 2)
            txtCF(6) = Mid(txtCF(6), 2)
         Else
            strCF02 = ""
            txtCF(6) = ""
         End If
         '附件路徑
         If Not IsNull(strCF02) Then
            SetList lstAtt, strCF02
         End If
         '2020/5/17 END
         
         CUID(1) = "" & .Fields("COR06")
         CUID(2) = "" & .Fields("COR07")
         CUID(3) = "" & .Fields("COR08")
         CUID(4) = "" & .Fields("COR09")
         CUID(5) = "" & .Fields("COR10")
         CUID(6) = "" & .Fields("COR11")
      End If
   End With
   UpdateCUID CUID, textCUID
   Set AdoRs = Nothing 'Add By Sindy 2020/5/17
End Sub

Private Sub ClearField()
   Dim oLabel As LABEL
   For Each oText In txtCOR
      oText.Text = Empty
   Next
   'add by sonia 2022/1/22 txtCOR(4)改為txtCOR04,txtCOR(5)改為txtCOR05
   txtCOR04 = ""
   txtCOR05 = ""
   'end 2022/1/22
   lbl1 = Empty
   textCUID = ""
   'Add By Sindy 2020/5/17
   lstAtt.Clear
   For Each oText In txtCF
      oText.Text = Empty
   Next
   '2020/5/17 END
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtCOR
      oText.Locked = bLocked
   Next
   'add by sonia 2022/1/22 txtCOR(4)改為txtCOR04,txtCOR(5)改為txtCOR05
   txtCOR04.Locked = bLocked
   txtCOR05.Locked = bLocked
   'end 2022/1/22
End Sub

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

' 更新 Create 及 Update 的人
'modify by sonia 2022/1/22 As TextBox=> Object
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

''Modified by Lydia 2019/09/10 +建檔人p_COR06
'Private Function GetCustData(p_stCust As String, Optional ByVal p_COR06 As String) As Boolean
'   Dim aiOrder(1 To 3) As Integer
'
'   GetCustData = False
'   Select Case Left(p_stCust, 1)
'      Case "X"
'         '2009/5/14 modify by sonia 開放M51權限
'         'modify by sonia 2015/6/5 開放01,09等級權限
'         'strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "' and cu13 ='" & strUserNum & "' "
'         'modify by sonia 2017/8/11 為判斷權限+cu13
'         'Modify by Amy 2017/08/11 cu02固定0-秀玲
'         'strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3,cu13 SAno from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "' "
'         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3,cu13 SAno from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='0' "
'         'cancel by sonia 2017/8/11 移到下面再判斷權限
'         'If Pub_StrUserSt03 <> "M51" And Pub_strUserST05 <> "01" And Pub_strUserST05 <> "09" Then
'         '   strExc(0) = strExc(0) & " and cu13 ='" & strUserNum & "' "
'         'End If
'         'end 2017/8/11
'         '2009/5/14 end
'      Case "R"
'         '2009/5/14 modify by sonia 開放M51權限
'         'modify by sonia 2015/6/5 開放01,09等級權限
'         'strExc(0) = "select '',poc03,'','',poc04 N3 from potcustomer1 where poc01='" & Left(p_stCust, 8) & "' and poc02='" & Right(p_stCust, 1) & "' and poc13 like '%" & strUserNum & "%' "
'         'modify by sonia 2017/8/11 為判斷權限+poc13
'         strExc(0) = "select '',NVL(POC03,NVL(RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26),POC27)),'','',poc04 N3,poc13 SAno from potcustomer1 where poc01='" & Left(p_stCust, 8) & "' and poc02='" & Right(p_stCust, 1) & "' "
'         'cancel by sonia 2017/8/11 移到下面再判斷權限
'         'If Pub_StrUserSt03 <> "M51" And Pub_strUserST05 <> "01" And Pub_strUserST05 <> "09" Then
'         '   strExc(0) = strExc(0) & " and poc13 like '%" & strUserNum & "%' "
'         'End If
'         'end 2017/8/11
'         '2009/5/14 end
'      Case Else
'         MsgBox "往來對象必須為 X 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
'   End Select
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   lbl1 = ""
'   If intI = 1 Then
'      For intI = 1 To 3
'         If Not IsNull(RsTemp(intI)) Then
'            lbl1 = RsTemp(intI)
'            Exit For
'         End If
'      Next
'      'add by sonia 2017/8/11 檢查權限
'      'Modified by Lydia 2019/09/10 +傳入客戶編號,建檔人
'      'Modified by Lydia 2019/10/18 取消QPGMR的判斷
'      'If CheckModifyLimit(RsTemp("SAno"), p_stCust, p_COR06) = False Then
'      If PUB_CheckModifyLimit_frm100101_19(RsTemp("SAno"), p_stCust, "") = False Then
'         Exit Function
'      End If
'      'end 2017/8/11
'
'   Else
'      MsgBox "往來對象輸入錯誤！"
'      Exit Function
'   End If
'
'   GetCustData = True
'End Function

''add by sonia 2017/8/11
''檢查主管權限
''Modified by Lydia 2019/09/10 +傳入客戶編號,建檔人
'Private Function CheckModifyLimit(p_SAno As String, bCustNo As String, bCOR06 As String) As Boolean
'Dim idx As Integer
'
'   '開放M51及01,09等級權限
'   If Pub_StrUserSt03 = "M51" Or Pub_strUserST05 = "01" Or Pub_strUserST05 = "09" Then
'      CheckModifyLimit = True
'      Exit Function
'   End If
'
'   'Add by Amy 2017/08/11 +MCTF 權限控制
'   If Left(p_SAno, 4) = "MCTF" Then
'        'MCTF人員
'        If InStr(Replace(Pub_GetSpecMan(p_SAno, False), ";", ","), strUserNum) > 0 Then
'            CheckModifyLimit = True
'            Exit Function
'        'MCTF主管
'        ElseIf InStr(Replace(Pub_GetSpecMan("MCTM", False), ";", ","), strUserNum) > 0 Then
'            CheckModifyLimit = True
'            Exit Function
'        End If
'   End If
'
'   '須為開發者或其案件主管, 方可維護此筆資料
'   'Modified by Lydia 2019/09/10 +整批匯入Create ID->QPGMR開放給全部人員查詢
'   If strUserNum = p_SAno Or bCOR06 = "QPGMR" Then
'      CheckModifyLimit = True
'      Exit Function
'   Else
'      strExc(0) = "SELECT A0908 FROM STAFF,ACC090 " & _
'                         "WHERE ST03=A0901(+) and ST01= '" & p_SAno & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If strUserNum = RsTemp(0) Then
'            CheckModifyLimit = True
'            Exit Function
'         End If
'      End If
'      'add by sonia 2017/10/17 帶人主管也可以看 82026可看X37109
'      strExc(0) = "SELECT ST52,ST53,ST54,ST55 FROM STAFF " & _
'                  "WHERE ST01= '" & p_SAno & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If strUserNum = "" & RsTemp(0) Then
'            CheckModifyLimit = True
'            Exit Function
'         ElseIf strUserNum = "" & RsTemp(1) Then
'            CheckModifyLimit = True
'            Exit Function
'         ElseIf strUserNum = "" & RsTemp(2) Then
'            CheckModifyLimit = True
'            Exit Function
'         ElseIf strUserNum = "" & RsTemp(3) Then
'            CheckModifyLimit = True
'            Exit Function
'         End If
'      End If
'      'end 2017/10/17
'   End If
'
'   'Added by Lydia 2019/09/10 若客戶存在於待活化客戶檔，開放同一所別所有人都可查詢往來記錄
'   If bCustNo <> "" And Left(bCustNo, 1) = "X" Then
'        strExc(0) = "select ocu01,st06 from oldcustomer,customer,staff " & _
'                         "where ocu01='" & Mid(ChangeCustomerL(bCustNo), 1, 8) & "' and ocu01=cu01 and cu02='0' and nvl(ocu03,0)=0 and cu13=st01(+) "
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'        If intI = 1 Then
'           If pub_strUserOffice = "" & RsTemp("st06") Then
'               CheckModifyLimit = True
'               Exit Function
'           End If
'        End If
'   End If
'
'   CheckModifyLimit = False
'   MsgBox "無查詢權限 !!!", vbInformation
'End Function
''end 2017/8/11

'Add By Sindy 2020/5/17
Private Sub lstAtt_DblClick()
   If cmdOpenAtt.Enabled = True Then
      cmdOpenAtt.Value = True
   End If
End Sub
