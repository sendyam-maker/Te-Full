VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020509 
   BorderStyle     =   1  '單線固定
   Caption         =   "TF基礎案號維護作業"
   ClientHeight    =   4176
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6924
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4176
   ScaleWidth      =   6924
   Begin VB.CommandButton Cmd1 
      Caption         =   "刪除"
      Height          =   324
      Index           =   1
      Left            =   3960
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1272
      Width           =   756
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "新增"
      Height          =   324
      Index           =   0
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1272
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   2412
      Left            =   72
      TabIndex        =   15
      Top             =   1632
      Width           =   6636
      _ExtentX        =   11705
      _ExtentY        =   4255
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      AllowUserResizing=   3
      FormatString    =   "V|TF基礎案號|申  請  國|本 所 案 號"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.TextBox txtDB 
      Height          =   300
      Index           =   1
      Left            =   4392
      MaxLength       =   3
      TabIndex        =   1
      Top             =   864
      Width           =   564
   End
   Begin VB.TextBox txtDB 
      Height          =   300
      Index           =   0
      Left            =   1368
      MaxLength       =   30
      TabIndex        =   0
      Top             =   864
      Width           =   2100
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   360
      Left            =   5856
      TabIndex        =   5
      Top             =   72
      Width           =   850
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&E)"
      Height          =   360
      Left            =   4704
      TabIndex        =   4
      Top             =   72
      Width           =   850
   End
   Begin VB.Label lblNation 
      Height          =   276
      Left            =   5016
      TabIndex        =   14
      Top             =   876
      Width           =   1380
   End
   Begin MSForms.Label lblFM2 
      Height          =   280
      Index           =   1
      Left            =   1008
      TabIndex        =   13
      Top             =   504
      Width           =   5712
      Size            =   "10075;494"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   280
      Index           =   0
      Left            =   1008
      TabIndex        =   12
      Top             =   168
      Width           =   2148
      Size            =   "3789;494"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTFB 
      Caption         =   "基礎案本所案號："
      Height          =   228
      Index           =   2
      Left            =   48
      TabIndex        =   11
      Top             =   1260
      Width           =   1476
   End
   Begin VB.Label lblTFB 
      Caption         =   "TF基礎案號數："
      Height          =   228
      Index           =   0
      Left            =   48
      TabIndex        =   10
      Top             =   900
      Width           =   1356
   End
   Begin VB.Label lblTFB 
      Caption         =   "申請國："
      Height          =   228
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   900
      Width           =   756
   End
   Begin VB.Label lblTFbaseCase 
      Caption         =   "非本所案件"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1560
      TabIndex        =   8
      Top             =   1272
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   280
      Index           =   1
      Left            =   48
      TabIndex        =   7
      Top             =   504
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   280
      Index           =   0
      Left            =   48
      TabIndex        =   6
      Top             =   168
      Width           =   900
   End
End
Attribute VB_Name = "frm020509"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2025/10/23
Option Explicit

Dim m_PrevForm As Form
Dim strCase(1 To 4) As String '本所案號
Dim oObj As Control
Dim intJ As Integer

Dim intR As Integer, strR1 As String
Dim rsRD As New ADODB.Recordset

Dim cmdState As Integer
Dim mRole As String 'U=處理 ; Q=查看
Dim mStatus As String   '是否可以輸入基礎案號
Dim mESeqNo  As String  '執行序號
Dim m_RowSeq As String, m_Flag1 As String  '暫存檔的RowSeq,FLAG1
Dim intLastRow As Integer '點選列

Public Sub SetParent(ByVal fm As Form, ByVal CNo As String, ByVal pRole As String)
   Set m_PrevForm = fm
   Call ChgCaseNo(CNo, strCase)
   mRole = pRole
End Sub

Private Sub Cmd1_Click(Index As Integer)
Dim strTmp1 As String, tmpBol As Boolean
Dim intErr As Integer
    
   intErr = -1
   If Trim(txtDB(0)) = "" Then
       MsgBox "請輸入TF基礎案號！", vbCritical
       intErr = 0
       GoTo JumpToExit
   End If
   If Trim(txtDB(1)) = "" Then
       MsgBox "請輸入申請國家！", vbCritical
       intErr = 1
       GoTo JumpToExit
   End If
   If Index = 0 Then
      If m_Flag1 <> "A" And Trim(txtDB(0)) = txtDB(0).Tag And Trim(txtDB(1)) = txtDB(1).Tag Then
         'MsgBox "未變更資料！", vbInformation
         Call ClearForm(False)
         Exit Sub
      Else
         strExc(0) = "select '1' as ord1,r001 as tfbn05, r002 as tfbn06 from rdatafactory  where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and nvl(r005,'N')<>'D' and r001='" & ChgSQL(Trim(txtDB(0))) & "' and r002='" & Trim(txtDB(1)) & "' "
         strExc(0) = strExc(0) & " Union select '2' as ord1, tfbn05, tfbn06 from TFBaseNo where tfbn01='" & strCase(1) & "' and tfbn02='" & strCase(2) & "' and tfbn03='" & strCase(3) & "' and tfbn04='" & strCase(4) & "' and tfbn05='" & ChgSQL(Trim(txtDB(0))) & "' and tfbn06='" & Trim(txtDB(1)) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "已存在相同的TF基礎案號設定！", vbCritical
            intErr = 0
            GoTo JumpToExit
         End If
      End If
   End If
   If Index = 1 Then
      If MsgBox("是否要刪除？", vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   cmdState = Index
   
   Select Case cmdState
      Case 0 '新增
          If Val(m_RowSeq) > 0 Then
             strSql = "Update rdatafactory Set R001='" & ChgSQL(Trim(txtDB(0))) & "',R002='" & Trim(txtDB(1)) & "' " & IIf(m_Flag1 = "Y", ",r005='U'", "") & _
                      ",r003='" & lblNation & "',r004='" & lblTFbaseCase & "' where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and rowseq='" & m_RowSeq & "' "
             cnnConnection.Execute strSql
          Else
             strExc(0) = "select max(rowseq) as mno from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                ' 無資料=無序號
                If Val(mESeqNo) = 0 Then
                   mESeqNo = "1"
                End If
                strSql = "insert into  rdatafactory (FORMNAME,ID,SEQNO,ROWSEQ,R001,R002,R003,R004,R005) values ('" & Me.Name & "'," & _
                           " '" & strUserNum & "', '" & mESeqNo & "', '" & Val("" & RsTemp.Fields("mno")) + 1 & "', '" & ChgSQL(Trim(txtDB(0))) & "'," & _
                           " '" & Trim(txtDB(1)) & "', '" & ChgSQL(lblNation) & "', '" & Trim(lblTFbaseCase) & "', 'A') "
                cnnConnection.Execute strSql
             End If
          End If
      Case 1 '刪除
          If Val(m_RowSeq) > 0 Then
             If m_Flag1 = "A" Then
                strSql = "delete from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and rowseq='" & m_RowSeq & "' "
                cnnConnection.Execute strSql
             Else
                strSql = "Update rdatafactory Set R005='D' where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and rowseq='" & m_RowSeq & "' "
                cnnConnection.Execute strSql
             End If
          End If
   End Select
   
   If cmdState >= 0 Then
      Call ClearForm(False)
      Call ReadData(False)
   End If
   
   Exit Sub
   
JumpToExit:
   If intErr >= 0 Then
      txtDB(intErr).SetFocus
      txtDB_GotFocus intErr
   End If
End Sub

Private Sub cmdExit_Click()
   If cmdState >= 0 Then
      If MsgBox("基礎案號設定有變更，是否放棄存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   Unload Me
End Sub

Private Sub cmdok_Click()
Dim bolConn As Boolean, strExSql As String

   If cmdState >= 0 Then  '有異動
      cmdOK.Enabled = False
      Screen.MousePointer = vbHourglass
On Error GoTo ErrHandle
      strR1 = "select r001 AS tfbn05,r002 as tfbn06,r004 AS tcase, r005 AS flag1 from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' order by rowseq "
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strR1)
      If intR = 1 Then
         rsRD.MoveFirst
         Do While Not rsRD.EOF
            strExSql = ""
            If bolConn = False And InStr("A,U,D", "" & rsRD.Fields("flag1")) > 0 Then
               bolConn = True
               cnnConnection.BeginTrans
            End If
            Select Case "" & rsRD.Fields("flag1")
               Case "A", "U"
                  If "" & rsRD.Fields("flag1") = "U" Then
                     strExSql = "delete from tfbaseno where tfbn01='" & strCase(1) & "' and tfbn02='" & strCase(2) & "' and tfbn03='" & strCase(3) & "' and tfbn04='" & strCase(4) & "' and (tfbn05,tfbn06) not in (select r001,r002 from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "') "
                     cnnConnection.Execute strExSql
                  End If
                  strExSql = "insert into tfbaseno (tfbn01,tfbn02,tfbn03,tfbn04,tfbn05,tfbn06,tfbn07,tfbn08,tfbn09) values ('" & strCase(1) & "','" & strCase(2) & "','" & strCase(3) & "','" & strCase(4) & "'," & _
                           "'" & rsRD.Fields("tfbn05") & "','" & rsRD.Fields("tfbn06") & "','" & strUserNum & "', to_char(sysdate,'YYYYMMDD'),substr(lpad(to_char(SYSDATE,'hh24miss'),6,'0'),1,4))"
                  cnnConnection.Execute strExSql
               Case "D"
                  If bolConn = False Then bolConn = True
                  strExSql = "delete from tfbaseno where tfbn01='" & strCase(1) & "' and tfbn02='" & strCase(2) & "' and tfbn03='" & strCase(3) & "' and tfbn04='" & strCase(4) & "' and tfbn05='" & ChgSQL("" & rsRD.Fields("tfbn05")) & "' and tfbn06='" & rsRD.Fields("tfbn06") & "' "
                  cnnConnection.Execute strExSql
            End Select
            If "" & rsRD.Fields("flag1") = "A" Or "" & rsRD.Fields("flag1") = "U" Then
                strExSql = PUB_GetTFbaseInfo(strCase(1), strCase(2), strCase(3), strCase(4), "" & rsRD.Fields("tfbn05"), "" & rsRD.Fields("tfbn06"), "1")
                Sleep 1000 '避免同一時分秒
            End If
            rsRD.MoveNext
         Loop
      End If
      If bolConn = True Then
         bolConn = False
         cnnConnection.CommitTrans
         PUB_SendMailCache
         cmdState = -1
      End If
      Call ClearForm(False)
      Call ReadData(True)
      cmdOK.Enabled = True
      Screen.MousePointer = vbDefault
   End If
   Unload Me '確定直接關掉---雅雯
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      cmdOK.Enabled = True
      Screen.MousePointer = vbDefault
      If bolConn = True Then cnnConnection.RollbackTrans
      MsgBox Err.Description, "存檔失敗"
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   ClearForm True
   Call SetGrid(True)
   
   If strCase(1) <> "" And strCase(2) <> "" Then
      'TF基礎案號設定：TF案未閉卷(無專用期 or 專用期未過)，卷宗性質為申請之母案案號，即TF-xxxxx0-0-00
      strExc(0) = "select * from trademark where tm01='" & strCase(1) & "' and tm02='" & strCase(2) & "' and tm03='" & strCase(3) & "' and tm04='" & strCase(4) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         lblFM2(0).Caption = strCase(1) & "-" & strCase(2) & "-" & strCase(3) & "-" & strCase(4)
         lblFM2(1).Caption = "" & RsTemp.Fields("tm05")
         If strCase(1) = "TF" And Right(strCase(2), 1) = "0" And strCase(4) = "00" And "" & RsTemp.Fields("tm28") = "1" And _
               "" & RsTemp.Fields("tm29") = "" And ("" & RsTemp.Fields("tm22") = "" Or ("" & RsTemp.Fields("tm22") <> "" And "" & RsTemp.Fields("tm22") >= strSrvDate(1))) Then
            mStatus = "Y"
         Else
            MsgBox "TF基礎案號設定僅提供：TF案未閉卷(無專用期 or 專用期未過)，" & vbCrLf & "卷宗性質為申請之母案案號，即TF-xxxxx0-0-00。", vbInformation
            mRole = "Q"
         End If
      End If
   End If
   If mRole = "Q" Then
      cmdOK.Visible = False
      Cmd1(0).Visible = False
      Cmd1(1).Visible = False
   Else
      cmdOK.Visible = True
      Cmd1(0).Visible = True
      Cmd1(1).Visible = True
   End If
   
   Call ReadData(True)
   
   If mStatus = "Y" Then
      txtDB(0).Locked = False
      txtDB(1).Locked = False
   Else
      txtDB(0).Locked = True
      txtDB(1).Locked = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If TypeName(m_PrevForm) <> "Nothing" Then
      strExc(0) = Pub_GetField("TFBaseNo", "TFBN01='" & strCase(1) & "' AND TFBN02='" & strCase(2) & "' AND TFBN03='" & strCase(3) & "' AND TFBN04='" & strCase(4) & "'", "TFBN05")
      If strExc(0) <> "" Then
         m_PrevForm.cmdTFBaseNo.BackColor = &HC0FFC0
      Else
         m_PrevForm.cmdTFBaseNo.BackColor = &H8000000F
      End If
   End If
   Set rsRD = Nothing
   Set frm020509 = Nothing
End Sub

Private Sub ClearForm(Optional ByVal pReset As Boolean = False)
   
   If pReset = True Then
      For Each oObj In lblFM2
         oObj.Caption = ""
      Next
      cmdState = -1
   End If
   For Each oObj In txtDB
      oObj.Text = ""
      oObj.Tag = ""
   Next
   lblNation.Caption = ""
   lblTFbaseCase.Caption = ""
   m_RowSeq = ""
   m_Flag1 = ""
End Sub

Private Sub ReadData(ByVal pReset As Boolean)
      
   Call SetGrid(True)
   If pReset = True Then
      cnnConnection.Execute "delete from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' "
      strR1 = "SELECT tfbn05,tfbn06 ,na03 as tfbn06n,decode(t1.tm01||t2.tm01,null,'非本所案件',decode(t1.tm01,null,decode(t2.tm28,'1',null,'N')||rtrim(t2.tm01||'-'||t2.tm02||'-'||t2.tm03||'-'||t2.tm04),decode(t1.tm28,'1',null,'N')||rtrim(t1.tm01||'-'||t1.tm02||'-'||t1.tm03||'-'||t1.tm04)) ) as tcase " & _
              ",'Y' as flag1 FROM tfbaseno,nation,trademark t1, trademark t2" & _
              " WHERE tfbn01='" & strCase(1) & "' and tfbn02='" & strCase(2) & "' and tfbn03='" & strCase(3) & "' and tfbn04='" & strCase(4) & "' " & _
              " and tfbn06=na01(+) AND tfbn05=t1.tm15(+) AND tfbn06=t1.tm10(+) AND tfbn05=t2.tm12(+) AND tfbn06=t2.tm10(+) order by tfbn08,tfbn09 "
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strR1)
      If intR = 1 Then
         Set RsTemp = PUB_CreateRecordset(rsRD, , , , Me.Name, mESeqNo)
      End If
   End If
   If Val(mESeqNo) > 0 Then
      strR1 = "SELECT '' as V, r001 AS tfbn05,r002 as tfbn06, r003 AS tfbn06n,r004 AS tcase, r005 AS flag1,rowseq " & _
              " from rdatafactory  where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and nvl(r005,'N')<>'D' " & _
              " order by rowseq "
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strR1)
      If intR = 1 Then
          Set MGrid1.Recordset = rsRD
          Call SetGrid(False)
      End If
   End If
End Sub

Private Sub SetGrid(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   '                        0        1           2         3            4         5       6
   arrGridHeadText = Array("V", "TF基礎案號", "TFBN06", "申請國家", "本所案號", "FLAG1", "ROWSEQ")
   arrGridHeadWidth = Array(200, 1800, 0, 1000, 1600, 0, 0)
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
       MGrid1.Clear
       MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
       MGrid1.row = 0
       MGrid1.col = iRow
       MGrid1.Text = arrGridHeadText(iRow)
       MGrid1.CellAlignment = flexAlignCenterCenter
       MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
   Next

   MGrid1.Visible = True
End Sub

Private Sub MGrid1_Click()

   'intLastRow = MGrid1.MouseRow '方便中斷的偵測,但是會影響Grid單選控制
   GridClick MGrid1, intLastRow, 0, 0
   
   If intLastRow > 0 Then
      ReadGrid
   End If
End Sub

Private Sub ReadGrid()

   Call ClearForm(False)
   
   If intLastRow > 0 Then
       m_Flag1 = "" & MGrid1.TextMatrix(intLastRow, 5)
       m_RowSeq = "" & MGrid1.TextMatrix(intLastRow, 6)
       txtDB(0).Text = "" & MGrid1.TextMatrix(intLastRow, 1)
       txtDB(1).Text = "" & MGrid1.TextMatrix(intLastRow, 2)
       lblNation.Caption = "" & MGrid1.TextMatrix(intLastRow, 3)
       lblTFbaseCase.Caption = "" & MGrid1.TextMatrix(intLastRow, 4)
       If Left(lblTFbaseCase, 1) = "N" Then
          lblTFbaseCase.ForeColor = vbRed
       Else
          lblTFbaseCase.ForeColor = &H80000012
       End If
       txtDB(0).Tag = txtDB(0).Text
       txtDB(1).Tag = txtDB(1).Text
   End If
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   If Index = 1 Then
      ' 申請國家不可輸入 001 - 008
      Select Case txtDB(Index)
         Case "001", "002", "003", "004", "005", "006", "007", "008":
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請國家不可輸入001-008"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            txtDB_GotFocus Index
            Exit Sub
         Case Else
      End Select
      ' 取得國家代碼
      lblNation = GetNationName(txtDB(Index), 0)
      If IsEmptyText(lblNation) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "國別代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtDB_GotFocus Index
         Exit Sub
      End If
   End If
   
   If Trim(txtDB(0)) <> "" And Trim(txtDB(1)) <> "" And mStatus = "Y" Then    '跳離開就要檢查
      lblTFbaseCase.Caption = PUB_GetTFbaseInfo(strCase(1), strCase(2), strCase(3), strCase(4), Trim(txtDB(0)), Trim(txtDB(1)), "")
      If Left(lblTFbaseCase, 1) = "N" Then
          lblTFbaseCase.ForeColor = vbRed
      Else
          lblTFbaseCase.ForeColor = &H80000012
      End If
   End If
End Sub
