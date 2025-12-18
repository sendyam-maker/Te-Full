VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090801_12 
   BorderStyle     =   1  '單線固定
   Caption         =   "接洽單-待收文區"
   ClientHeight    =   5870
   ClientLeft      =   2800
   ClientTop       =   3720
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5870
   ScaleWidth      =   8950
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   60
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   8865
      Begin VB.TextBox txtMsg 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   0
         TabIndex        =   9
         Text            =   "資料送出中...請稍候..."
         Top             =   0
         Width           =   8865
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   8
         Top             =   450
         Width           =   8820
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   6705
      TabIndex        =   6
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "修改(&M)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   4
      Left            =   5730
      TabIndex        =   5
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消關連"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   3
      Left            =   2085
      TabIndex        =   4
      Top             =   90
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "設定關連"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   930
      TabIndex        =   3
      Top             =   90
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回輸入畫面"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   7680
      TabIndex        =   2
      Top             =   90
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "全部送出"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   4560
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   90
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   5280
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   8850
      _ExtentX        =   15593
      _ExtentY        =   9313
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|接洽單編號|組別|相同案號|客戶|申請國家|案號|案件名稱|案件性質|本所期限|法定期限|總費用"
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
   End
End
Attribute VB_Name = "frm090801_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2022/7/20
Option Explicit

Public frmParent As Form '前一畫面


Private Sub cmdok_Click(Index As Integer)
Dim bolUpd As Boolean, ii As Integer, jj As Integer, i As Integer
Dim bolHadDel As Boolean
Dim bolFind As Boolean
Dim showNote As String
Dim bolConn As Boolean
Dim dblMaxWidth As Double
Dim strTotRow As String
Dim strCRL65 As String, tmpArr As Variant 'Add By Sindy 2023/5/4
Dim strLOS15 As String '案源案號 Add By Sindy 2024/5/7
Dim strCRL07 As String 'Add By Sindy 2025/6/12
   
   bolConn = False
   Select Case Index
      Case 0 '全部送出
         '檢查...
         For ii = 1 To Grid1.Rows - 1
            If Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) <> "" And Grid1.RowHeight(ii) > 0 Then
               '檢查案源B開頭的,要有2筆接洽單才能送出
               'Modified by Morgan 2025/4/18 轉B類補收文除外
               strExc(0) = "select CRL55" & _
                           " from ConsultRecordList" & _
                           " where CRL01 ='" & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) & "'" & _
                           " and (substr(CRL74,1,1)='B' and length(CRL74)=2) and CRL55 is not null" & _
                           " and not exists(select * from lawofficesource where los15=CRL55 and los06 is not null)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strLOS15 = GetCRL55toLOS15("" & RsTemp.Fields("CRL55")) '案源案號 Add By Sindy 2024/5/7
                  '改判斷DB資料筆數,因送出時有可能一筆送成功,另一筆未送成功,重新操作送出
                  'Modify By Sindy 2024/5/7 and CRL55='" & RsTemp.Fields("CRL55") & "' => 改 and instr(CRL55,'" & strLOS15 & "')>0
                  strExc(0) = "select CRL07,CRL55" & _
                              " from ConsultRecordList" & _
                              " where (substr(CRL74,1,1)='B' and length(CRL74)=2) and instr(CRL55,'" & strLOS15 & "')>0"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.RecordCount <= 1 Then
'                        'Modify By Sindy 2023/2/17 FCT,FCP都是另外填寫接洽單(不在系統中),所以只會有法務案1張接洽單
'                        If Left(RsTemp.Fields("CRL07"), 1) <> "F" Then
'                        '2023/2/17 END
                           MsgBox "B類案源尚未輸入完畢，不能送出！"
                           Exit Sub
'                        End If
                     End If
                  End If
'                  '檢查案源B開頭的,Grid要有2筆接洽單才能送出
'                  bolFind = False
'                  For jj = 1 To Grid1.Rows - 1
'                     If InStr(Trim(Grid1.TextMatrix(jj, PUB_MGridGetId("相同案號", Grid1))), RsTemp.Fields("CRL55")) > 0 And _
'                        InStr(Trim(Grid1.TextMatrix(jj, PUB_MGridGetId("接洽單編號", Grid1))), Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)))) = 0 Then
'
'                        bolFind = True
'                        Exit For
'                     End If
'                  Next jj
'                  If bolFind = False Then
'                     MsgBox "B類案源尚未輸入完畢，不能送出！"
'                     Exit Sub
'                  End If
               End If
            End If
         Next ii
         
On Error GoTo CheckingErr
         
         'Add By Sindy 2023/5/4 MCTF人員有設定關連,則可先取得連續案號
         If InStr(Pub_GetSpecMan("MCTF", True, "收信人員"), strUserNum) > 0 Then
            '檢查是否有設關連
            strCRL65 = ""
            For ii = 1 To Grid1.Rows - 1
               If Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("CRL65", Grid1))) <> "" _
                  And Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("CRL06", Grid1))) = "Y" _
                  And Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("CRL08", Grid1))) = "" Then
                  If InStr(strCRL65, Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("CRL65", Grid1)))) = 0 Then
                     strCRL65 = strCRL65 & "," & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("CRL65", Grid1)))
                  End If
               End If
            Next ii
            If strCRL65 <> "" Then
               strCRL65 = Mid(strCRL65, 2)
               cmdOK(0).Visible = False '鎖住,以免按二次
               Frame1.Top = 0
               Text2.Width = 0
               dblMaxWidth = 8820 'Text2.Width
               txtMsg = "新案正在取案號中...請稍候...": DoEvents
               Frame1.Visible = True
               tmpArr = Split(strCRL65, ",")
               strTotRow = UBound(tmpArr) + 1
               For ii = 0 To UBound(tmpArr)
                  Text2.Width = dblMaxWidth / Val(strTotRow) * (ii + 1): DoEvents
                  Call Pub_GetContinuousCaseNo(CStr(tmpArr(ii)), False)
               Next ii
               Me.Enabled = False
               Call QueryData(False)
               Me.Enabled = True
               Frame1.Visible = False 'Add By Sindy 2022/12/22
               DoEvents
            End If
         End If
         '2023/5/4 END
         
         cmdOK(0).Visible = False '鎖住,以免按二次
         'Add By Sindy 2022/12/22
         Frame1.Top = 0
         Text2.Width = 0
         dblMaxWidth = 8820 'Text2.Width
         strTotRow = Grid1.Rows - 1
         txtMsg = "資料送出中...請稍候...": DoEvents
         Frame1.Visible = True
         '2022/12/22 END
         For ii = 1 To Grid1.Rows - 1
            If Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) <> "" And Grid1.RowHeight(ii) > 0 Then
               '送出
               Text2.Width = dblMaxWidth / Val(strTotRow) * ii: DoEvents
               cnnConnection.BeginTrans: bolConn = True
               If PUB_AddConsultRecvFlow(Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)))) = True Then
                  cnnConnection.CommitTrans: bolConn = False
                  PUB_SendMailCache '發信(因電子收文)
                  Grid1.RowHeight(ii) = 0
                  'Added by Lydia 2023/11/28 原本在Service1，改在這裡寄Email(智財顧問之專業時數調整：智權人員收文智財顧問時，系統自動傳送信函至杜燕文之信箱。)
                  strExc(0) = "SELECT CRL07,CRL08,CRL09,CRL10,CRC03,CRC08" & _
                              " From ConsultRecordList, ConsultRecCMP" & _
                              " WHERE CRL01='" & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) & "' AND CRL01=CRC01(+) AND CRL07='ACS' AND CRC03='112' "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(1) = Pub_GetSpecMan("全所智權部主管")
                     If strExc(1) <> "" Then
                        '調整此郵件在收件人請假時不必發給職代。
                        PUB_SendMail strUserNum, strExc(1), "" & RsTemp.Fields("CRC08"), "已收文智財顧問案" & RsTemp.Fields("CRL07") & "-" & RsTemp.Fields("CRL08") & IIf("" & RsTemp.Fields("CRL09") & RsTemp.Fields("CRL10") <> "000", "-" & RsTemp.Fields("CRL09") & "-" & RsTemp.Fields("CRL10"), "") & "，請確認預設之執行事務", , , , , , , , , , , True
                     End If
                  End If
                  'end 2023/11/28
               Else
                  Frame1.Visible = False 'Add By Sindy 2022/12/22
                  GoTo CheckingErr
               End If
            End If
         Next ii
         Call FlowBatchSendMail(strUserNum) 'Add By Sindy 2022/10/4 整批發通知信: 電子收文(或部分簽核)通知信
         Frame1.Visible = False 'Add By Sindy 2022/12/22
         GoToBack
         Exit Sub
         
      Case 1 '回前畫面
         frmParent.m_blnCallPrint = False
         frmParent.Text5 = ""
         Unload Me
         frmParent.Show
         Exit Sub
         
      Case 2 '設定關連
         'Modify By Sindy 2025/6/12 排除L案號不可設關連,因目前還是人工收文
         Grid1.Enabled = False
         For ii = 1 To Grid1.Rows - 1
            If Me.Grid1.TextMatrix(ii, 0) = "V" Then
               strExc(10) = Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("案號", Grid1)))
               strCRL07 = SystemNumber(strExc(10), 1) '系統別
               If InStr(strCRL07, "L") > 0 Then
                  Grid1.row = ii
                  Grid1.col = 0
                  '資料列清除反白
                  Me.Grid1.Text = ""
                  For i = 0 To Grid1.Cols - 1
                     If i <> 2 Then
                        Grid1.col = i
                        Grid1.CellBackColor = QBColor(15)
                     End If
                  Next i
                  Call SetColor(CDbl(ii))
               End If
            End If
         Next ii
         Grid1.Enabled = True
         '2025/6/12 END
         If PUB_SetCRLGroup(Grid1, PUB_MGridGetId("接洽單編號", Grid1), PUB_MGridGetId("CRL65", Grid1), bolUpd, _
            PUB_MGridGetId("CRL06", Grid1), PUB_MGridGetId("相同案號", Grid1), PUB_MGridGetId("CRL74", Grid1), _
            PUB_MGridGetId("案號", Grid1), PUB_MGridGetId("客戶", Grid1), PUB_MGridGetId("案件名稱", Grid1)) = False Then Exit Sub
         If bolUpd = True Then Call QueryData(False)
         
      Case 3 '取消關連
         If PUB_CancelCRLGroup(Grid1, PUB_MGridGetId("接洽單編號", Grid1), PUB_MGridGetId("CRL65", Grid1), bolUpd, _
            PUB_MGridGetId("CRL74", Grid1)) = False Then Exit Sub
         If bolUpd = True Then Call QueryData(False)
         
      Case 4 '修改
         For ii = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(ii, 0) = "V" And Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)) <> "" Then
               'Add By Sindy 2023/5/18
               strExc(0) = "select *" & _
                           " from flow003" & _
                           " where f0301 ='" & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) & "'" & _
                           " and f0309 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  MsgBox "此接洽單（" & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) & "）已送出，將重新查詢資料！"
                  Call QueryData(False)
                  Exit Sub
               End If
               '2023/5/18 END
               
               '查詢接洽記錄單
               frmParent.m_blnCallPrint = False
               frmParent.Text5 = Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)))
               frmParent.cmdOK(0).Caption = "存檔"
               frmParent.cmdOK(1).Tag = ""
               Grid1.TextMatrix(ii, 0) = ""
               For jj = 0 To Grid1.Cols - 1
                  If jj <> 2 Then
                     Grid1.col = jj
                     Grid1.row = ii
                     Grid1.CellBackColor = QBColor(15)
                  End If
               Next jj
               'Call frmParent.cmdOK_Click(4) 'Modify By Sindy 2024/6/24 mark
               'Me.Hide
               Unload Me
               frmParent.Show
               'Modify By Sindy 2024/6/24
               frmParent.Enabled = False
               Call frmParent.cmdok_Click(4)
               frmParent.Enabled = True
               '2024/6/24 END
               Exit For
            End If
         Next ii
         
      Case 5 '刪除
         If MsgBox("確定要刪除勾選的資料？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
            Exit Sub
         Else
            showNote = ""
         End If
         For ii = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(ii, 0) = "V" And Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)) <> "" And Grid1.RowHeight(ii) > 0 Then
               'Add By Sindy 2023/5/18
               strExc(0) = "select *" & _
                           " from flow003" & _
                           " where f0301 ='" & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) & "'" & _
                           " and f0309 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  MsgBox "此接洽單（" & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) & "）已送出，將重新查詢資料！"
                  Call QueryData(False)
                  Exit Sub
               End If
               '2023/5/18 END
               
               '檢查是否為案源資料
               Dim strCRL55 As String, strCRL74 As String
               Dim strLOS17 As String, strLOS18 As String, strChkNo As String
               strExc(0) = "select CRL55,CRL74" & _
                           " from ConsultRecordList" & _
                           " where CRL01 ='" & Trim(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) & "'" & _
                           " and length(CRL74)=2 and CRL55 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCRL55 = "" & RsTemp.Fields("CRL55")
                  strLOS15 = GetCRL55toLOS15(strCRL55) '案源案號 Add By Sindy 2024/5/7
                  strCRL74 = "" & RsTemp.Fields("CRL74")
                  
                  strExc(0) = "select LOS15,LOS17,LOS18 from LawOfficeSource where LOS15='" & strLOS15 & "' and (LOS17='" & Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)) & "' or LOS18='" & Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)) & "')"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strLOS17 = "" & RsTemp.Fields("LOS17")
                     strLOS18 = "" & RsTemp.Fields("LOS18")
                  End If
                  strChkNo = ""
                  If InStr(showNote, strLOS17) = 0 Or InStr(showNote, strLOS18) = 0 Then
                     'Add By Sindy 2024/9/24
                     strExc(10) = ""
                     If strLOS17 <> "" Then
                        If strExc(10) = "" Then
                           strExc(10) = strLOS17
                        Else
                           strExc(10) = strExc(10) & "," & strLOS17
                        End If
                     End If
                     If strLOS18 <> "" Then
                        If strExc(10) = "" Then
                           strExc(10) = strLOS18
                        Else
                           strExc(10) = strExc(10) & "," & strLOS18
                        End If
                     End If
                     If MsgBox("此為法律案源接洽單，確定（" & strExc(10) & "）要刪除？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                     '2024/9/24 END
                        Exit Sub
                     Else
                        showNote = showNote & IIf(strLOS17 <> "", "," & strLOS17, "") & IIf(strLOS18 <> "", "," & strLOS18, "")
                        '檢查是否有另一張要勾選
                        If strLOS17 <> Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)) And strLOS17 <> "" Then strChkNo = strLOS17
                        If strLOS18 <> Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)) And strLOS18 <> "" Then strChkNo = strLOS18
                        If strChkNo <> "" Then
                           For jj = 1 To Grid1.Rows - 1
                              If Trim(Grid1.TextMatrix(jj, 0)) = "" And Grid1.TextMatrix(jj, PUB_MGridGetId("接洽單編號", Grid1)) = strChkNo Then
                                 Grid1.col = 0
                                 Grid1.row = jj
                                 Grid1.Text = "V"
                                 For i = 0 To Grid1.Cols - 1
                                    If i <> 2 Then
                                       Grid1.col = i
                                       Grid1.CellBackColor = &HFFC0C0
                                    End If
                                 Next i
                                 Exit For
                              End If
                           Next jj
                        End If
                     End If
                  End If
               End If
            End If
         Next ii
         For ii = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(ii, 0) = "V" And Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1)) <> "" And Grid1.RowHeight(ii) > 0 Then
               cnnConnection.BeginTrans: bolConn = True
               If PUB_DelCRLAllData(Grid1.TextMatrix(ii, PUB_MGridGetId("接洽單編號", Grid1))) = True Then
                  cnnConnection.CommitTrans: bolConn = False
                  bolHadDel = True
                  Grid1.RowHeight(ii) = 0
               Else
                  If bolHadDel = True Then QueryData
                  Exit Sub
               End If
            End If
         Next ii
         If bolHadDel = True Then QueryData
         If (Grid1.Rows - 1) = 1 Then
            If Grid1.TextMatrix(1, PUB_MGridGetId("接洽單編號", Grid1)) = "" Then
               Call GoToBack
               Exit Sub
            End If
         End If
         
      Case Else
   End Select
   
   Exit Sub
   
CheckingErr:
   Frame1.Visible = False 'Add By Sindy 2023/5/4
   If bolConn = True Then
      cnnConnection.RollbackTrans: bolConn = False
   End If
   cmdOK(Index).Visible = True
   If Err.Description <> "" Then
      MsgBox (Err.Description), , "全部送出"
   Else
      MsgBox "全部送出失敗！", vbExclamation, "全部送出"
   End If
End Sub

Private Sub GoToBack()
   frmParent.m_blnCallPrint = False
   frmParent.cmdSend.Tag = "無資料"
   Unload Me
   frmParent.Show
End Sub

Private Sub Form_Load()
'   Dim oLabel As LABEL
'
'   m_AttachPath = App.path & "\" & strUserNum
'   Me.Height = 4230
   
   MoveFormToCenter Me, True
'   lstUsers(0).Clear
'   txtUserNo(0) = ""
'   lblName(0) = ""
'
'   For Each oLabel In Label2
'      oLabel.BackColor = &H8000000F
'   Next
'   SetcboLawMan
'   SetCaseType
   QueryData
'   Me.Tag = Me.Caption
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DestroyToolTip '清除物件
   PUB_SendMailCache '發信
   Call FlowBatchSendMail(strUserNum) 'Add By Sindy 2022/10/4 整批發通知信: 電子收文(或部分簽核)通知信
   Set frm090801_12 = Nothing
End Sub

Public Function QueryData(Optional ByVal bolFirst As Boolean = True) As Boolean
Dim strSql As String, intQ As Integer
Dim RsQ As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim stCP10 As String, strVal As String, strValCon As String
Dim strShowType As String
Dim strCRL49 As String
   
   QueryData = True
   Me.Grid1.Clear
   SetDataListWidth
   
   strVal = "select CRL01"
   ',ConsultRecCMP / " and crl01=CRC01(+) and CRC02 is not null"
   strValCon = " From ConsultRecordList,flow003" & _
            " where crl02>=" & 接洽單電子收文啟用日 & _
            " and CRL01=f0301(+) and f0301 is not null and f0309 is null" & _
            " and crl78='" & strUserNum & "'"
   
   '自動更新關連表單編號
   If bolFirst = True Then
      strSql = "select CRL55,count(*)" & strValCon & _
               " and CRL55 is not null and CRL74 is null" & _
               " group by CRL55"
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, strSql)
      If intQ = 1 Then
         RsQ.MoveFirst
         Do While Not RsQ.EOF
            If Val(RsQ.Fields(1)) > 1 Then '相同案號有輸入一筆以上接洽單
               strExc(1) = ""
               '檢查是否已有關連表單編號
               strExc(0) = "select CRL65" & _
                           " from ConsultRecordList" & _
                           " where CRL01 in(" & strVal & strValCon & ")" & _
                           " and CRL55='" & RsQ.Fields(0) & "' and CRL65 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = RsTemp.Fields("CRL65")
               Else
                  strExc(0) = "select CRL01" & _
                              " from ConsultRecordList" & _
                              " where CRL01 in(" & strVal & strValCon & ")" & _
                              " and CRL55='" & RsQ.Fields(0) & "'" & _
                              " order by CRL01 asc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(1) = RsTemp.Fields("CRL01")
                  End If
               End If
               If strExc(1) <> "" Then
                  strSql = "UPDATE CONSULTRECORDLIST SET CRL65='" & strExc(1) & "' WHERE CRL01 in(" & strVal & strValCon & ")" & _
                           " and (CRL55='" & RsQ.Fields(0) & "' or instr(CRL07||'-'||CRL08||'-'||CRL09||'-'||CRL10,'" & RsQ.Fields(0) & "')>0)"
                  cnnConnection.Execute strSql, intI
               End If
            End If
            RsQ.MoveNext
         Loop
      End If
   End If
   
   '設定組別
   cnnConnection.Execute "delete from r020115 where id='" & strUserNum & "'"
   strSql = "insert into r020115(r001002,r001005,ID)" & _
            " select CRL65,rownum,'" & strUserNum & "' from (select CRL65" & _
            strValCon & _
            " and CRL65 is not null" & _
            " group by CRL65" & _
            " order by CRL65 asc)"
   cnnConnection.Execute strSql, intI
   
   'Add By Sindy 2023/1/13
   '檢查:新案舊客戶同系統別同申請國家,發生同客戶多筆接洽單,有的收據公司"有值"有的"沒值" ex:CFT-023392
   strExc(0) = "select CRA05,CRL07,CRL15 From ConsultRecordList, consultrecapp" & _
               " where CRL01 in(" & strVal & strValCon & ")" & _
               " and CRL01=CRA01 and CRL06='Y' and CRA03 is null" & _
               " group by CRA05,CRL07,CRL15 having count(*)>1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strExc(0) = "select CRL49 From ConsultRecordList, consultrecapp" & _
                     " where CRL01 in(" & strVal & strValCon & ")" & _
                     " and CRL01=CRA01 and CRA05='" & RsTemp.Fields("CRA05") & "'" & _
                     " and CRL07='" & RsTemp.Fields("CRL07") & "'" & _
                     " and CRL15='" & RsTemp.Fields("CRL15") & "'" & _
                     " group by CRL49" ' and CRL49 is not null
         intI = 1
         Set RsQ = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsQ.MoveFirst
            Do While Not RsQ.EOF
               strCRL49 = "" & RsQ.Fields("CRL49")
   '            '更新收據公司
   '            strSql = "update ConsultRecordList set CRL49='" & strCRL49 & "' where CRL01 in(" & _
   '                     "select CRL01 From ConsultRecordList, consultrecapp" & _
   '                     " where CRL01 in(" & strVal & strValCon & ")" & _
   '                     " and CRL01=CRA01 and CRA05='" & RsTemp.Fields("CRA05") & "'" & _
   '                     " and CRL07='" & RsTemp.Fields("CRL07") & "' and CRL49 is null)"
   '            cnnConnection.Execute strSql, intI
               strExc(0) = "select CRL01,CRL49 From ConsultRecordList, consultrecapp" & _
                           " where CRL01 in(" & strVal & strValCon & ")" & _
                           " and CRL01=CRA01 and CRA05='" & RsTemp.Fields("CRA05") & "'" & _
                           " and CRL07='" & RsTemp.Fields("CRL07") & "'" & _
                           " and CRL15='" & RsTemp.Fields("CRL15") & "'"
               If strCRL49 = "" Then
                  strExc(0) = strExc(0) & " and CRL49 is not null"
               Else
                  strExc(0) = strExc(0) & " and (CRL49 is null or CRL49<>'" & strCRL49 & "')"
               End If
               intI = 1
               Set Rs = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "有新案舊客戶(" & RsTemp.Fields("CRA05") & ")同系統別(" & RsTemp.Fields("CRL07") & ")同申請國家(" & RsTemp.Fields("CRL15") & ")的接洽單，" & vbCrLf & vbCrLf & "收據公司不一致，請逐一檢查改正後再存檔送出。", vbExclamation
                  cmdOK(0).Enabled = False
                  GoTo ReadData
                  'RsTemp.MoveLast
               End If
               RsQ.MoveNext
            Loop
         End If
         RsTemp.MoveNext
      Loop
   End If
   
ReadData:
   '讀取顯示資料
   'Modify By Sindy 2023/5/4 + ,CRL08
   strSql = "select '' V,CRL01 接洽單編號,r001005 組別,decode(CRL74,null,CRL55,'') 相同案號" & _
            ",GetCRAName(crl01) 客戶,na03 申請國家" & _
            ",CRL07||'-'||decode(CRL08,null,'',CRL08)||decode(CRL09,null,'','-'||CRL09)||decode(CRL10,null,'','-'||CRL10) 案號" & _
            ",CRL17 案件名稱,CRL73 商品類別,GetCRCaseNmFee(crl01,'1') 案件性質,sqldatet(CRL12) 本所期限,sqldatet(CRL13) 法定期限" & _
            ",sum(CRC04) 總費用,CRL65,CRL06,CRL74,CRL79,CRL80,CRL90,CRL08" & _
            " from ConsultRecordList,ConsultRecCMP,nation,r020115" & _
            " where CRL01 in(" & strVal & strValCon & ") and CRL01=CRC01(+) and crl15=na01(+) and id(+)='" & strUserNum & "' and r001002(+)=CRL65" & _
            " group by CRL65,CRL01,CRL15,CRL07,CRL08,CRL09,CRL10,CRL17,CRL12,CRL13,CRL55,CRL56,na03,r001005,CRL06,CRL74,CRL79,CRL80,CRL73,CRL90,CRL08" & _
            " order by CRL65 asc,CRL79 asc,CRL80 asc,CRL01 asc" ',CRL08,CRL09,CRL10,CRL17,CRL12,CRL13,CRL55,CRL56,na03,r001005
            '排序:關連表單編號,系統別,申請國家,專利種類,接洽單編號
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
      Set Grid1.Recordset = RsQ
   End If
   SetColor
   
   Set RsQ = Nothing
   Set Rs = Nothing
End Function

Private Sub SetDataListWidth()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   '                        0    1             2       3           4       5           6       7           8           9           10          11          12        13       14       15       16       17       18       19
   arrGridHeadText = Array("V", "接洽單編號", "組別", "相同案號", "客戶", "申請國家", "案號", "案件名稱", "商品類別", "案件性質", "本所期限", "法定期限", "總費用", "CRL65", "CRL06", "CRL74", "CRL79", "CRL80", "CRL90", "CRL08")
   arrGridHeadWidth = Array(200, 0, 0, 800, 1200, 800, 900, 1000, 450, 1300, 800, 800, 800, 0, 0, 0, 0, 0, 0, 0)
   Grid1.Visible = False
   Grid1.Cols = UBound(arrGridHeadText) + 1
   Grid1.Rows = 2
   For iRow = 0 To Grid1.Cols - 1
      Grid1.row = 0
      Grid1.col = iRow
      Grid1.Text = arrGridHeadText(iRow)
      Grid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      Grid1.CellAlignment = flexAlignCenterCenter
   Next
   Grid1.Visible = True
End Sub

Private Sub Grid1_DblClick()
   Call cmdok_Click(4)
End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Static iRow As Integer, iCol As Integer
   
   If Grid1.MouseRow <> 0 And (Grid1.MouseCol = PUB_MGridGetId("相同案號", Grid1) Or Grid1.MouseCol = PUB_MGridGetId("客戶", Grid1) Or _
      Grid1.MouseCol = PUB_MGridGetId("案件名稱", Grid1) Or Grid1.MouseCol = PUB_MGridGetId("案件性質", Grid1)) Then
      If iRow <> Grid1.MouseRow Or iCol <> Grid1.MouseCol Then
         If Grid1.TextMatrix(Grid1.MouseRow, Grid1.MouseCol) <> "" Then
            CreateToolTip GetHWndForToolTip(Grid1), Grid1.TextMatrix(Grid1.MouseRow, Grid1.MouseCol)
            iRow = Grid1.MouseRow
            iCol = Grid1.MouseCol
         End If
      End If
   End If
End Sub

Private Sub Grid1_SelChange()
Dim intRow As Integer, intCol As Integer
Dim ii As Integer, i As Integer

Grid1.Visible = False
Grid1.row = Grid1.MouseRow
Grid1.col = 0
intRow = Me.Grid1.row
intCol = Me.Grid1.col
If Grid1.row <> 0 Then
   'Me.Grid1.row = intRow
   'Me.Grid1.col = intCol
   '資料列清除反白
   If Me.Grid1.TextMatrix(intRow, 0) = "V" Then
      Me.Grid1.Text = ""
      For i = 0 To Grid1.Cols - 1
         If i <> 2 Then
            Grid1.col = i
            Grid1.CellBackColor = QBColor(15)
         End If
      Next i
   '目前資料列反白
   Else
      Grid1.Text = "V"
      For i = 0 To Grid1.Cols - 1
         If i <> 2 Then
            Grid1.col = i
            Grid1.CellBackColor = &HFFC0C0
         End If
      Next i
   End If
End If
Grid1.Visible = True
End Sub

Private Sub SetColor(Optional intSetRow As Double = 0)
   Dim ii As Integer, jj As Integer, intNum As Integer
   
   With Grid1
   If .Rows > 1 Then
      .Visible = False
      For ii = IIf(intSetRow = 0, 1, intSetRow) To IIf(intSetRow = 0, .Rows - 1, intSetRow)
         '標示顏色註記
         If Val(Trim(.TextMatrix(ii, PUB_MGridGetId("組別", Grid1)))) > 0 Then
            intNum = Val(.TextMatrix(ii, PUB_MGridGetId("組別", Grid1))) Mod 6
'            If intNum = 0 Then
'               intNum = 14
'            Else
'               intNum = intNum + 8
'            End If
            For jj = 3 To 5
               .col = jj
               .row = ii
               .CellBackColor = PGColor(intNum) '0~5 QBColor(intNum) '9~14
            Next jj
         End If
         If Trim(.TextMatrix(ii, PUB_MGridGetId("CRL90", Grid1))) = "Y" Then '急件
            For jj = 0 To 12
               If Not (jj >= 1 And jj <= 5) Then
                  .col = jj
                  .row = ii
                  .CellBackColor = &H8080FF '紅色
               End If
            Next jj
         End If
      Next ii
      If intSetRow = 0 Then .TopRow = 1
      .Visible = True
   End If
   End With
End Sub
