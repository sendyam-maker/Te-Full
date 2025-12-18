VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210149 
   BorderStyle     =   1  '單線固定
   Caption         =   "待處理區"
   ClientHeight    =   5748
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8952
   Begin VB.FileListBox File1 
      Height          =   252
      Left            =   1680
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "顯示E化函待確收"
      Height          =   225
      Left            =   6960
      TabIndex        =   10
      Top             =   510
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "全選"
      Height          =   300
      Left            =   60
      TabIndex        =   5
      Top             =   420
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "刪指示信"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5220
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3450
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   90
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新"
      Default         =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   6180
      TabIndex        =   2
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "明細資料"
      Height          =   360
      Index           =   2
      Left            =   7200
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   3
      Left            =   8145
      TabIndex        =   4
      Top             =   60
      Width           =   765
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm210149.frx":0000
      Height          =   4995
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   8835
      _ExtentX        =   15579
      _ExtentY        =   8805
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|表單編號|表單類別|智權人員|本所案號|總收文號|案件性質|本所期限|法定期限|目前表單狀態"
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
      _Band(0).Cols   =   10
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   1500
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2646;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "程序人員："
      Height          =   240
      Left            =   30
      TabIndex        =   9
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "表單類別："
      Height          =   240
      Left            =   2490
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm210149"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/27 Form2.0已修改
'Create by Sindy 2015/1/12
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public cmdState As Integer '紀錄作用按鍵
Public m_ProState As String '系統別
Dim m_AttachPath As String 'Added by Morgan 2015/11/24
Public idxCP10 As Integer, idxNP22 As Integer, idxCP09 As Integer, idxCaseNo As Integer

'Add by Amy 2018/10/08
Private Sub cmdChoose_Click()
    Dim strV As String
    
    If cmdChoose.Caption = "全選" Then strV = "V"
    GRD1.Visible = False
    If GRD1.Rows > 1 Then
        If GRD1.TextMatrix(1, 1) <> "" Then
            For j = 1 To GRD1.Rows - 1
                GRD1.col = 0
                GRD1.row = j
                GRD1.Text = strV
                For i = 0 To GRD1.Cols - 1
                    GRD1.col = i
                    If cmdChoose.Caption = "全選" Then
                        GRD1.CellBackColor = &HFFC0C0
                    Else
                         GRD1.CellBackColor = QBColor(15)
                    End If
                Next i
                If cmdChoose.Caption <> "全選" Then Call SetColColor(GRD1.MouseRow)
            Next j
        End If
    End If
    If cmdChoose.Caption = "全選" Then
        cmdChoose.Caption = "全部取消"
    Else
        cmdChoose.Caption = "全選"
    End If
    GRD1.Visible = True
End Sub
'end 2018/10/08

Private Sub cmdDelete_Click()
   Dim ii As Integer
   
   With GRD1
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) = "V" Then
         If .TextMatrix(ii, 10) = "4" Then
            If MsgBox("是否確定要刪除 " & .TextMatrix(ii, 4) & " 指示信？", vbYesNo + vbDefaultButton2 + vbQuestion, "刪除指示信 (編號:" & .TextMatrix(ii, 1) & ")") = vbYes Then
               If DeleteAppForm(.TextMatrix(ii, idxCP09), .TextMatrix(ii, 1)) Then
                  .RowHeight(ii) = 0
               End If
            End If
         Else
            MsgBox "表單編號 " & .TextMatrix(ii, 1) & " 非指示信不可刪除！", vbCritical
            Exit For
         End If
      End If
   Next
   End With
   Call QueryData
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Dim bolDone As Boolean, strSendDate  As String, strSendTime As String, bolInTrans As Boolean 'Added by Morgan 2015/11/24
   Dim stLetter As String 'Added by Morgan 2018/8/24
   Dim frmNext As Form, cp() As String 'Added by Morgan 2020/1/16
   Dim intStep As Integer, stCP09 As String, stNP22 As String, stNP07 As String 'Add by Amy 2020/12/10
   Dim strCmd As String, intExc As Integer 'Add by Amy 2022/03/30
   Dim strMsg As String 'Add by Amy 2024/10/08
   Dim strNP11 As String 'Add by Amy 2025/01/22
   Dim intFCState As String, strST15 As String, strSysKind As String, strNation As String 'Add by Amy 2025/04/10
   Dim strCCM18 As String 'Add by Amy 2025/06/19
   Dim o_A1K04 As String, o_A1K27 As String, o_A1K28 As String, o_A1K29 As String, o_TM56 As String, o_TM69 As String, o_CCM20 As String 'Add by Amy 2025/11/13
   
On Error GoTo ErrHnd 'Added by Morgan 2015/11/24
   
   Select Case cmdState
      Case 1 '查詢
         If QueryData = False Then ShowNoData
         
      Case 2 '明細資料
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            Me.Tag = "" 'Add By Sindy 2018/10/18
            If Trim(GRD1.Text) = "V" Then
               GRD1.col = 0
               GRD1.Text = ""
               For j = 0 To GRD1.Cols - 1
                  GRD1.col = j
                  GRD1.CellBackColor = QBColor(15)
               Next j
               GRD1.col = 4
               'Modify By Sindy 2018/10/18 Mark,會導至form.hide下列寄送時,會有Form不見的問題
'               'Add by Amy 2018/10/08
'                If fnSaveParentForm(Me) = False Then
'                  Me.Enabled = True
'                  Exit Sub
'                End If
                'end 2018/10/08
               If Not IsNull(GRD1.Text) Then
                  Screen.MousePointer = vbHourglass
                  
                  'Added by Morgan 2019/1/11
                  '6 客戶函
                  If GRD1.TextMatrix(GRD1.row, 10) = "6" Then
                     '判發退回(05)
                     If GRD1.TextMatrix(GRD1.row, 11) = "05" Then
                        EditLetter GRD1.TextMatrix(GRD1.row, idxCP09)
                     End If
                  'end 2019/1/11
                  
                  'Added by Morgan 2015/11/17
                  '4 指示信
                  ElseIf GRD1.TextMatrix(GRD1.row, 10) = "4" Then
                     '未寄送(11)
                     If GRD1.TextMatrix(GRD1.row, 11) = "11" Then
                        If PUB_SendOrderLetterP(GRD1.TextMatrix(GRD1.row, idxCP09)) Then  'Added by Morgan 2016/3/30
                        'Removed by Morgan 2018/8/30 移到 PUB_SendOrderLetterP
                        '      '電子表單
                        '      strSql = "update flow003 set f0309='03' where f0301='" & GRD1.TextMatrix(GRD1.row, 1) & "' and f0309='09'"
                        '      cnnConnection.Execute strSql, intI
                        'end 2018/8/30
                           GRD1.RowHeight(GRD1.row) = 0
                        End If
                     '未上傳(10)或判發退回(05)
                     Else
                        EditLetter GRD1.TextMatrix(GRD1.row, idxCP09), "2" 'Modified by Morgan 2019/1/19 改用函數
                     End If
                     
                  'Added by Morgan 2018/2/6
                  '5 帳單
                  ElseIf GRD1.TextMatrix(GRD1.row, 10) = "5" Then
                     strExc(0) = "select a1524 from acc150 where a1501='" & GRD1.TextMatrix(GRD1.row, 1) & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        'Modified by Morgan 2022/1/12
                        'MsgBox RsTemp(0), vbInformation, "帳單退回意見"
                        UniMsgBox RsTemp(0), vbInformation, "帳單退回意見"
                        'end 2022/1/12
                        Forms(0).mnu110401_Click 1
                        If strFormName = "Frmacc2150" Then
                           Set Frmacc2150.m_ParentForm = Me
                           Frmacc2150.Text2 = GRD1.TextMatrix(GRD1.row, 1)
                           Frmacc2150.Command3.Value = True
                           KeyEnter vbKeyF3
                        Else
                           MsgBox "無法開啟帳單維護畫面!!!", vbCritical
                        End If
                     End If
                  'end 2018/2/6
                  
                  'Added by Morgan 2018/10/25
                  '7 E化函
                  ElseIf GRD1.TextMatrix(GRD1.row, 10) = "7" Then
                     If GRD1.TextMatrix(GRD1.row, 11) = Flow_待寄送 Then
                        If PUB_SendECustLetter(GRD1.TextMatrix(GRD1.row, idxCP09)) Then
                           GRD1.RowHeight(GRD1.row) = 0
                        End If
                     'Added by Morgan 2021/3/30
                     ElseIf GRD1.TextMatrix(GRD1.row, 11) = Flow_待確收 Then
                        If PUB_RecpConfirm(GRD1.TextMatrix(GRD1.row, idxCP09), Me) Then
                           GRD1.RowHeight(GRD1.row) = 0
                        End If
                     End If
                  
                  'Added by Morgan 2020/1/16
                  '8 C類未發文
                  ElseIf GRD1.TextMatrix(GRD1.row, 10) = "8" Then
                     Set frmNext = mdiMain.GetForm("frm040104_1")
                     With frmNext
                     cp() = Split(GRD1.TextMatrix(GRD1.row, idxCaseNo), "-")
                     .mCP09 = GRD1.TextMatrix(GRD1.row, idxCP09)
                     Set .mPreForm = Me
                     .Show
                     .Option1(0).Value = True
                     .Text1 = cp(0)
                     .Text2 = cp(1)
                     .Text3 = cp(2)
                     .Text4 = cp(3)
                     .Command1.Value = True
                     End With
                     
                  'Add by Amy 2018/08/30 +T延展結案 (智權人員不需走流程)
                  ElseIf GRD1.TextMatrix(GRD1.row, 1) = "延展結案" Then
                    'Add by Amy 2020/12/10 +結案檢查(同frm210149_1) ex:T-168693 智權人員操作延展結案後,又收文延展(AA8049058),程序操作未彈訊息做了閉卷
                    cp() = Split(GRD1.TextMatrix(GRD1.row, idxCaseNo), "-")
                    stCP09 = GRD1.TextMatrix(GRD1.row, idxCP09)
                    stNP07 = GRD1.TextMatrix(GRD1.row, idxCP10)
                    stNP22 = GRD1.TextMatrix(GRD1.row, idxNP22)
                    'Add by Amy 2025/01/20 內商人員共用待處理區,避免同時處理同一筆資料,造成後續資料有問題
                    If Pub_ChkLock(0, "frm110101_2", "C", "解除期限", cp(0) & cp(1) & cp(2) & cp(3)) = False Then
                        Screen.MousePointer = vbDefault
                        Me.Enabled = True
                        Exit Sub
                    End If
                    'Modify by Amy 2025/01/23  避免同一個人,同時開2個商標系統(同時開待處理區),另一個畫面於解除期限已結案,回另一個待處理區時又按同一筆,訊息又沒注意看,導致結案又把ti06上 Y-與Sindy 討論後先+np11
                    If ChkNotCloseStep(intStep, cp(0), cp(1), cp(2), cp(3), stCP09, stNP22, stNP07, strNP11) = True Then
                        'Modify by Amy 2022/03/30 +是否選項,是-更新取消延展(ti06='Y'),ex:T-129927 做了結案又收文,導致專業部無法進行後續操作
                        If intStep = 1 Then
                            If strNP11 <> MsgText(601) Then
                              strMsg = cp(0) & "-" & cp(1) & IIf(cp(2) & cp(3) = "000", "", cp(2) & "-" & cp(3)) & "「" & IIf(stNP07 = "102", "延展", "註冊費") & "」期限" & _
                                             "已操作過「解除期限」畫面將更新"
                              MsgBox strMsg
                              Call cmdok_Click(1)
                              Screen.MousePointer = vbDefault
                              Me.Enabled = True
                              Exit Sub
                            End If
                    'end 2025/01/23
                            'MsgBox "已無「" & IIf(stNP07 = "102", "延展", "註冊費") & "」期限" & vbCrLf & "請向智權人員確認！"
                            'Modify by Amy 2024/10/08 DeBug用
                            strMsg = "已無「" & IIf(stNP07 = "102", "延展", "註冊費") & "」期限" & vbCrLf & _
                                             "已向智權人員確認" & vbCrLf & _
                                             "要取消「延展」結案！"
                            If MsgBox(strMsg, vbYesNo + vbDefaultButton2) = vbYes Then
                                strCmd = "Update T102Inform Set ti06='Y' Where ti02='" & stCP09 & "' And ti04='" & stNP22 & "' "
                                cnnConnection.Execute strCmd, intExc
                                If intExc > 0 Then
                                    Pub_SeekTbLog strCmd 'Add by Amy 2024/10/04 加入以便知道為何已結案,ti06又設 Y
                                    strExc(9) = GetPrjSalesNM(strUserNum) & "(" & strUserNum & ") 於[待處理區] 操作" & cp(0) & "-" & cp(1) & "-" & cp(2) & "-" & cp(3) & "(" & stCP09 & ")" & vbCrLf & _
                                                      "彈[" & strMsg & "]" & vbCrLf & _
                                                      "按[確定]-->要取消「延展」結案！"
                                    PUB_SendMail strUserNum, "A2004", "", "T102Inform.Ti06 已上Y 請確認是否User 按錯", strExc(9), , , , , , , , , , , , , True
                            'end 2024/10/08
                                    MsgBox cp(0) & "-" & cp(1) & "-" & cp(2) & "-" & cp(1) & vbCrLf & _
                                                    "已取消「延展」結案！"
                                    Call cmdok_Click(1)
                                End If
                            End If
                        End If
                        'end 2022/03/30
                        Screen.MousePointer = vbDefault
                        Me.Enabled = True
                        Exit Sub
                    End If
                    'end 2020/12/10
                    Set frm110101_2.mPrev01 = Me
                    Me.Tag = "延展結案" 'Add By Sindy 2018/10/18
                    frm110101_2.Pre_ProState = m_ProState 'Add by Amy 2023/02/13 內商人員共用待處理區,避免同時處理同一筆資料,造成後續資料有問題
                    frm110101_2.Show
                    Me.Hide
                  Else
                  'end 2015/11/17
                     Me.Hide
                     'Add by Amy 2025/04/10 +FC結案單
                     intFCState = 0 '非FC結案單
                     strST15 = PUB_GetStaffST15(GRD1.TextMatrix(i, PUB_MGridGetId("F0316", GRD1)), 1)
                     strSysKind = GRD1.TextMatrix(i, PUB_MGridGetId("SYSKIND", GRD1))
                     strNation = GetPrjNation1(GRD1.TextMatrix(i, PUB_MGridGetId("本所案號", GRD1)))
                     If strSrvDate(1) >= FCP結案單電子化啟用日 Then
                        'Modify by Amy 2025/06/26 發現舊資料會頁籤判斷會有問題FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案
                        '       ex:FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案 / 外商承辦使用國內結案單操作結案 ex:T-242111(結案單號11203939)
                        strCCM18 = Pub_GetField("CloseCaseMain", "CCM01='" & GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1)) & "'", "CCM18")
                        If strCCM18 = "F" Then
                           If strSysKind = "FCP" Or strSysKind = "FG" Or strSysKind = "P" Or strSysKind = "CFP" Then
                              intFCState = 2
                           Else
                              intFCState = 1
                           End If
                        End If
                        'end 2025/06/26
                     End If
                     frm210149_1.intFCState = intFCState
                     frm210149_1.m_NP07 = GRD1.TextMatrix(i, PUB_MGridGetId("CP10", GRD1))
                     'end 2025/04/10
                     
                     Call frm210149_1.SetParent(Me)
                     frm210149_1.Hide
                     'Add by Amy 2023/02/13 內商人員共用待處理區,避免同時處理同一筆資料,造成後續資料有問題
                     If m_ProState = "T" Then
                        frm210149_1.Pre_ProState = m_ProState
                     End If
                     frm210149_1.txtF0301 = GRD1.TextMatrix(i, 1) '表單編號
                     'frm210149_1.QueryData 'Mark by Amy 2025/08/18 往下搬
                     Me.Tag = "結案單" 'Add By Sindy 2018/10/18
                     frm210149_1.Show
                     frm210149_1.QueryData 'Add by Amy 2025/08/18 從上面搬下來,外商 Run Pub_CloseShowfrm210133_INV 按鈕才能顯示
                     frm210149_1.ChkData 'Add By Sindy 2015/5/13
                     'Add by Amy 2025/11/13 外商結案單
                     If strCCM18 = "F" And intFCState = 1 Then
                        cp() = Split(GRD1.TextMatrix(GRD1.row, idxCaseNo), "-")
                        Call Pub_GetCloseA1KData(3, Me.Name, cp(0), cp(1), cp(2), cp(3), o_A1K29, GRD1.TextMatrix(i, PUB_MGridGetId("CP10", GRD1)), o_A1K04, o_A1K27, o_A1K28, o_TM56, o_TM69, GRD1.TextMatrix(i, 1), o_CCM20, strMsg)
                        If strMsg <> "" Then
                           MsgBox "此結案單有設定請款單資訊如下：" & Replace(strMsg, ";", vbCrLf) & vbCrLf & _
                                       "請確認！"
                        End If
                     End If
                  End If 'Added by Morgan 2015/11/17
                  
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
         Call QueryData
      Case 3 '結束
         Unload Me
      Case Else
   End Select
   
'Added by Morgan 2015/11/24
ErrHnd:
   If Err.Number <> 0 Then
      If bolInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   Set frmNext = Nothing 'Added by Morgan 2020/1/16
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String
Dim strCon2 As String 'Added by Morgan 2015/11/12
Dim strCon3 As String 'Added by Morgan 2018/2/6
Dim strCon4 As String 'Added by Morgan 2018/6/22
Dim strCon7 As String 'Added by Morgan 2018/8/24
Dim strCon8 As String 'Added by Morgan 2018/10/25
Dim strCon9 As String 'Added by Morgan 2018/10/25
Dim strCon10 As String 'Added by Morgan 2019/1/9
Dim strCon11 As String 'Added by Morgan 2019/1/9
Dim strCon12 As String 'Added by Morgan 2020/1/16
'Add by Amy 2018/06/25
Dim strCon5 As String, strCon6 As String, strNP As String, strBase As String, strNP2 As String, strBase2 As String
Dim strTB As String, strField As String, strField2 As String
Dim strT102InForm As String 'Add by Amy 2018/08/30 T延展結案語法
Dim bolRightCloser As Boolean 'Added by Morgan 2019/5/3 結案程序人員是否為處理人員
Dim strCTB As String 'Add by Amy 2025/04/10 結案單主檔
   
   cmdChoose.Caption = "全選" 'Add by Amy 2018/10/08
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   SetGrd
   
   strCon = " And F0309 in('" & Flow_處理中 & "','" & Flow_判發退回 & "')"
   strCon2 = ""
   strCon3 = ""
   strCon4 = ""
   strCon7 = "" 'Added by Morgan 2018/8/24
   'Added by Morgan 2018/10/25 非FMP,指定E化客戶(LP26=E),已判發及確認的信函
   'Modified by Morgan 2021/4/1 掛號直寄未確認的要等1整個工作天(-2個工作天)後才可寄，平信直寄會自動確認可馬上寄
   'Modified by Morgan 2022/2/14 +cp27>19221111 (C類來函或報價定稿)
   strCon8 = " and substr(cp12,1,1)<>'F' and cp27>19221111 and lp26='E' and lp10='Y' and lp31 is null and lp05>0 and ((lp07>0 and lp11<>'2') or (lp11='Y' and cp127<" & CompWorkDay(2, strSrvDate(1), 1) & "))"
   strCon9 = strCon8
   'end 2018/10/25
   strCon10 = "" 'Added by Morgan 2019/1/9
   strCon11 = "" 'Added by Morgan 2019/1/9
   strCon12 = "" 'Added by Morgan 2020/1/16
   
   'Modify by Amy 2018/06/25 除法務/顧問外其他結案電子化
'   If m_ProState = "CFP" Then
'      strCon = strCon & " And pa01='CFP'"
'      If Trim(Combo1.Text) <> "" Then
'         strCon = strCon & " And F0308='" & Left(Combo1, 5) & "'"
'      End If
'      strCon2 = strCon2 & " and cp01='CFP'" 'Added by Morgan 2015/11/12
'      'Modified by Morgan 2018/6/22
'      'strCon3 = strCon3 & " and cp01 in ('CFP','CPS')" 'Added by Morgan 2018/2/6
'      strCon3 = strCon3 & " and cp01='CFP'"
'      strCon4 = strCon4 & " and cp01='CPS'"
'      'end 2018/6/22
'   ElseIf m_ProState = "P" Then
'      strCon = strCon & " And pa01='P'"
'      strCon2 = strCon2 & " and cp01='P'" 'Added by Morgan 2015/11/12
'      'Modified by Morgan 2018/6/22
'      'strCon3 = strCon3 & " and cp01 in ('P','PS')" 'Added by Morgan 2018/2/6
'      strCon3 = strCon3 & " and cp01='P'"
'      strCon4 = strCon4 & " and cp01='PS'"
'      'end 2018/6/22
'      If Left(Combo1, 1) = "1" Then '台灣案
'         strCon = strCon & " And PA09='000'"
'         strCon2 = strCon2 & " and PA09='000'" 'Added by Morgan 2015/11/12
'         strCon3 = strCon3 & " and PA09='000'" 'Added by Morgan 2018/2/6
'         strCon4 = strCon4 & " and SP09='000'" 'Added by Morgan 2018/6/22
'      ElseIf Left(Combo1, 1) = "2" Then '非台灣案
'         strCon = strCon & " And PA09<>'000'"
'         strCon2 = strCon2 & " and PA09<>'000'" 'Added by Morgan 2015/11/12
'         strCon3 = strCon3 & " and PA09<>'000'" 'Added by Morgan 2018/2/6
'         strCon4 = strCon4 & " and SP09<>'000'" 'Added by Morgan 2018/6/22
'      End If
'   End If
   strCon5 = " And F0309 in('" & Flow_處理中 & "','" & Flow_判發退回 & "')"
   strNP = " And F0309 in('" & Flow_處理中 & "','" & Flow_判發退回 & "')"
   strNP2 = " And F0309 in('" & Flow_處理中 & "','" & Flow_判發退回 & "')"
   strBase = " And F0309 in('" & Flow_處理中 & "','" & Flow_判發退回 & "')"
   strBase2 = " And F0309 in('" & Flow_處理中 & "','" & Flow_判發退回 & "')"
   'Add by Amy 2025/04/10 +FC結案
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strBase = strBase & " And length(CCM02)>=10 And F0301=CCM01(+)"
      strBase2 = strBase2 & " And length(CCM02)>=10 And F0301=CCM01(+)"
   Else
      strBase = strBase & " And length(F0303)>=10"
      strBase2 = strBase2 & " And length(F0303)>=10"
   End If
   strTB = ",Patent": strField = ",pa09": strField2 = ",PA01,PA02,PA03,PA04,'','',0,0,pa09"
   If InStr(m_ProState, "T") > 0 Then strTB = ",TradeMark": strField = ",tm10 as pa09": strField2 = ",TM01,TM02,TM03,TM04,'','',0,0,tm10 as pa09"
   'Modify by Amy 2025/04/10 +FCP結案
   If m_ProState = "CFP" Or m_ProState = "P" Or m_ProState = "FCP" Then
      'Add by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
      If strSrvDate(1) >= FCP結案單電子化啟用日 Then
         strCTB = UCase(",CloseCaseMain")
         strCon = strCon & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
                                          " And ccm03 is null And length(ccm02)=9 And ccm02=cp09 And F0301=CCM01(+)"
         strCon5 = strCon5 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
                                          " And ccm03 is null And length(ccm02)=9 And ccm02=cp09 And F0301=CCM01(+)"
          strNP = strNP & " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+)" & _
                                          " And CCM03 is not null and CCM02=NP01(+) and CCM03=NP22(+) And F0301=CCM01(+)"
         strNP2 = strNP2 & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+)" & _
                                          " And CCM03 is not null and CCM02=NP01(+) and CCM03=NP22(+) And F0301=CCM01(+)"
         strBase = strBase & " and substr(CCM02,1,length(CCM02)-9)=pa01(+) and substr(CCM02,length(CCM02)-8,6)=pa02(+)" & _
                                             " and substr(CCM02,length(CCM02)-2,1)=pa03(+) and substr(CCM02,length(CCM02)-1,2)=pa04(+)"
         strBase2 = strBase2 & " and substr(CCM02,1,length(CCM02)-9)=sp01(+) and substr(CCM02,length(CCM02)-8,6)=sp02(+)" & _
                                                   " and substr(CCM02,length(CCM02)-2,1)=sp03(+) and substr(CCM02,length(CCM02)-1,2)=sp04(+)"
      Else
         strCon = strCon & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
                                           " And F0304 is null and length(F0303)=9 and F0303=cp09 "
         strCon5 = strCon5 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
                                          " And F0304 is null and length(F0303)=9 and F0303=cp09 "
         strNP = strNP & " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) " & _
                                          "And F0304 is not null and F0303=NP01(+) and F0304=NP22(+) "
         strNP2 = strNP2 & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) " & _
                                          "And F0304 is not null and F0303=NP01(+) and F0304=NP22(+) "
         strBase = strBase & " and substr(F0303,1,length(F0303)-9)=pa01(+) and substr(F0303,length(F0303)-8,6)=pa02(+) and substr(F0303,length(F0303)-2,1)=pa03(+) and substr(F0303,length(F0303)-1,2)=pa04(+)"
         strBase2 = strBase2 & " and substr(F0303,1,length(F0303)-9)=sp01(+) and substr(F0303,length(F0303)-8,6)=sp02(+) and substr(F0303,length(F0303)-2,1)=sp03(+) and substr(F0303,length(F0303)-1,2)=sp04(+)"
      End If
      strCon2 = strCon2 & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
      strCon6 = strCon6 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)"
      strCon3 = strCon3 & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
      strCon4 = strCon4 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)"
      strCon8 = strCon8 & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa01 is not null" 'Added by Morgan 2018/10/25
      strCon9 = strCon9 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp01 is not null" 'Added by Morgan 2018/10/25
      strCon10 = strCon10 & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" 'Added by Morgan 2019/1/9
      strCon11 = strCon11 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" 'Added by Morgan 2019/1/9
      strCon12 = strCon12 & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" 'Added by Morgan 2020/1/16
      
      If m_ProState = "CFP" Then
        bolRightCloser = True 'Added by Morgan 2019/5/3
        strCon = strCon & " And pa01='CFP'"
        strCon5 = strCon5 & " And sp01='CPS'"
        strNP = strNP & " And pa01='CFP'"
        strNP2 = strNP2 & " And sp01='CPS'"
        strBase = strBase & " And pa01='CFP'"
        strBase2 = strBase2 & " And sp01='CPS'"
        strCon2 = strCon2 & " and cp01='CFP'"
        strCon6 = strCon6 & " and cp01='CPS'"
        strCon3 = strCon3 & " and cp01='CFP'"
        strCon4 = strCon4 & " and cp01='CPS'"
        strCon8 = strCon8 & " and cp01='CFP'" 'Added by Morgan 2018/10/25
        strCon9 = strCon9 & " and cp01='CPS'" 'Added by Morgan 2018/10/25
        strCon10 = strCon10 & " and cp01='CFP'" 'Added by Morgan 2019/1/9
        strCon11 = strCon11 & " and cp01='CPS'" 'Added by Morgan 2019/1/9
      'Add by Amy 2025/04/10 +FCP
      ElseIf m_ProState = "FCP" Then
         bolRightCloser = True
         strCTB = strCTB & UCase(",CaseProgress F0")
         'P 寰華案由外專程序處理
         strCon = strCon & " And pa01 in ('FCP','P')"
         strCon5 = strCon5 & " And sp01='FG'"
         strNP = strNP & " And pa01 in ('FCP','P') And NP01=CP09(+)"
         strNP2 = strNP2 & " And sp01='FG' And NP01=CP09(+)"
         strBase = strBase & " And pa01 in ('FCP','P') " & _
                                             " And pa01=cp01(+) And pa02=cp02(+) And pa03=cp03(+) And pa04=cp04(+)"
         strBase2 = strBase2 & " And sp01='FG'" & _
                                                   " And sp01=cp01(+) And sp02=cp02(+) And sp03=cp03(+) And sp04=cp04(+)"
         strCon2 = strCon2 & " and cp01 in ('FCP','P')"
         strCon6 = strCon6 & " and cp01='FG'"
         strCon3 = strCon3 & " and cp01 in ('FCP','P')"
         strCon4 = strCon4 & " and cp01='FG'"
         strCon8 = strCon8 & " and cp01 in ('FCP','P')"
         strCon9 = strCon9 & " and cp01='FG'"
         strCon10 = strCon10 & " and cp01 in ('FCP','P')"
         strCon11 = strCon11 & " and cp01='FG'"
      Else
         bolRightCloser = True 'Added by Morgan 2025/2/11
         
        'Modify by Amy 2025/04/10 +FC結案單
        If strSrvDate(1) >= FCP結案單電子化啟用日 Then
            strCTB = strCTB & UCase(",CaseProgress F0")
            
            strCon = strCon & " And pa01='P'"
            strCon5 = strCon5 & " And sp01='PS'"
            strNP = strNP & " And pa01='P' And NP01=CP09(+)"
            strNP2 = strNP2 & " And sp01='PS' And NP01=CP09(+)"
            strBase = strBase & " And pa01='P' And pa01=cp01(+) And pa02=cp02(+) And pa03=cp03(+) And pa04=cp04(+)"
            strBase2 = strBase2 & " And sp01='PS' And sp01=cp01(+) And sp02=cp02(+) And sp03=cp03(+) And sp04=cp04(+)"
        Else
            strCon = strCon & " And pa01='P'"
            strCon5 = strCon5 & " And sp01='PS'"
            strNP = strNP & " And pa01='P'"
            strNP2 = strNP2 & " And sp01='PS'"
            strBase = strBase & " And pa01='P'"
            strBase2 = strBase2 & " And sp01='PS'"
        End If
        strCon2 = strCon2 & " and cp01='P'"
        strCon6 = strCon6 & " and cp01='PS'"
        strCon3 = strCon3 & " and cp01='P'"
        strCon4 = strCon4 & " and cp01='PS'"
        strCon8 = strCon8 & " and cp01='P'" 'Added by Morgan 2018/10/25
        strCon9 = strCon9 & " and cp01='PS'" 'Added by Morgan 2018/10/25
        strCon10 = strCon10 & " and cp01='P'" 'Added by Morgan 2019/1/9
        strCon11 = strCon11 & " and cp01='PS'" 'Added by Morgan 2019/1/9
        strCon12 = strCon12 & " and cp01='P'" 'Added by Morgan 2020/1/16
      End If
      
      'Modified by Morgan 2025/1/10
      'If m_ProState = "CFP" Then
      If m_ProState = "CFP" Or strSrvDate(1) >= P業務區劃分啟用日 Then
        
        If Trim(Combo1.Text) <> "" Then
            strCon = strCon & " And F0308='" & Left(Combo1, 5) & "'"
            strCon5 = strCon5 & " And F0308='" & Left(Combo1, 5) & "'"
            strNP = strNP & " And F0308='" & Left(Combo1, 5) & "'"
            strNP2 = strNP2 & " And F0308='" & Left(Combo1, 5) & "'"
            strBase = strBase & " And F0308='" & Left(Combo1, 5) & "'"
            strBase2 = strBase2 & " And F0308='" & Left(Combo1, 5) & "'"
            'Added by Morgan 2018/8/24
            strCon2 = strCon2 & " And AF15='" & Left(Combo1, 5) & "'"
            strCon6 = strCon6 & " And AF15='" & Left(Combo1, 5) & "'"
            strCon7 = strCon7 & " And '" & Left(Combo1, 5) & "' in (a1516,a1519)"
            'end 2018/8/24
            
            'Added by Morgan 2021/11/3 E化函
            'Modified by Morgan 2022/6/22 考慮請假時由職代發文，EMail應該還是管制人要處理(LP38新增時寫入)
            'strCon8 = strCon8 & " And cp83='" & Left(Combo1, 5) & "'"
            'strCon9 = strCon9 & " And cp83='" & Left(Combo1, 5) & "'"
            strCon8 = strCon8 & " And nvl(lp38,cp83)='" & Left(Combo1, 5) & "'"
            strCon9 = strCon9 & " And nvl(lp38,cp83)='" & Left(Combo1, 5) & "'"
            'end 2019/1/28
            
            'Added by Morgan 2019/1/28 客戶函判發退回
            'Modified by Morgan 2025/9/3 請假時由職代輸入來函，但若有後續可能會需要由原管制人處理(承辦人改回原管制人) Ex:P-126378
            'strCon10 = strCon10 & " And cp83='" & Left(Combo1, 5) & "'"
            'strCon11 = strCon11 & " And cp83='" & Left(Combo1, 5) & "'"
            strCon10 = strCon10 & " And cp14='" & Left(Combo1, 5) & "'"
            strCon11 = strCon11 & " And cp14='" & Left(Combo1, 5) & "'"
            'end 2025/9/3
            'end 2019/1/28
            
            'Added by Morgan 2025/2/11
            strCon12 = strCon12 & " and cp14='" & Left(Combo1, 5) & "'"
            'end 2025/2/11
        End If
      Else
        If Left(Combo1, 1) = "1" Then '台灣案
            strCon = strCon & " And PA09='000'"
            strCon5 = strCon5 & " And SP09='000'"
            strNP = strNP & " And PA09='000'"
            strNP2 = strNP2 & " And SP09='000'"
            strBase = strBase & " And PA09='000'"
            strBase2 = strBase2 & " And SP09='000'"
            strCon2 = strCon2 & " and PA09='000'"
            strCon6 = strCon6 & " and SP09='000'"
            strCon3 = strCon3 & " and PA09='000'"
            strCon4 = strCon4 & " and SP09='000'"
            strCon8 = strCon8 & " and PA09='000'" 'Added by Morgan 2018/10/25
            strCon9 = strCon9 & " and SP09='000'" 'Added by Morgan 2018/10/25
            strCon10 = strCon10 & " and PA09='000'" 'Added by Morgan 2019/1/9
            strCon11 = strCon11 & " and SP09='000'" 'Added by Morgan 2019/1/9
            strCon12 = strCon12 & " and PA09='000'" 'Added by Morgan 2020/1/16
        ElseIf Left(Combo1, 1) = "2" Then '非台灣案
            strCon = strCon & " And PA09<>'000'"
            strCon5 = strCon5 & " And SP09<>'000'"
            strNP = strNP & " And PA09<>'000'"
            strNP2 = strNP2 & " And SP09<>'000'"
            strBase = strBase & " And PA09<>'000'"
            strBase2 = strBase2 & " And SP09<>'000'"
            strCon2 = strCon2 & " and PA09<>'000'"
            strCon6 = strCon6 & " and SP09<>'000'"
            strCon3 = strCon3 & " and PA09<>'000'"
            strCon4 = strCon4 & " and SP09<>'000'"
            strCon8 = strCon8 & " and PA09<>'000'" 'Added by Morgan 2018/10/25
            strCon9 = strCon9 & " and SP09<>'000'" 'Added by Morgan 2018/10/25
            strCon10 = strCon10 & " and PA09<>'000'" 'Added by Morgan 2019/1/9
            strCon11 = strCon11 & " and SP09<>'000'" 'Added by Morgan 2019/1/9
            strCon12 = strCon12 & " and PA09<>'000'" 'Added by Morgan 2020/1/16
        End If
      End If
   'Modify by Amy 2025/04/10 +FCT結案
   ElseIf m_ProState = "CFT" Or m_ProState = "T" Or m_ProState = "FCT" Then
      'Add by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
      If strSrvDate(1) >= FCP結案單電子化啟用日 Then
         strCTB = UCase(",CloseCaseMain")
         strCon = strCon & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
                                             " And ccm03 is null And length(ccm02)=9 And ccm02=cp09 And F0301=CCM01(+)"
         strCon5 = strCon5 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
                                                " And ccm03 is null And length(ccm02)=9 And ccm02=cp09 And F0301=CCM01(+)"
         strNP = strNP & " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+)" & _
                                       " And CCM03 is not null and CCM02=NP01(+) and CCM03=NP22(+) And F0301=CCM01(+)"
         strNP2 = strNP2 & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) " & _
                                             " And CCM03 is not null and CCM02=NP01(+) and CCM03=NP22(+) And F0301=CCM01(+)"
         strBase = strBase & "  and substr(CCM02,1,length(CCM02)-9)=tm01(+) and substr(CCM02,length(CCM02)-8,6)=tm02(+) and substr(CCM02,length(CCM02)-2,1)=tm03(+) and substr(CCM02,length(CCM02)-1,2)=tm04(+)"
         strBase2 = strBase2 & " and substr(CCM02,1,length(CCM02)-9)=sp01(+) and substr(CCM02,length(CCM02)-8,6)=sp02(+) and substr(CCM02,length(CCM02)-2,1)=sp03(+) and substr(CCM02,length(CCM02)-1,2)=sp04(+)"
      Else
         strCon = strCon & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
                                           " And F0304 is null and length(F0303)=9 and F0303=cp09 "
         strCon5 = strCon5 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
                                          " And F0304 is null and length(F0303)=9 and F0303=cp09 "
         strNP = strNP & " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+)" & _
                                       " And F0304 is not null and F0303=NP01(+) and F0304=NP22(+) "
         strNP2 = strNP2 & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) " & _
                                             " And F0304 is not null and F0303=NP01(+) and F0304=NP22(+) "
         strBase = strBase & " and substr(F0303,1,length(F0303)-9)=tm01(+) and substr(F0303,length(F0303)-8,6)=tm02(+) and substr(F0303,length(F0303)-2,1)=tm03(+) and substr(F0303,length(F0303)-1,2)=tm04(+)"
         strBase2 = strBase2 & " and substr(F0303,1,length(F0303)-9)=sp01(+) and substr(F0303,length(F0303)-8,6)=sp02(+) and substr(F0303,length(F0303)-2,1)=sp03(+) and substr(F0303,length(F0303)-1,2)=sp04(+)"
      End If
      strCon2 = strCon2 & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)"
      strCon6 = strCon6 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)"
      strCon3 = strCon3 & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)"
      strCon4 = strCon4 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)"
      'Added by Morgan 2021/10/14
      strCon8 = strCon8 & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01 is not null"
      strCon9 = strCon9 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp01 is not null"
      'end 2021/10/14
      
      If m_ProState = "CFT" Then
        bolRightCloser = True 'Added by Morgan 2019/5/3
        strCon = strCon & " And tm01='CFT'"
        strCon5 = strCon5 & " And sp01 in('CFC','S')"
        strNP = strNP & " And tm01='CFT'"
        strNP2 = strNP2 & " And sp01 in('CFC','S')"
        strBase = strBase & " And tm01='CFT'"
        strBase2 = strBase2 & " And sp01 in('CFC','S')"
        strCon2 = strCon2 & " and cp01='CFT'"
        strCon6 = strCon6 & " and sp01 in('CFC','S')"
        strCon3 = strCon3 & " and cp01='CFT'"
        strCon4 = strCon4 & " and sp01 in('CFC','S')"
        'Added by Morgan 2021/10/14
        strCon8 = strCon8 & " and cp01='CFT'"
        strCon9 = strCon9 & " and cp01 in('CFC','S')"
        'end 2021/10/14
        
        If Trim(Combo1.Text) <> "" Then
            strCon = strCon & " And F0308='" & Left(Combo1, 5) & "'"
            strCon5 = strCon5 & " And F0308='" & Left(Combo1, 5) & "'"
            strNP = strNP & " And F0308='" & Left(Combo1, 5) & "'"
            strNP2 = strNP2 & " And F0308='" & Left(Combo1, 5) & "'"
            strBase = strBase & " And F0308='" & Left(Combo1, 5) & "'"
            strBase2 = strBase2 & " And F0308='" & Left(Combo1, 5) & "'"
        End If
      'Modify by Amy 2025/08/29 +FCT 案件 (FCT/S 台灣,T內商結)
      'Memo 有 CFT/CFC/S 非台灣 不應該出現,若有資料再視狀況調整
      ElseIf m_ProState = "FCT" Then
         bolRightCloser = True
         strCon = strCon & " And tm01='FCT'"
         strCon5 = strCon5 & " And sp01='S'"
         strNP = strNP & " And tm01='FCT'"
         strNP2 = strNP2 & " And sp01='S'"
         strBase = strBase & " And tm01='FCT'"
         strBase2 = strBase2 & " And sp01='S'"
         strCon2 = strCon2 & " and cp01='FCT'"
         strCon6 = strCon6 & " and sp01='S'"
         strCon3 = strCon3 & " and cp01='FCT'"
         strCon4 = strCon4 & " and sp01='S'"
         strCon8 = strCon8 & " and cp01='FCT'"
         strCon9 = strCon9 & " and cp01='S'"
         
         If Trim(Combo1.Text) <> "" Then
            strCon = strCon & " And F0308='" & Left(Combo1, 5) & "'"
            strCon5 = strCon5 & " And F0308='" & Left(Combo1, 5) & "'"
            strNP = strNP & " And F0308='" & Left(Combo1, 5) & "'"
            strNP2 = strNP2 & " And F0308='" & Left(Combo1, 5) & "'"
            strBase = strBase & " And F0308='" & Left(Combo1, 5) & "'"
            strBase2 = strBase2 & " And F0308='" & Left(Combo1, 5) & "'"
        End If
      Else
        'Modify by Amy 2019/12/23 TradeMark 原:tm01='T',TF資料會出不來
        strCon = strCon & " And Substr(tm01,1,1)='T'"
        strCon5 = strCon5 & " And SubStr(sp01,1,1)='T'"
        strNP = strNP & " And Substr(tm01,1,1)='T'"
        strNP2 = strNP2 & " And SubStr(sp01,1,1)='T'"
        strBase = strBase & " And Substr(tm01,1,1)='T'"
        strBase2 = strBase2 & " And SubStr(sp01,1,1)='T'"
        strCon2 = strCon2 & " and SubStr(cp01,1,1)='T'"
        strCon6 = strCon6 & " and SubStr(cp01,1,1)='T'"
        'Modify By Sindy 2021/1/19
        'strCon3 = strCon3 & " and SubStr(cp01,1,1)='T'"
        'strCon4 = strCon4 & " and SubStr(cp01,1,1)='T'"
        strCon3 = strCon3 & " and tm01 in('T','TF')" '帳單
        strCon4 = strCon4 & " and sp01 not in('T','TF')" '帳單
        '2021/1/19 END
        'end 2019/12/23
        
        'Added by Morgan 2021/10/14
        strCon8 = strCon8 & " and SubStr(cp01,1,1)='T'"
        strCon9 = strCon9 & " and SubStr(cp01,1,1)='T'"
        'end 2021/10/14
        
        'Add by Amy 2019/12/04 客戶函判發退回
        If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            strCon10 = strCon10 & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01='T'"
            strCon11 = strCon11 & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and SubStr(sp01,1,1)='T'"
            If Trim(Combo1.Text) <> "" Then
                'Modified by Morgan 2025/9/3 請假時由職代輸入來函，但若有後續可能會需要由原管制人處理(承辦人改回原管制人) Ex:P-126378
                'strCon10 = strCon10 & " And cp83='" & Left(Combo1, 5) & "'"
                'strCon11 = strCon11 & " And cp83='" & Left(Combo1, 5) & "'"
                strCon10 = strCon10 & " And cp14='" & Left(Combo1, 5) & "'"
                strCon11 = strCon11 & " And cp14='" & Left(Combo1, 5) & "'"
                'end 2025/9/3
            End If
        End If
        'end 2019/12/04
        'Memo 2018/06/25 除法務/顧問外其他結案電子化上線前通知目前只有一人操作故不需分台灣非台灣
'        If Left(Combo1, 1) = "1" Then '台灣案
'           strCon = strCon & " And TM10='000'"
'           strCon5 = strCon5 & " And SP09='000'"
'           strNP = strNP & " And TM10='000'"
'           strNP2 = strNP2 & " And SP09='000'"
'           strBase = strBase & " And TM10='000'"
'           strBase2 = strBase2 & " And SP09='000'"
'           strCon2 = strCon2 & " and TM10='000'"
'           strCon6 = strCon6 & " and SP09='000'"
'           strCon3 = strCon3 & " and TM10='000'"
'           strCon4 = strCon4 & " and SP09='000'"
'        ElseIf Left(Combo1, 1) = "2" Then '非台灣案
'           strCon = strCon & " And TM10<>'000'"
'           strCon5 = strCon5 & " And SP09<>'000'"
'           strNP = strNP & " And TM10<>'000'"
'           strNP2 = strNP2 & " And SP09<>'000'"
'           strBase = strBase & " And TM10<>'000'"
'           strBase2 = strBase2 & " And SP09<>'000'"
'           strCon2 = strCon2 & " and TM10<>'000'"
'           strCon6 = strCon6 & " and SP09<>'000'"
'           strCon3 = strCon3 & " and TM10<>'000'"
'           strCon4 = strCon4 & " and SP09<>'000'"
'        End If
         'Add by Amy 2018/08/30 +T延展結案單
         'Modified by Morgan 2019/1/19 +oprno
         'Modify by Amy 2020/05/21 +ti06(取消延展) Null 才出現
         strT102InForm = " Union Select '延展結案' F0301,'1' F0302,'' F0309,np10 F0316,np02,np03,np04,np05,np01,np07,np08,np09,tm10,''||np22 np22,'' oprno" & _
                                 " From NextProgress,Trademark,Customer,t102InForm,CasePropertyMap Where tm01 in('T','TF') And np07 in (102,716) And np11 is null" & _
                                 " And (np01,np22) in (Select ti02,ti04 From T102InForm Where ti01<=" & strSrvDate(1) & ")" & _
                                 " And np02=tm01(+) And np03=tm02(+) And np04=tm03(+) And np05=tm04(+) And substr(tm23,1,8)=cu01(+) And substr(tm23,9,1)=cu02(+)" & _
                                 " And np01=ti02(+) And np22=ti04(+) And np02=cpm01(+) And to_char(np07)=cpm02(+) And ti06 is null"
      End If
   End If
   'end 2018/06/25
   If InStr(Combo2.Text, "全部") = 0 Then
      strCon = strCon & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
      'Add by Amy 2018/08/30
      strCon5 = strCon5 & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
      strNP = strNP & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
      strNP2 = strNP2 & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
      strBase = strBase & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
      strBase2 = strBase2 & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
      'T延展結案單
      If strT102InForm <> MsgText(601) And Trim(Left(Combo2.Text, 2)) <> "1" Then
        strT102InForm = ""
      End If
      'end 2018/08/30
      'Added by Morgan 2015/11/12"
      If Trim(Left(Combo2.Text, 2)) <> "4" Then
         strCon2 = strCon2 & " and 1=0 "
         strCon6 = strCon6 & " and 1=0 " 'Added by Morgan 2018/8/27
      End If
      'end 2015/11/12
      
      'Modified by Morgan 2018/8/24
      If Trim(Left(Combo2.Text, 2)) <> "5" Then
         strCon7 = strCon7 & " and 1=0 "
      End If
      'end 2018/2/6
      
      'Added by Morgan 2019/1/9
      If Trim(Left(Combo2.Text, 2)) <> "6" Then
         strCon10 = strCon10 & " and 1=0 "
         strCon11 = strCon11 & " and 1=0 "
      End If
      'end 2019/1/9
      
      'Added by Morgan 2018/10/25
      If Trim(Left(Combo2.Text, 2)) <> "7" Then
         strCon8 = strCon8 & " and 1=0 "
         strCon9 = strCon9 & " and 1=0 "
      End If
      'end 2018/10/25
      
      'Added by Morgan 2021/3/29
      If Trim(Left(Combo2.Text, 2)) <> "8" Then
         strCon12 = strCon12 & " and 1=0 "
      End If
      'end 2021/3/29
      
      'end 2018/2/6
   End If
   
   Combo1.Tag = Combo1 'Added by Morgan 2018/9/5
   Combo2.Tag = Combo2 'Added by Morgan 2018/9/5
   
   Screen.MousePointer = vbHourglass
   'Modified by Morgan 2015/11/12 將指示信判發退回併入
   'Modified by Morgan 2015/11/24 將指示信待寄送併入
   'strSql = "select '' V,F0301 表單編號,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,st02 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,CP09 總收文號,DECODE(PA09,'000',CPM03,CPM04) 案件性質" & _
            ",sqldatet(cp06) 本所期限,sqldatet(cp07) 法定期限,decode(F0309," & ShowFlow表單狀態中文 & ") 目前表單狀態" & _
            " from (" & _
            "select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09 from flow003,caseprogress,patent where " & strCon & " and F0304 is null and length(F0303)=9 and F0303=cp09 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,pa09 from flow003,nextprogress,patent where " & strCon & " and F0304 is not null and F0303=NP01(+) and F0304=NP22(+) and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+)" & _
            " Union select flow003.*,PA01,PA02,PA03,PA04,'','',0,0,pa09 from flow003,patent where " & strCon & " and length(F0303)>=10 and substr(F0303,1,length(F0303)-9)=pa01(+) and substr(F0303,length(F0303)-8,6)=pa02(+) and substr(F0303,length(F0303)-2,1)=pa03(+) and substr(F0303,length(F0303)-1,2)=pa04(+)" & _
            "),CASEPROPERTYMAP,staff" & _
            " where cp01=cpm01(+) and cp10=cpm02(+) and f0316=st01(+)" & _
            " order by F0301 DESC"
   'Modified by Morgan 2018/2/6 +帳單退回
   'Modified by Morgan 2018/6/22 新增 PS,CPS 帳單退回語法
   'Modify by Amy 2018/06/25 除法務/顧問外其他結案電子化
'   strSql = "select '' V,F0301 表單編號,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,st02 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,CP09 總收文號,DECODE(PA09,'000',CPM03,CPM04) 案件性質" & _
'            ",sqldatet(cp06) 本所期限,sqldatet(cp07) 法定期限,decode(F0309," & ShowFlow表單狀態中文 & ") 目前表單狀態, F0302 表單類別代碼, F0309 表單狀態代碼" & _
'            " from (select F0301,F0302,F0309,F0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09 from flow003,caseprogress,patent where " & strCon & " and F0304 is null and length(F0303)=9 and F0303=cp09 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
'            " Union select F0301,F0302,F0309,F0316,np02,np03,np04,np05,np01,np07,np08,np09,pa09 from flow003,nextprogress,patent where " & strCon & " and F0304 is not null and F0303=NP01(+) and F0304=NP22(+) and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+)" & _
'            " Union select F0301,F0302,F0309,F0316,PA01,PA02,PA03,PA04,'','',0,0,pa09 from flow003,patent where " & strCon & " and length(F0303)>=10 and substr(F0303,1,length(F0303)-9)=pa01(+) and substr(F0303,length(F0303)-8,6)=pa02(+) and substr(F0303,length(F0303)-2,1)=pa03(+) and substr(F0303,length(F0303)-1,2)=pa04(+)" & _
'            " Union select nvl(cp140,cp09) F0301,'4' F0302,decode(af09,null,'10','05') F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09 from appform,caseprogress,patent where af07=0 and af06 is not null and not exists(select * from casepaperpdf where cpp01=af01 and instr(upper(cpp02),'.DATA.PDF')>0) and cp09(+)=af01  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & strCon2 & _
'            " Union select nvl(cp140,cp09) F0301,'4' F0302,'11' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09 from appform,caseprogress,patent where af07>to_char(sysdate-30,'yyyymmdd') and af06 is not null and af11=0 and cp09(+)=af01  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & strCon2 & _
'            " Union select a1501 F0301,'5' F0302,'04' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09 from (select a1501,min(axf02) axf02  from acc150,staff,acc151 where a1521='R' and st01(+)=a1516 and axf01(+)=a1501 group by a1501),caseprogress,patent where cp09(+)=axf02 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & strCon3 & _
'            " Union select a1501 F0301,'5' F0302,'04' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09 from (select a1501,min(axf02) axf02  from acc150,staff,acc151 where a1521='R' and st01(+)=a1516 and axf01(+)=a1501 group by a1501),caseprogress,servicepractice where cp09(+)=axf02 and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & strCon4 & _
'            "),CASEPROPERTYMAP,staff" & _
'            " where cp01=cpm01(+) and cp10=cpm02(+) and f0316=st01(+)" & _
'            " order by F0301 DESC"
    'Modified by Morgan 2018/8/24 改帳單退回語法
    'Modify by Amy  2018/08/30 欄位加cp10/np22 及T延展結案單語法
    'Modified by Morgan 2019/1/18 總收文號->處理人員(發文人員,目前只有客戶函及指示信用),另加CP09於最後,內層+oprno
    'Modified by Morgan 2019/5/3 +oprno(處理人員ID)
    'Modified by Morgan 2020/6/29 +卷宗區指示信要排除有刪除註記的
    'Modify by Amy 2025/04/10 +FC結案,f0316,cp01
    strSql = "select '' V,F0301 表單編號,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,s1.st02 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,s2.st02 處理人員,DECODE(PA09,'000',CPM03,CPM04) 案件性質" & _
            ",sqldatet(cp06) 本所期限,sqldatet(cp07) 法定期限,Decode(F0309,null,'',decode(F0309," & ShowFlow表單狀態中文 & ")) 目前表單狀態, F0302 表單類別代碼, F0309 表單狀態代碼,cp10,np22,cp09,oprno,F0316,CP01 as SYSKIND" & _
            " from (select F0301,F0302,F0309,F0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22," & IIf(bolRightCloser = True, "F0308", "''") & " oprno from flow003,caseprogress f0" & strTB & Replace(strCTB, ",CASEPROGRESS F0", "") & " where 1=1 " & strCon & _
            " Union select F0301,F0302,F0309,F0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09,'' np22," & IIf(bolRightCloser = True, "F0308", "''") & " oprno from flow003,caseprogress f0,servicepractice" & Replace(strCTB, ",CASEPROGRESS F0", "") & " where 1=1 " & strCon5 & _
            " Union select F0301,F0302,F0309,F0316,np02,np03,np04,np05,np01,np07,np08,np09" & strField & ",'' np22," & IIf(bolRightCloser = True, "F0308", "''") & " oprno from flow003,nextprogress" & strTB & strCTB & " where 1=1 " & strNP & _
            " Union select F0301,F0302,F0309,F0316,np02,np03,np04,np05,np01,np07,np08,np09,sp09,'' np22," & IIf(bolRightCloser = True, "F0308", "''") & " oprno from flow003,nextprogress,servicepractice" & strCTB & " where 1=1 " & strNP2 & _
            " Union select F0301,F0302,F0309,F0316" & strField2 & ",'' np22," & IIf(bolRightCloser = True, "F0308", "''") & " oprno from flow003" & strTB & strCTB & " where 1=1 " & strBase & _
            " Union select F0301,F0302,F0309,F0316,SP01,SP02,SP03,SP04,'','',0,0,sp09,'' np22," & IIf(bolRightCloser = True, "F0308", "''") & " oprno from flow003,servicepractice" & strCTB & " where 1=1 " & strBase2 & _
            strT102InForm & _
            " Union select nvl(cp140,cp09) F0301,'4' F0302,decode(af09,null,'10','05') F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22,af15 oprno from appform,caseprogress" & strTB & " where af11=0 and af06 is not null and not exists(select * from casepaperpdf where cpp01=af01 and instr(upper(cpp02),'.DATA.PDF')>0 AND CPP10<>'D') and cp09(+)=af01 " & strCon2 & _
            " Union select nvl(cp140,cp09) F0301,'4' F0302,'11' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22,af15 oprno from appform,caseprogress" & strTB & " where af11=0 and af07>to_char(sysdate-30,'yyyymmdd') and af06 is not null and exists(select * from casepaperpdf where cpp01=af01 and instr(upper(cpp02),'.DATA.PDF')>0 AND CPP10<>'D') and cp09(+)=af01 " & strCon2 & _
            " Union select nvl(cp140,cp09) F0301,'4' F0302,decode(af09,null,'10','05') F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09,'' np22,af15 oprno from appform,caseprogress,servicepractice where af11=0 and af06 is not null and not exists(select * from casepaperpdf where cpp01=af01 and instr(upper(cpp02),'.DATA.PDF')>0 AND CPP10<>'D') and cp09(+)=af01 " & strCon6 & _
            " Union select nvl(cp140,cp09) F0301,'4' F0302,'11' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09,'' np22,af15 oprno from appform,caseprogress,servicepractice where af11=0 and af07>to_char(sysdate-30,'yyyymmdd') and af06 is not null  and exists(select * from casepaperpdf where cpp01=af01 and instr(upper(cpp02),'.DATA.PDF')>0 AND CPP10<>'D') and cp09(+)=af01 " & strCon6 & _
            " Union select a1501 F0301,'5' F0302,'04' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22, oprno from (select a1501,min(axf02) axf02,max(nvl(a1519,a1516)) oprno  from acc150,acc151 where a1521='R' and a1507 is null" & strCon7 & " and axf01(+)=a1501 group by a1501),caseprogress" & strTB & " where cp09(+)=axf02 " & strCon3 & _
            " Union select a1501 F0301,'5' F0302,'04' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09,'' np22, oprno from (select a1501,min(axf02) axf02,max(nvl(a1519,a1516)) oprno  from acc150,acc151 where a1521='R' and a1507 is null" & strCon7 & " and axf01(+)=a1501 group by a1501),caseprogress,servicepractice where cp09(+)=axf02 " & strCon4
   
   'Added by Morgan 2019/1/9
   '客戶函退回
   'Modify by Amy 2019/12/04 改判斷登入者部門非F開頭部門
   'If m_ProState = "CFP" Or m_ProState = "P" Then
   'Modified by Morgan 2021/8/17 +排除從CFT進來(未電子化)，否則非F部門(如電腦中心)就會當掉(strCon10,strCon11未設定)
   If m_ProState <> "CFT" And ((strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(Pub_StrUserSt03, 1) <> "F") _
    Or (strSrvDate(1) < T商標電子化第2階段啟用日 And (m_ProState = "CFP" Or m_ProState = "P"))) Then
      'Modified by Morgan 2025/9/3 請假時由職代輸入來函，但若有後續可能會需要由原管制人處理(承辦人改回原管制人) Ex:P-126378
      strSql = strSql & _
         " Union select cp09 F0301,'6' F0302,'05' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22,cp14 oprno from letterprogress,caseprogress" & strTB & " where lp07=0 and lp05=0 and lp04 is not null and lp36 is not null and not exists(select * from casepaperpdf where cpp01=lp01 and instr(upper(cpp02),'.CUS.PDF')>0) and cp09(+)=lp01 " & strCon10 & _
         " Union select cp09 F0301,'6' F0302,'05' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09,'' np22,cp14 oprno from letterprogress,caseprogress,servicepractice where lp07=0 and lp05=0 and lp04 is not null and lp36 is not null and not exists(select * from casepaperpdf where cpp01=lp01 and instr(upper(cpp02),'.CUS.PDF')>0) and cp09(+)=lp01 " & strCon11
   End If
   'end 2019/1/9
   
   'Added by Morgan 2018/10/25
   'If strSrvDate(1) >= e化客戶啟用日 Then 'Removed by Morgan 2021/3/26 有資料才會顯示，可先移除方便測試
      'Modified by Morgan 2021/10/14 +T
      If m_ProState = "CFP" Or m_ProState = "P" Or m_ProState = "T" Then
         'Modified by Morgan 2019/5/3 處理人員抓發文人
         strSql = strSql & _
               " Union select lp01 F0301,'7' F0302,'11' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22,nvl(lp38,cp83) oprno from letterprogress,caseprogress" & strTB & " where lp39=0 and cp09(+)=lp01 " & strCon8 & _
               " Union select lp01 F0301,'7' F0302,'11' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09,'' np22,nvl(lp38,cp83) oprno from letterprogress,caseprogress,servicepractice where lp39=0 and cp09(+)=lp01 " & strCon9
               
         'Added by Morgan 2021/3/26 處理人員抓寄發人
         'Removed by Morgan 2021/10/20 確收先不上(可能會改智權做)
         'Modified by Morgan 2021/11/18 先改可勾選(Demom用)
         If Check1.Value = vbChecked Then
            strSql = strSql & _
               " Union select lp01 F0301,'7' F0302,'13' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22,lp38 oprno from letterprogress,caseprogress" & strTB & " where lp39>0 and lp47=0 and cp09(+)=lp01 " & strCon8 & _
               " Union select lp01 F0301,'7' F0302,'13' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,sp09,'' np22,lp38 oprno from letterprogress,caseprogress,servicepractice where lp39>0 and lp47=0 and cp09(+)=lp01 " & strCon9
         End If
      End If
   'End If
   'end 2018/10/15
      
   'Added by Morgan 2020/1/16 C類未發文
   If m_ProState = "P" Then
      strSql = strSql & _
      " Union select cp09 F0301,'8' F0302,'12' F0309, cp13 f0316,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07" & strField & ",'' np22,cp14 oprno from caseprogress,staff" & strTB & " where cp158=0 and cp159=0 and cp09>'C' and cp09<'D' and substr(cp12,1,1)<>'F' and st01(+)=cp14 and st03='P12'" & strCon12
   End If
   'end 2020/1/16
   
   strSql = strSql & "),CASEPROPERTYMAP,staff s1,staff s2" & _
            " where cp01=cpm01(+) and cp10=cpm02(+) and s1.st01(+)=f0316 and s2.st01(+)=oprno" & _
            " order by F0301 DESC,CP01||CP02||CP03||CP04"
   'end 2018/06/25
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      'ShowNoData
      Exit Function
   End If
   
   'Add By Sindy 2015/10/16
   GRD1.Visible = False
   For i = 1 To GRD1.Rows - 1
      Call SetColColor(i)
      'Added by Morgan 2015/11/24
      'Modified by Morgan 2024/9/19
      'If GRD1.TextMatrix(i, 10) = "4" Then
      If GRD1.TextMatrix(i, idxCP09) > "C" Then
      'end 2024/9/19
         GRD1.TextMatrix(i, 6) = GRD1.TextMatrix(i, 6) & PUB_GetRelateCasePropertyName(GRD1.TextMatrix(i, idxCP09), "1")
      End If
      'end 2015/11/24
   Next i
   GRD1.Visible = True
   '2015/10/16 END
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      GRD1.Text = "V"
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2015/10/16
'判發退回以紅色標註
Private Sub SetColColor(intRow As Integer)
Dim i As Integer
   
   GRD1.row = intRow
   If GRD1.TextMatrix(intRow, 9) = "判發退回" Then
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         'Added by Morgan 2015/11/13
         '指示信退回以黃色標註
         'Modified by Morgan 2019/1/11 +客戶函退回
         If GRD1.TextMatrix(intRow, 10) = "4" Or GRD1.TextMatrix(intRow, 10) = "6" Then
            GRD1.CellBackColor = &H80FFFF
         Else
            GRD1.CellBackColor = &H8080FF
         End If
      Next i
   End If
   
End Sub

'Added by Morgan 2018/9/5
Private Sub Combo1_Click()
   'Modify by Amy 2025/04/10 +FC結案單
   If (m_ProState <> "CFP" And Left(m_ProState, 2) <> "FC") Or Combo1.Visible = False Then Exit Sub
   If Combo1.Tag <> Combo1 Then
      'Modify By Sindy 2025/6/4 下拉選單在點選資料列時,無資料不用彈訊息
      'cmdOK(1).Value = True
      QueryData
      '2025/6/4 END
   End If
End Sub

'Added by Morgan 2018/9/5
Private Sub Combo2_Click()
   'Modify by Amy 2025/04/10 +FC結案單
   If (m_ProState <> "CFP" And Left(m_ProState, 2) <> "FC") Or Combo2.Visible = False Then Exit Sub
   If Combo2.Tag <> Combo2 Then
      cmdOK(1).Value = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   '表單類別
   'Modify by Amy 2023/02/13 +Me.Name
   Call Flow_SetF0302Combo(Combo2, True, m_ProState, Me.Name)
   'Add by Amy 2025/07/08
   If Left(m_ProState, 2) = "FC" Then
      Label2.Visible = False
      Combo2.Visible = False
   End If
   
   'Modify by Amy 2018/06/25 除法務/顧問外其他結案電子化
   'Modify by Amy 2025/04/10 +FC結案單電子化
   If Left(m_ProState, 2) = "FC" Then
      Call SetFCCombo(Me.Name, Combo1, m_ProState, Label1)
   ElseIf InStr(m_ProState, "P") > 0 Then
        Call SetPatentP12Combo(Combo1, m_ProState, Label1)
   '商標相關
   Else
        Call SetTradeMarkCombo(Combo1, m_ProState, Label1)
   End If
   'end 2018/06/25
   Call QueryData
   
   'Added by Morgan 2015/11/24
   'Modify by Amy 2025/04/10 發現與frm210149_1 資料夾不同,改為一致,避免抓不到資料
   '                                                     且先判斷檔案是否開啟否則刪檔會錯-桂英(從解除期限回此支)
'    m_AttachPath = App.path & "\" & strUserNum
'    KillTemp
   'end 2015/11/24
   If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
      MsgBox "附件資料夾建立失敗" & vbCrLf & _
                     strExc(9) & vbCrLf & "請洽電腦中心!"
   End If
   If ChkOpenFile(m_AttachPath, strExc(9)) = True Then
      '刪不了暫存檔下次進入再刪
      MsgBox "檔案正在使用中,需關閉之檔案如下:" & vbCrLf & _
                     Replace(strExc(9), ";", vbCrLf)
   Else
      Call Pub_SetFilePathDelTmp("Close", 2, strExc(9), m_AttachPath)
   End If
   'end 2025/04/10
   
   
   If Pub_StrUserSt03 = "M51" Or m_ProState = "CFP" Then cmdDelete.Visible = True 'Added by Morgan 2018/9/5
   
   'Added by Morgan 2021/11/18
   If strSrvDate(1) >= e化客戶啟用日 And Pub_StrUserSt03 = "M51" Then
      Check1.Visible = True
   End If
   'end 2021/11/18
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strText As String
'   '一進入系統,檢查是否有須要開啟此作業
'   If pub_CallNextABSForm = True Then
'      strText = ChkIsAbsenceMustPro
'      Me.Hide
'      If InStr(1, strText, "C") > 0 Then
'         frm180203_1.Show
'      ElseIf InStr(1, strText, "D") > 0 Then
'         frm160102.intChoose = 1
'         frm160102.Hide
'         Call frm160102.cmdOK_Click(0)
'      Else
'         pub_CallNextABSForm = False
'      End If
'   End If
   
   'Modify by Amy 2025/04/10 發現與frm210149_1 資料夾不同,改為一致,避免抓不到資料
   '                                                     且先判斷檔案是否開啟否則刪檔會錯-桂英(從解除期限回此支)
   'PUB_ClearTempFolder App.path & "\" & strUserNum 'Added by Morgan 2024/9/19
   If ChkOpenFile(m_AttachPath, strExc(9)) = True Then
      '刪不了暫存檔下次進入再刪
      MsgBox "檔案正在使用中,需關閉之檔案如下:" & vbCrLf & _
                     Replace(strExc(9), ";", vbCrLf)
   Else
      PUB_ClearTempFolder m_AttachPath
   End If
   'end 2025/04/10
   
   Set frm210149 = Nothing
'   If pub_CallNextABSForm = False Then
'      Call Forms(0).SysStartCallForm
'   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modified by Morgan 2015/11/24 +10,11
   'Modify by Amy 2018/08/30 +cp10/np22 for T延展結案用
   '                                          0                1           2                        3                   4               5                         6            7                      8                          9                          10                    11                     12        13
   'Modified by Morgan 2019/1/18
   'arrGridHeadText = Array("V", "表單編號", "表單類別", "智權人員", "本所案號", "總收文號", "案件性質", "本所期限", "法定期限", "目前表單狀態", "表單類別代碼", "表單狀態代碼", "cp10", "np22")
   'arrGridHeadWidth = Array(200, 950, 950, 800, 1400, 950, 1180, 800, 800, 1000, 0, 0, 0, 0)
   'Modify by Amy 2025/04/10 +F0316,SYSKIND
   arrGridHeadText = Array("V", "表單編號", "表單類別", "智權人員", "本所案號", "處理人員", "案件性質", "本所期限", "法定期限", "目前表單狀態", "表單類別代碼", "表單狀態代碼", "cp10", "np22", "cp09", "F0316", "SYSKIND")
   arrGridHeadWidth = Array(200, 950, 950, 800, 1400, 950, 1180, 800, 800, 1000, 0, 0, 0, 0, 0, 0, 0)
   'end 2025/04/10
   idxCP10 = 12
   idxNP22 = 13
   idxCP09 = 14
   idxCaseNo = 4 'Added by Morgan 2020/1/16
   'end 2019/1/18
   'end 2018/08/30
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

'Add by Amy 2018/10/08 第一次會無法點選,原寫於GRD1_SelChange
Private Sub Grd1_Click()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
         Call SetColColor(GRD1.MouseRow)
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub
'Add by Amy 2025/06/20-薛經理說其他支都可點2下進明細,此支也加
Private Sub GRD1_DblClick()
   'Memo 此支有「刪指示信」鈕,與Sindy確認後和其他支一樣
   Call cmdok_Click(2)
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "表單編號" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub KillTemp()
'Mark by Amy 2025/04/10 改共用
'On Error GoTo ErrHnd
'   If Dir(m_AttachPath & "\.") <> "" Then
'      Kill m_AttachPath & "\*.*"
'   End If
'   Exit Sub
'
'ErrHnd:
'   Resume Next
End Sub

'Add by Amy 2018/06/25 設定商標人員下拉選單
Private Sub SetTradeMarkCombo(objCbo As Object, strSysID As String, objLbl As LABEL)
    Dim rsTmp As New ADODB.Recordset
    Dim intQ As Integer, strQ As String
    Dim strTmp As String, arrTmp
   
    objCbo.Clear
    If strSysID = "CFT" Then
        objLbl.Caption = "承辦人員："
        '顯示自已及職代(案件職代->人事職代)
        strTmp = GetSOAgent(2, strUserNum)
        If strTmp = MsgText(601) Then
            strTmp = GetSOAgent(3, strUserNum)
        End If
        objCbo.AddItem strUserNum & " " & GetPrjSalesNM(strUserNum)
        If strTmp <> MsgText(601) Then
            strQ = "Select st01,st02 From Staff Where St01 in('" & Replace(strTmp, ",", "','") & "') Order by St01"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strQ)
            If intI = 1 Then
                rsTmp.MoveFirst
                 Do While Not rsTmp.EOF
                    objCbo.AddItem rsTmp.Fields("st01") & " " & rsTmp.Fields("st02")
                    rsTmp.MoveNext
                 Loop
            End If
        End If
        objCbo = strUserNum & " " & GetPrjSalesNM(strUserNum)
    ElseIf strSysID = "T" Then
        'T處理人員同一個人不需區分台灣非台灣,故國家下拉隱藏
        objLbl.Visible = False
        Combo1.Visible = False
        Label2.Left = -30
        Combo2.Left = 900
    End If
End Sub

'Added by Morgan 2018/9/5
Private Function DeleteAppForm(pAF01 As String, pNo As String) As Boolean
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   
   strSql = "delete appform  where af01='" & pAF01 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   If pNo <> "" And pAF01 <> pNo Then OrderLetterFlowStatusUpdate pNo
      
   cnnConnection.CommitTrans
   DeleteAppForm = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function OpenWord(pDoc As String) As Boolean

   Dim iResumeCnt As Integer
   
On Error GoTo ErrHnd
   
   If TypeName(g_WordAp) <> "Application" Then
      Set g_WordAp = New Word.Application
   End If
   
   g_WordAp.Documents.Open pDoc
   g_WordAp.Visible = True
   OpenWord = True
   Exit Function
   
ErrHnd:
   'Resume
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤 : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Function
'Added by Morgan 2019/1/17
'開啟客戶函編輯畫面
Private Sub EditLetter(pLP01 As String, Optional pType As String = "1")
   Dim stType As String, stDoc As String
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   Dim stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String, stCP10 As String
   Dim bolOK As Boolean
   
On Error GoTo ErrHnd
   
   '檢查原始檔是否有客戶函/指示信DOC
   stSQL = "select cpf13,cp01,cp02,cp03,cp04,cp10 from CasePaperFile,caseprogress" & _
      " where cpf01='" & pLP01 & "' and cp09(+)=cpf01"
   If pType = "1" Then
      stSQL = stSQL & " and substr(upper(cpf02),-8)='.CUS.DOC'"
   Else
      stSQL = stSQL & " and substr(upper(cpf02),-9)='.DATA.DOC'"
   End If
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stDoc = App.path & "\$TEMP"
      With RsQ
      stCP01 = .Fields("cp01")
      stCP02 = .Fields("cp02")
      stCP03 = .Fields("cp03")
      stCP04 = .Fields("cp04")
      stCP10 = .Fields("cp10")
      If PUB_GetFtpFile(.Fields("cpf13"), stDoc, "CASEPAPERFILE", True) Then
         If OpenWord(stDoc) = True Then
            bolOK = True
         End If
      End If
      End With
   Else
      'Moidfy by Amy 2019/12/04 +if T商標電子化第2階段啟用日,改抓ld27副檔名
      If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
        '客戶函
        If pType = "1" Then
            stSQL = " And InStr(Upper(ld27),'CUS')>0 "
        '指示信
        Else
            stSQL = " And InStr(Upper(ld27),'DATA')>0 "
        End If
        stSQL = "select cp01,cp02,cp03,cp04,cp10,ld01,ld04,ld10,ld11,ld02,ld03 From CaseProgress,LetterDemand " & _
                    "Where cp09='" & pLP01 & "' And ld18(+)=cp09 " & stSQL & " " & _
                    "Order by LD02 Desc,LD03 Desc"
      Else
        '客戶函
        If pType = "1" Then
           stSQL = "select cp01,cp02,cp03,cp04,cp10,ld01,ld04,ld10,ld11,ld02,ld03 from caseprogress,letterdemand" & _
              " where cp09='" & pLP01 & "' and ld18(+)=cp09 and ld12='1' ORDER BY LD02 DESC,LD03 DESC"
        '指示信
        Else
           If m_ProState = "CFP" Then
              stSQL = "select cp01,cp02,cp03,cp04,cp10,ld01,ld04,ld10,ld11,ld02,ld03 from caseprogress,letterdemand where cp09='" & pLP01 & "' and ld18(+)=cp09 and ld05='CFP' and ld12 in ('3','4') ORDER BY LD02 DESC,LD03 DESC"
           Else
              stSQL = "select cp01,cp02,cp03,cp04,cp10,ld01,ld04,ld10,ld11,ld02,ld03 from caseprogress,letterdemand where cp09='" & pLP01 & "' and ld18(+)=cp09 and ld05='P' and ld12='6' ORDER BY LD02 DESC,LD03 DESC"
           End If
        End If
      End If
      
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With RsQ
         stCP01 = .Fields("cp01")
         stCP02 = .Fields("cp02")
         stCP03 = .Fields("cp03")
         stCP04 = .Fields("cp04")
         stCP10 = .Fields("cp10")
         NowPrint .Fields("ld04"), .Fields("ld10"), .Fields("ld11"), True, .Fields("ld01"), , , , , , , , , False, , , , pLP01
         End With
         bolOK = True
      Else
         MsgBox "系統無定稿，請自行上傳至卷宗區!!!"
      End If
   End If
   
   If bolOK Then
      Set frm1105_1.m_PrevForm = Me
      frm1105_1.m_RecNo = pLP01
      frm1105_1.m_PdfName = PUB_CaseNo2FileName(stCP01, stCP02, stCP03, stCP04) & "." & stCP10 & "." & IIf(pType = "1", "CUS", "DATA") & ".PDF"
      frm1105_1.Show
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   Set RsQ = Nothing
End Sub

'判斷暫存檔是否開啟
Private Function ChkOpenFile(ByVal stAttachPath As String, ByRef stMsg As String) As Boolean
   Dim i As Integer
   
   ChkOpenFile = False
   stMsg = ""
   '讀取資料夾檔案
   If stAttachPath <> "" And Right(stAttachPath, "1") <> "\" Then stAttachPath = stAttachPath & "\"
      File1.path = stAttachPath
   File1.Refresh
   If File1.ListCount = 0 Then Exit Function
   
   For i = 0 To File1.ListCount - 1
      If PUB_ChkFileOpening(stAttachPath & File1.List(i), , False) = True Then
         ChkOpenFile = True
         stMsg = stMsg & ";" & File1.List(i)
      End If
   Next i
   If stMsg <> "" Then stMsg = Mid(stMsg, 2)
End Function
