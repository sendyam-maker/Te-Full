VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210124 
   BorderStyle     =   1  '單線固定
   Caption         =   "定稿報價查詢"
   ClientHeight    =   5600
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5600
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdOK 
      Caption         =   "最後收文人員(&S)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   2880
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   0
      Top             =   510
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7560
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "修改報價/收據抬頭(&E)"
      Height          =   400
      Index           =   1
      Left            =   4440
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   1875
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "確認報價(&C)"
      Height          =   400
      Index           =   0
      Left            =   6390
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   60
      Width           =   1110
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8415
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4365
      Left            =   90
      TabIndex        =   7
      Top             =   1110
      Width           =   9150
      _ExtentX        =   16157
      _ExtentY        =   7691
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   330
      Left            =   1260
      TabIndex        =   13
      Top             =   510
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2220
      TabIndex        =   12
      Top             =   510
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "輸入日期前面有 * 代表修改過報價或收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   225
      TabIndex        =   11
      Top             =   900
      Width           =   3885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "3 個工作天"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5715
      TabIndex        =   8
      Top             =   540
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "報價超過                                 將視同確認不再顯示"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4860
      TabIndex        =   10
      Top             =   630
      Width           =   4020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(  例： 106/11/20 報價， 106/11/23 不再顯示 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4860
      TabIndex        =   9
      Top             =   840
      Width           =   3630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   225
      TabIndex        =   6
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "frm210124"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/04 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblSalesName
'Memo by Lydia 2019/07/01 表單名稱:專業部定稿報價查詢=>定稿報價查詢
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/4/23
Option Explicit

Dim ii As Integer
'Add by Amy 2014/05/19
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim arrID, m_strListPer As String 'Add By Sindy 2016/5/4 'Add By Sindy 2022/2/21


Private Sub cmdConfirm_Click(Index As Integer)
Dim ii As Integer, bChecked As Boolean
   
   '2011/8/10 add by sonia 因68096的帳號,杜副總開給某些人使用,若開放確認會查不出是誰操作的故限制使用68096帳號登入者只可查詢不可做確認動作
   If strUserNum = "68096" Then
      MsgBox "使用68096帳號登入者只可查詢不可做確認動作！"
      Exit Sub
   End If
   '2011/8/10 end
   
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) <> "" Then
            bChecked = True
            Exit For
         End If
      Next
      If bChecked = False Then
         MsgBox "至少需點選一筆資料！"
         Exit Sub
      End If
   End With
   If Index = 1 Then
      GetData1
   ElseIf Index = 0 Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      If FormSave = False Then
         Screen.MousePointer = vbDefault
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
         Exit Sub
      Else
         GetData
      End If
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'cancel by sonia 2020/1/9
''2011/8/10 add by sonia
'Private Sub cmdok_Click()
'Dim i As Integer
'Dim strCust As String
'
'   Me.Enabled = False
'   Screen.MousePointer = vbHourglass
'   GrdDataList.MousePointer = flexArrowHourGlass
'   For i = 1 To GrdDataList.Rows - 1
'      GrdDataList.col = 0
'      GrdDataList.row = i
'      If Trim(GrdDataList.Text) = "V" Then
'         GrdDataList.col = 0
'         GrdDataList.Text = ""
'         GrdDataList.col = 2
'         strCust = GrdDataList.Text
'         GrdDataList.col = 22
'         '抓該客戶所有案件最後收文之智權人員,包含離職人員
'         If GrdDataList.Text <> "" Then
'            strExc(0) = "select st02 from staff,(select max(cp05||cp09||cp13) cp13 from ( " & _
'                        "      Select cp05,cp09,cp13 From patent, caseprogress Where pa26='" & GetPrjPeopleNum1(GrdDataList.Text) & "' and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and cp09<'B' " & _
'                        "union Select cp05,cp09,cp13 From trademark, caseprogress Where tm23='" & GetPrjPeopleNum1(GrdDataList.Text) & "' and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and cp09<'B' " & _
'                        "union Select cp05,cp09,cp13 From lawcase, caseprogress Where lc11='" & GetPrjPeopleNum1(GrdDataList.Text) & "' and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and cp09<'B' " & _
'                        "union Select cp05,cp09,cp13 From servicepractice, caseprogress Where sp08='" & GetPrjPeopleNum1(GrdDataList.Text) & "' and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and cp09<'B' " & _
'                        "union Select cp05,cp09,cp13 From hirecase, caseprogress Where hc05='" & GetPrjPeopleNum1(GrdDataList.Text) & "' and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp09<'B' " & _
'                        ")) aa where substr(aa.cp13,18)=st01(+) "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'                MsgBox strCust & "(" & GrdDataList.Text & ")" & "所有案件最後收文智權人員為 " & RsTemp.Fields(0).Value & " ！"
'            End If
'         End If
'      End If
'   Next i
'   GrdDataList.MousePointer = flexDefault
'   Screen.MousePointer = vbDefault
'   Me.Enabled = True
'End Sub
''2011/8/10 end
'end 2020/1/9

Public Sub cmdSearch_Click()
Dim Cancel As Boolean
   
   'Add By Sindy 2022/2/21 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus '讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
      If Combo3 = MsgText(601) Then
          Call Combo3_Validate(Cancel)
          If Cancel = True Then
              Combo3.SetFocus
              Exit Sub
          End If
      ElseIf txtSales = MsgText(601) Then
          txtSales = Mid(Combo3, 1, Val(InStr(Combo3, " ")) - 1)
      End If
   End If
   '2022/2/21 END
    
   Call txtSales_Validate(Cancel)
   If Cancel = True Then
      'Add By Sindy 2022/2/21
      '有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      '排除隱藏
      ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      '2022/2/21 END
      Exit Sub
   End If
   If CheckSales = True Then
      GetData
   End If

'cancel by sonia 2020/1/9
'   '2011/8/10 ADD BY SONIA
'   If txtSales = "68006" Or txtSales = "68096" Or strUserNum = "68006" Then
'      CmdOk.Visible = True
'   Else
'      CmdOk.Visible = False
'   End If
'   '2011/8/10 END
'end 2020/1/9
End Sub

Private Function CheckSales() As Boolean
   'Add by Amy 2014/05/19 彥葶A2023登入可不輸智權人員
   'Modify by Amy 2015/02/03 總經理業務工作代理人員
   If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) Then
        CheckSales = True
        Exit Function
   End If
   'end 2014/05/19
   If txtSales = "" Then
      If Pub_StrUserSt03 = "M51" Then
         CheckSales = True
         Exit Function
      End If
      MsgBox "智權人員不可空白！", vbExclamation
      txtSales.SetFocus
   ElseIf lblSalesName = "" Then
      MsgBox "智權人員輸入錯誤！", vbExclamation
      txtSales.SetFocus
   'Add By Sindy 2022/2/21 在m_strListPer裡的就是有權限代處理
   ElseIf Combo3.Visible = True And InStr(m_strListPer, txtSales) > 0 And m_strListPer <> "" Then
   '2022/2/21 END
   'Modify By Sindy 2022/5/12 查詢權限改使用共用檢查的函數
   Else
      'Modify By Sindy 2023/12/22 +, , bolSpecMan, strSpecCode
      If PUB_ChkSalePerLimit(txtSales, strUserNum, , bolSpecMan, strSpecCode) = False Then
         If txtSales.Visible = True Then
            txtSales.SetFocus
            txtSales_GotFocus
         End If
         Exit Function
      Else
         CheckSales = True
      End If
'   ElseIf PUB_CheckSalesRight(txtSales, True) Then
'      CheckSales = True
   '2022/5/12 END
   End If
End Function

'Add By Sindy 2022/2/21
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo3_LostFocus()
   If Trim(Combo3) <> "" And Trim(Combo3) <> "全部" Then
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   ElseIf Trim(Combo3) <> "全部" Then
      txtSales = ""
   End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
Dim strEmp As String
Dim stTmp As String
   
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      lblSalesName.Caption = GetStaffName(txtSales)
      If lblSalesName.Caption = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo3.SetFocus
         Cancel = True
      End If
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Modify By Sindy 2024/8/5 mark; 因 txtSales_Validate 會檢查相關的權限
   Else
      txtSales = ""
'   ElseIf Combo3 = MsgText(601) Then
''        '下拉選單無區主管智權人員不可為空
''        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) And Pub_StrUserSt03 <> "M51" Then
''           MsgBox "非區主管職代智權人員不可空白！"
''           Cancel = True
''           Combo3.SetFocus
''           Exit Sub
''        End If
   '2024/8/5 END
   End If
End Sub
'2022/2/21 END

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth True
   
'   'Add by Amy 2015/02/03 +總經理業務工作代理人員
'   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'       bolSpecMan = True
'       strSpecCode = "總經理業務工作代理人員"
'   'Modify by Amy 2014/05/19 開放專利處部份智權同仁資料給彥葶代為處理
'   ElseIf CheckLevel(strUserNum, "A8") = True Then
'        bolSpecMan = True
'        strSpecCode = "A8"
'   End If
'
'   If bolSpecMan = True Then
'        'Add by Amy 2015/02/03 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then txtSales.Enabled = True: txtSales = strUserNum
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True: txtSales = ""
'   Else
'        txtSales = strUserNum
'   End If
'   'end 2014/05/19
   'Modify By Sindy 2023/5/18
   Call PUB_SetFormSaleDept(strUserNum, , , , txtSales, bolSpecMan, strSpecCode, , , , , , True)
   
   'Add By Sindy 2022/2/21
   '檢查當時是否需要為他人職代
   Combo3.Clear
   'Add By Sindy 2023/5/18
   If txtSales <> strUserNum And txtSales <> "" Then
      Combo3.AddItem txtSales & " " & GetPrjSalesNM(txtSales)
   End If
   '2023/5/18 END
   Combo3.AddItem strUserNum & " " & strUserName
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
'      'Add by Amy 2020/03/25 判斷下拉選單是否有區主管
'      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
'         bolAreaMan = True
'      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Form 2.0物件無法覆蓋Form 1.0
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   '2022/2/21 END
   
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If pub_CallNextForm = True Then
'      strSql = "select * from executelog where el01='frm210134' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI <> 1 Then
         pub_CallNextForm = True
         frm210134.Show
         frm210134.cmdSearch_Click
'      End If
   End If
   Set frm210124 = Nothing
End Sub

Private Sub SetDataListWidth(Optional ByVal p_bClear As Boolean = False)
Dim i As Integer, jj As Integer
   
   With grdDataList
      If p_bClear = True Then
         .Clear
         .Rows = 2
         .Cols = 12
      End If
      .FormatString = "選|輸入日期|申請人|收據抬頭|申請國家|本所號/分所號|案件名稱|案件性質|參考|費用說明"
      i = 0
      .ColWidth(i) = 300
      .ColAlignment(i) = flexAlignCenterCenter
      i = i + 1
      .ColWidth(i) = 870
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 1250
      .ColAlignment(i) = flexAlignLeftCenter
      'Added by Morgan 2015/7/2
      i = i + 1
      .ColWidth(i) = 0
      .ColAlignment(i) = flexAlignLeftCenter
      'end 2015/7/2
      i = i + 1
      .ColWidth(i) = 670
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 1335
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 1650
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 870
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 480
      .ColAlignment(i) = flexAlignCenterCenter
      i = i + 1
      'Modify by Morgan 2009/6/19
      '.ColWidth(i) = 2655
      .ColWidth(i) = 5000
      .ColAlignment(i) = flexAlignLeftCenter
      'Modify by Morgan 2010/1/11
      'jj = .Cols - 3
      'Modify by Morgan 2010/3/25
      'jj = .Cols - 4
      jj = .Cols - 5
      For i = i + 1 To jj
         .ColWidth(i) = 0
      Next
      .ColWidth(i) = 1335
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 0
      'Add by Morgan 2010/1/11
      i = i + 1
      .ColWidth(i) = 0
      'Add by Morgan 2010/3/25
      i = i + 1
      .ColWidth(i) = 0
'      '2011/8/10 add by sonia
'      i = i + 1
'      .ColWidth(i) = 0
'      '2011/8/10 end
      .Enabled = False
   End With
End Sub

Private Sub GetData()
Dim stDate As String, stCon As String
   
   If txtSales <> "" Then
      '2010/5/13 modify by sonia考慮中所跨區帶人離職時,帶人主管要看到離職智權人員資料,故傳入操作人員部門
      'stCon = " AND NP10 IN (" & PUB_GetSalesList(txtSales) & ")"
      Select Case strUserNum
         '蔣律師,杜副總,杜燕文,劉大愛,王協理,葉經理,小真,林永生,簡協理不限制
         'modify by sonia 2014/6/9 +美珍77027,並取消蔣律師79037
         'modify by sonia 2016/2/24 +69008
         'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
         Case "68006", "74018", "79053", "71011", "67002", "69008", "65001", "71003", "69005", "77027"
         'Add by Lydia 2014/10/29
            'stCon = " AND NP10 IN (" & PUB_GetSalesList(txtSales) & ")"
            stCon = " AND NVL(NP10,CP13) IN (" & PUB_GetSalesList(txtSales) & ")"
         Case Else
            Select Case PUB_GetST05(strUserNum)
               '電腦中心,財務,總經理看全部
               '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
               Case "00", "01", "08"
                 ' stCon = " AND NP10 IN (" & PUB_GetSalesList(txtSales) & ")"
                  stCon = " AND NVL(NP10,CP13) IN (" & PUB_GetSalesList(txtSales) & ")"
               Case Else
                  'stCon = " AND NP10 IN (" & PUB_GetSalesList(txtSales, PUB_GetStaffST15(txtSales, 1), PUB_GetStaffST15(txtSales, 1)) & ")"
                  stCon = " AND NVL(NP10,CP13) IN (" & PUB_GetSalesList(txtSales, PUB_GetStaffST15(txtSales, 1), PUB_GetStaffST15(txtSales, 1)) & ")"
            End Select
      End Select
      '2010/5/13 END
   'Modify by Amy 2014/05/19
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            ''彥葶A2023代為處理A7人員
            'stCon = " and np10 in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "')"
            stCon = " and NVL(NP10,CP13) in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "')"
          'end 'Add by Lydia 2014/10/29
        End If
   'end 2014/05/19
   End If
   '系統日往前 ? 個工作天日期
   stDate = CompWorkDay(報價確認天數, strSrvDate(1), 1)
   'Modify by Morgan 2010/1/11 +欄位 LC15
   '2011/6/17 MODIFY BY SONIA 原判斷顯示案件性質抓CP10或NP07以CPM02=DECODE(NP07,'601',CP10,NP07),改為CPM02=DECODE(LC11,CP66,CP10,NP07),如此商標的註冊證及延展才能抓出正確的案件性質 CFT-010250
   '2011/8/10 MODIFY BY SONIA 加本所案號完整欄
   'Add by Lydia 2014/10/29 　LC02="0"＝＞為自動發證國家之證書號輸入(frm05010403_2)時產生，因為無下一程序所以預設LC02=NP22=0
   'Modified by Morgan 2015/7/1 +收據抬頭
   'Modified by Morgan 2020/2/14 專利取消分所號
   'Modified by Morgan 2022/4/11 CPM02=DECODE(LC11,CP66,CP10,NP07)-->CPM02=DECODE(NVL(LC19,LC11),CP66,CP10,NP07)
   strExc(0) = "SELECT SQLDATET(NVL(LC06,lc11)) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),PA26) 申請人,LC16 收據抬頭" & _
      ", NA03 申請國家, NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 案號" & _
      ", NVL(PA05,NVL(PA07,PA06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質" & _
      ", '' 參考, GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,PA26,PA09,PA08,NP07,LC11,'1',NP10,NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號,NP10||' '||ST02 智權人員,LC15,CP64 備註,np02||'-'||np03||'-'||np04||'-'||np05" & _
      " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, PATENT, CUSTOMER, Nation, CASEPROPERTYMAP, staff" & _
      " WHERE LC07 IS NULL and LC13 is null and NVL(LC06,lc11)>=" & stDate & " AND NP01(+)=LC01 AND NP22(+)=LC02" & _
      " AND CP09(+)=NP01" & stCon & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
      " AND NA01(+)=PA09 AND CPM01(+)=NP02 AND CPM02=DECODE(NVL(LC19,LC11),CP66,CP10,NP07)" & _
      " AND ST01(+)=NP10 and LC02>0 " '排除自動發證
      
  '暫保留 strExc(0) = " SELECT SQLDATET(NVL(LC06,NVL(LC19,LC11))) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),PA26) 申請人," & _
               " NA03 申請國家, NVL(PA47,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04)) 案號," & _
               " NVL(PA05,NVL(PA07,PA06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質, '' 參考," & _
               " GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,PA26,PA09,PA08,NP07,LC11, '1',CP13," & _
               " CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號,CP13||' '||ST02 智權人員," & _
               " LC15,CP64 備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04 " & _
               " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, PATENT, CUSTOMER, Nation, CASEPROPERTYMAP, staff " & _
               " Where LC07 Is Null And LC13 Is Null and NVL(LC06,lc11)>=" & stDate & " AND " & _
               " NP01(+)=LC01 AND NP22(+)=LC02 AND CP09(+)=LC01" & stCon & _
               " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL " & _
               " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) AND NA01(+)=PA09 " & _
               " AND CPM01(+)=CP01 AND ((LC02='0' AND CPM02=CP10) or CPM02=DECODE(LC11,CP66,CP10,NP07)) " & _
               " AND ST01(+)=CP13 " '當LC02='0 為自動發證國家證書報價定稿無NP
               
               '因為上面的sql的基準已改成caseprogress,所以即使自動發證國家證書報價定稿無NP22也能帶入。
               '指定結合自動發證國家證書報價定稿(frm05010403_2)
   'Modified by Morgan 2015/7/2 代碼改 3(原為 1), 案件性質改抓CP10(原固定放 601)
   'Modified by Morgan 2017/9/19 + AND NP22(+)=LC02 故意抓不到期限,否則若多筆期限時資料會重複
   'Modified by Morgan 2020/2/14 專利取消分所號
   strExc(0) = strExc(0) & " UNION ALL SELECT SQLDATET(NVL(LC06,lc11)) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),PA26) 申請人,LC16 收據抬頭," & _
               " NA03 申請國家, CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 案號," & _
               " NVL(PA05,NVL(PA07,PA06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質, '' 參考," & _
               " GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,PA26,PA09,PA08,CP10 as NP07,LC11, '3',CP13 as NP10," & _
               " CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號,CP13||' '||ST02 智權人員," & _
               " LC15,CP64 備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04 " & _
               " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, PATENT, CUSTOMER, Nation, CASEPROPERTYMAP, staff " & _
               " Where LC07 Is Null And LC13 Is Null and NVL(LC06,lc11)>=" & stDate & " AND " & _
               " NP01(+)=LC01 AND NP22(+)=LC02 AND CP09(+)=LC01 and LC02=0 " & stCon & _
               " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL " & _
               " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) AND NA01(+)=PA09 " & _
               " AND CPM01(+)=CP01 and  CPM02= CP10 AND ST01(+)=CP13 "
                              

   'Add by Morgan 2008/10/24
   If Val(strSrvDate(1)) >= 20081117 Then
   '2009/10/29 MODIFY BY SONIA 商標加大陸註冊證之領證費
   'strExc(0) = strExc(0) & " UNION ALL" & _
      " SELECT SQLDATET(NVL(LC06,lc11)) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),TM23) 申請人" & _
      ", NA03 申請國家, NVL(TM34,NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05)) 案號" & _
      ", NVL(TM05,NVL(TM07,TM06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質" & _
      ", '' 參考, GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,TM23,TM10,TM08,NP07,LC11,'2',NP10,NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號,NP10||' '||ST02 智權人員" & _
      " From LETTERCACHE, NEXTPROGRESS, TRADEMARK, CUSTOMER, Nation, CASEPROPERTYMAP, staff" & _
      " WHERE LC07 IS NULL and LC13 is null and NVL(LC06,lc11)>=" & stDate & " AND NP01(+)=LC01 AND NP22(+)=LC02" & stCon & _
      " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05 AND TM01 IS NOT NULL" & _
      " AND CU01(+)=SUBSTR(TM23,1,8) AND CU02(+)=SUBSTR(TM23,9,1)" & _
      " AND NA01(+)=TM10 AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND ST01(+)=NP10"
   'Modify by Morgan 2010/1/11 +欄位 LC15
   '2011/6/17 MODIFY BY SONIA 原判斷顯示案件性質抓CP10或NP07以CPM02=DECODE(CP01||CP10,'T1701',CP10,DECODE(CP01||CP10,'CFT1701',CP10,DECODE(CP01||CP10,'CFT1713',CP10,NP07))),改為CPM02=DECODE(LC11,CP66,CP10,NP07),如此商標的註冊證及延展才能抓出正確的案件性質 CFT-010250
   '2011/8/10 MODIFY BY SONIA 加本所案號完整欄
   'Add by Lydia 2014/10/29 　LC02="0"＝＞為自動發證國家之證書號輸入(frm05010403_2)時產生，因為無下一程序所以預設LC02=NP22=0
  'Modified by Morgan 2015/6/16 +LC02>0,自動發證也改LC02=0
  'Modified by Morgan 2022/4/11 CPM02=DECODE(LC11,CP66,CP10,NP07)-->CPM02=DECODE(NVL(LC19,LC11),CP66,CP10,NP07)
  strExc(0) = strExc(0) & " UNION ALL" & _
      " SELECT SQLDATET(NVL(LC06,lc11)) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),TM23) 申請人,LC16 收據抬頭" & _
      ", NA03 申請國家, NVL(TM34,NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05)) 案號" & _
      ", NVL(TM05,NVL(TM07,TM06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質" & _
      ", '' 參考, GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,TM23,TM10,TM08,NP07,LC11,'2',NP10,NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號,NP10||' '||ST02 智權人員,LC15,CP64 備註,np02||'-'||np03||'-'||np04||'-'||np05" & _
      " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, TRADEMARK, CUSTOMER, Nation, CASEPROPERTYMAP, staff" & _
      " WHERE LC07 IS NULL and LC13 is null and NVL(LC06,lc11)>=" & stDate & " and LC02>0 AND NP01(+)=LC01 AND NP22(+)=LC02" & _
      " AND CP09(+)=NP01" & stCon & _
      " AND TM01(+)=NP02 AND TM02(+)=NP03 AND TM03(+)=NP04 AND TM04(+)=NP05 AND TM01 IS NOT NULL" & _
      " AND CU01(+)=SUBSTR(TM23,1,8) AND CU02(+)=SUBSTR(TM23,9,1)" & _
      " AND NA01(+)=TM10 AND CPM01(+)=NP02 AND CPM02=DECODE(NVL(LC19,LC11),CP66,CP10,NP07)" & _
      " AND ST01(+)=NP10"
      
'暫保留  strExc(0) = strExc(0) & " UNION ALL" & _
             " SELECT SQLDATET(NVL(LC06,lc11)) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),TM23) 申請人," & _
             " NA03 申請國家, NVL(TM34,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04)) 案號," & _
             " NVL(TM05,NVL(TM07,TM06)) 案件名稱,DECODE(NA01,'000',CPM03,CPM04) 案件性質, '' 參考," & _
             " GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,TM23,TM10,TM08,NP07,LC11,'2',CP13," & _
             " CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號," & _
             " CP10||' '||ST02 智權人員,LC15,CP64 備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04 " & _
             " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, TRADEMARK, CUSTOMER, Nation, CASEPROPERTYMAP, staff " & _
             " WHERE LC07 IS NULL and LC13 is null and NVL(LC06,lc11)>=" & stDate & " AND " & _
             " NP01(+)=LC01 AND NP22(+)=LC02 AND CP09(+)=LC01 " & stCon & _
             " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL " & _
             " AND CU01(+)=SUBSTR(TM23,1,8) AND CU02(+)=SUBSTR(TM23,9,1) AND NA01(+)=TM10 " & _
             " AND CPM01(+)=CP01 AND CPM02=DECODE(LC11,CP66,CP10,NP07) AND ST01(+)=CP13 "

   '2009/10/29 END
   
   'Added by Morgan 2015/6/16 自動發證領證報價
  strExc(0) = strExc(0) & " UNION ALL" & _
      " SELECT SQLDATET(NVL(LC06,lc11)) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),TM23) 申請人,LC16 收據抬頭" & _
      ", NA03 申請國家, NVL(TM34,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04)) 案號" & _
      ", NVL(TM05,NVL(TM07,TM06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質" & _
      ", '' 參考, GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,TM23,TM10,TM08,CP10,LC11,'4',CP13 as NP10,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號,CP13||' '||ST02 智權人員,LC15,CP64 備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04" & _
      " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, TRADEMARK, CUSTOMER, Nation, CASEPROPERTYMAP, staff" & _
      " WHERE LC07 IS NULL and LC13 is null and NVL(LC06,lc11)>=" & stDate & " and LC02=0 AND NP01(+)=LC01 AND NP22(+)=LC02" & _
      " AND CP09(+)=LC01" & stCon & _
      " AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND TM01 IS NOT NULL" & _
      " AND CU01(+)=SUBSTR(TM23,1,8) AND CU02(+)=SUBSTR(TM23,9,1)" & _
      " AND NA01(+)=TM10 AND CPM01(+)=CP01 AND CPM02=CP10" & _
      " AND ST01(+)=CP13"
   'end 2015/6/16
   End If
   'Add By Sindy 2014/12/1 內商註冊證輸入frm02010404_3,TC案件有領證費的報價,也要給智權人員做報價確認
   'Modified by Morgan 2020/2/14 PS,CPS取消分所號(目前並無報價定稿)
   strExc(0) = strExc(0) & " UNION ALL" & _
      " SELECT SQLDATET(NVL(LC06,lc11)) 輸入日期,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),SP08) 申請人,LC16 收據抬頭" & _
      ", NA03 申請國家, NVL(DECODE(CP01,'P','','CPS','',SP28),CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04)) 案號" & _
      ", NVL(SP05,NVL(SP07,SP06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質" & _
      ", '' 參考, GetVarDesc(LC01,LC02) 費用說明,LC01,LC02,SP08,SP09,'',CP10,LC11,'4',CP13,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號,CP13||' '||ST02 智權人員,LC15,CP64 備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04" & _
      " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, ServicePractice, CUSTOMER, Nation, CASEPROPERTYMAP, staff" & _
      " WHERE LC07 IS NULL and LC13 is null and NVL(LC06,lc11)>=" & stDate & " AND NP01(+)=LC01 AND LC02=0" & _
      " AND CP09(+)=LC01" & stCon & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND CU01(+)=SUBSTR(SP08,1,8) AND CU02(+)=SUBSTR(SP08,9,1)" & _
      " AND NA01(+)=SP09 AND CPM01(+)=CP01 AND CPM02=CP10" & _
      " AND ST01(+)=CP13"
   '2014/12/1 END
   strExc(0) = strExc(0) & " ORDER BY np10,2,3,4,5"
  '暫保留  strExc(0) = strExc(0) & " ORDER BY CP13,2,3,4,5"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set grdDataList.Recordset = RsTemp.Clone
      SetDataListWidth
      SetFlag
      grdDataList.Enabled = True
   Else
      SetDataListWidth True
      MsgBox "無資料！"
   End If
   
End Sub

Private Sub SetFlag()
Dim ii As Integer
Dim stAppNo As String, stAppCountry As String, strAppKind As String, strProperty As String, strDate As String
Dim strKey As String, strSys As String
Dim dblYear As Double 'Add by Morgan 2010/1/11 繳費年度
Dim bolShowTitle As Boolean 'Added by Morgan 2015/7/2

   With grdDataList
      For ii = 1 To .Rows - 1
         'Modified by Morgan 2015/7/1 加收據抬頭欄位,索引+1
         strKey = .TextMatrix(ii, 10) & .TextMatrix(ii, 11)
         stAppNo = .TextMatrix(ii, 12)
         stAppCountry = .TextMatrix(ii, 13)
         strAppKind = .TextMatrix(ii, 14)
         strProperty = .TextMatrix(ii, 15)
         strDate = .TextMatrix(ii, 16)
         strSys = .TextMatrix(ii, 17)
         dblYear = Val("" & .TextMatrix(ii, 21)) 'Add by Morgan 2010/1/11 繳費年度
         'Add by Lydia 2014/10/29 　LC02="0"＝＞為自動發證國家之證書號輸入(frm05010403_2)時產生，因為無下一程序所以預設LC02=NP22=0
         'Modify By Sindy 2014/12/1
         'If .TextMatrix(ii, 10) = "0" Then strSys = "3"
         'Removed by Morgan 2015/7/2 考慮商標自動發證,改抓資料時就改代碼
         'If .TextMatrix(ii, 11) = "0" And strSys = "1" Then
         '   strSys = "3"
         'End If
         'end 2015/7/2
         '2014/12/1 END
         
         'Added by Morgan 2015/7/2
         If .TextMatrix(ii, 11) = "0" Then
            '自動發證證書費預設收據抬頭
            If .TextMatrix(ii, 3) = "" Then
               .TextMatrix(ii, 3) = GetReceiptTitle(.TextMatrix(ii, 23), stAppNo)
            End If
            bolShowTitle = True
         End If
         'end 2015/7/2
            
         If PUB_GetOldPrice(stAppNo, stAppCountry, strAppKind, strProperty, , strDate, strKey, strSys, dblYear) = True Then
            .TextMatrix(ii, 8) = "有"
         End If
         'Add by Morgan 2008/6/4 若有改過報價(只要有到修改畫面並按確認都算)加標示
         strExc(0) = "select * from lettercachevar where lcv01||lcv02='" & strKey & "' and lcv06 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            .TextMatrix(ii, 1) = "*" & .TextMatrix(ii, 1)
         End If
      Next
   End With
   
   'Added by Morgan 2015/7/2
   If bolShowTitle = True Then
      grdDataList.ColWidth(2) = 870
      grdDataList.ColWidth(3) = 870
      grdDataList.ColWidth(6) = 1200
   End If
   'end 2015/7/2
End Sub

Private Sub GetData1()
Dim stLC01 As String, stLC02 As String
Dim ii As Integer, bolRefresh As Boolean
Dim bolShowFrame1 As Boolean 'Add By Sindy 2019/9/17
Dim strOldAddr As String 'Add By Sindy 2020/7/3
Dim strNewAddr As String 'Add By Sindy 2020/7/3
   
   With grdDataList
      ii = 1
      Do While ii < .Rows
         If .TextMatrix(ii, 0) <> "" Then
            'Modified by Morgan 2015/7/1 加收據抬頭欄位,索引+1
            stLC01 = .TextMatrix(ii, 10)
            stLC02 = .TextMatrix(ii, 11)
            
            'Add By Sindy 2020/7/3 CFT變更費用
            strOldAddr = ""
            strNewAddr = ""
            strExc(0) = "SELECT *" & _
               " FROM LETTERCACHEVAR" & _
               " WHERE LCV01='" & stLC01 & "' AND LCV02 ='" & stLC02 & "'" & _
               " and LCV03 in('原註冊地址','新地址')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If RsTemp.Fields("LCV03") = "原註冊地址" Then strOldAddr = "" & RsTemp.Fields("LCV04")
                  If RsTemp.Fields("LCV03") = "新地址" Then strNewAddr = "" & RsTemp.Fields("LCV04")
                  RsTemp.MoveNext
               Loop
            End If
            '2020/7/3 END
            
            'Modify by Morgan 2008/11/18 +LCV08(智權人員是否可修改)
            strExc(0) = "SELECT nvl(a.LCV07,a.LCV03) 項目,a.LCV04||'('||b.LCV04||')' 專業部金額" & _
               ",NVL(a.LCV06,a.LCV04) 智權部金額,nvl(b.LCV06,b.LCV04) 點數,a.LCV03,a.LCV08" & _
               " FROM LETTERCACHEVAR a, LETTERCACHEVAR b" & _
               " WHERE a.LCV01='" & stLC01 & "' AND a.LCV02 ='" & stLC02 & "' AND a.LCV05='Y'" & _
               " and b.LCV01(+)=a.LCV01 and b.LCV02(+)=a.LCV02 and b.LCV03(+)=a.LCV03||'點數' order by 1*a.lcv04 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Set frm210124_1.grdDataList.Recordset = RsTemp.Clone
               frm210124_1.SetDataListWidth
               frm210124_1.m_LC01 = stLC01
               frm210124_1.m_LC02 = stLC02
               frm210124_1.m_iRowID = ii
               If .TextMatrix(ii, 8) = "有" Then
                  frm210124_1.cmdOK(3).Enabled = True
               Else
                  frm210124_1.cmdOK(3).Enabled = False
               End If
               frm210124_1.lblAppName = .TextMatrix(ii, 2)
               frm210124_1.lblCountry = .TextMatrix(ii, 4)
               frm210124_1.lblCaseNo = .TextMatrix(ii, 5)
               frm210124_1.lblCaseName = .TextMatrix(ii, 6)
               frm210124_1.lblProperty = .TextMatrix(ii, 7)
               frm210124_1.lblMemo = .TextMatrix(ii, 22) 'Add by Morgan 2010/3/25
               'Add by Morgan 2008/7/23 CFP領證的確認報價畫面加說明
               frm210124_1.lblDesc = ""
               If Left(.TextMatrix(ii, 19), 3) = "CFP" And .TextMatrix(ii, 15) = "601" Then
                  strExc(0) = "select YF08 from caseprogress,patent,patentyearfee where cp09='" & .TextMatrix(ii, 10) & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and yf01(+)=pa09 and yf02=pa08 and yf04='601' and yf05='1' and yf08 is not null"
                  'Modified by Morgan 2023/3/29
                  If strSrvDate(1) >= PA179啟用日 Then
                     strExc(0) = strExc(0) & " and yf03=decode(pa179,'1','Y00000002','3','Y00000003','Y00000000')"
                  Else
                  'end 2023/3/29
                     strExc(0) = strExc(0) & " and yf03=decode(instr(pa91,'大個體'),0,'Y00000000','Y00000002')"
                  End If 'Added by Morgan 2023/3/29
                  
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     frm210124_1.lblDesc = "報價說明：" & vbCrLf & RsTemp(0)
                  End If
               End If
               'end 2008/7/23
               
               frm210124_1.Combo1.Clear
               frm210124_1.Combo1.Enabled = False
               'Add By Sindy 2020/7/3
               If strOldAddr <> "" Or strNewAddr <> "" Then
                  frm210124_1.Frame2.Visible = True
                  frm210124_1.txtOldAddr = strOldAddr
                  frm210124_1.txtNewAddr = strNewAddr
                  frm210124_1.Width = 8000
                  'Modified by Lydia 2022/01/04 Height = 5700=>5775
                  frm210124_1.Height = 5775
               '2020/7/3 END
               'Added by Morgan 2015/7/2
               ElseIf .TextMatrix(ii, 11) = "0" Then 'LC02
                  frm210124_1.Combo1.Enabled = True
                  SetReceiptTitle frm210124_1.Combo1, .TextMatrix(ii, 12)
                  If .TextMatrix(ii, 3) <> "" Then
                     frm210124_1.Combo1 = .TextMatrix(ii, 3)
                     frm210124_1.Combo1.Tag = frm210124_1.Combo1
                  End If
                  frm210124_1.SetCheck 'Added by Morgan 2015/12/2
                  
                  'Add By Sindy 2019/9/17 若有輸入收據抬頭且字數>=4，
                  '請檢查若不存在於客戶檔及抬頭檔則加畫面讓使用者輸入收據抬頭的相關資料
                  bolShowFrame1 = False: frm210124_1.Frame1.Visible = False
                  frm210124_1.Tag = .TextMatrix(ii, 10)
                  If frm210124_1.Combo1.Enabled = True And frm210124_1.Combo1.Text <> "" Then
                     If Len(Trim(frm210124_1.Combo1.Text)) >= 4 Then
                        If PUB_ChkTitleNmExist(frm210124_1.Combo1.Text, False) = "" Then
                           bolShowFrame1 = True: frm210124_1.Frame1.Visible = True
                        End If
                     End If
                  End If
                  If bolShowFrame1 = False Then
                     frm210124_1.Width = 6615
                     'Modified by Lydia 2022/01/04 Height = 5700=>5775
                     frm210124_1.Height = 5775
                  End If
                  '2019/9/17 END
               Else
'                  frm210124_1.Combo1.Clear
'                  frm210124_1.Combo1.Enabled = False
                  frm210124_1.Width = 6615 'Add By Sindy 2019/8/26
                  'Modified by Morgan 2016/1/13
                  'frm210124_1.Height = frm210124_1.Height - 300
                  frm210124_1.Height = frm210124_1.Combo1.Top + (frm210124_1.Height - frm210124_1.ScaleHeight)
                  'end 2016/1/13
               End If
               'end 2015/7/2
               frm210124_1.Show vbModal
               Select Case Me.Tag
                  Case "0" '確定
                     bolRefresh = True
                     If ii < .Rows - 1 Then
                        .RemoveItem ii
                        ii = ii - 1
                     End If
                  Case "1" '取消
                     Exit Do
                  Case "2" '下一筆
                     grdSelected ii
               End Select
            Else
               MsgBox "無費用欄位須確認！"
            End If
         End If
         ii = ii + 1
      Loop
   End With
   If bolRefresh = True Then
      GetData
   Else
      ClearCheck
   End If
End Sub

'清除點選
Private Sub ClearCheck()
Dim ii As Integer
   
   With grdDataList
      .Visible = False
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) <> "" Then
            grdSelected ii
         End If
      Next
      .Visible = True
   End With
End Sub

Private Sub grdSelected(p_iRow As Integer)
Dim stCheck As String, lColor As Long, ii As Integer
   
   With grdDataList
      .row = p_iRow
      .col = 0
      If .Text = "" Then
         .Text = "V"
         lColor = &HFFC0C0
      Else
         .Text = ""
         lColor = &H80000018
      End If
      For ii = 1 To .Cols - 1
         .col = ii
         .CellBackColor = lColor
      Next
   End With
End Sub

Private Sub GrdDataList_Click()
Dim iRow As Integer
   
   With grdDataList
      If .MouseRow > 0 And .MouseRow < .Rows Then
         .Visible = False
         iRow = .MouseRow
         grdSelected iRow
         .Visible = True
      End If
   End With
End Sub

Private Function FormSave() As Boolean
Dim stLCV01 As String, stLCV02 As String
   
   cnnConnection.BeginTrans
On Error GoTo ErrorHandler

   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) <> "" Then
            'Modified by Morgan 2015/7/1 加收據抬頭欄位,索引+1
            stLCV01 = .TextMatrix(ii, 10)
            stLCV02 = .TextMatrix(ii, 11)
            
            strSql = "UPDATE LETTERCACHE SET LC07='" & strUserNum & "',LC08='" & strSrvDate(1) & "',LC09=TO_CHAR(SYSDATE,'HH24MISS')" & _
               " where lc01='" & stLCV01 & "' and lc02='" & stLCV02 & "'"
            cnnConnection.Execute strSql, intI
         End If
      Next
   End With
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = StaffQuery(txtSales)
   'cancel by sonia 2016/6/30 因為S29不足五碼
   'Else
   '   lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2022/2/21
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   
   'ADD BY SONIA 2016/3/30 因73009及A3030離職,客戶轉S29
   If Len(txtSales) < 4 Then
      lblSalesName = StaffQuery(txtSales)
   End If
   
   'Add by Amy 2014/05/19
   If bolSpecMan = True Then
        'Modify by Amy 2015/02/03 +總經理業務工作代理人員
        If txtSales = strUserNum Or txtSales = "" Then
            Exit Sub
        Else
            If InStr(strSpecCode, "總經理業務工作代理人員") > 0 And InStr(Pub_GetSpecMan("總經理員工編號"), txtSales) > 0 Then
                Exit Sub
            ElseIf InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), txtSales) > 0 Then
                Exit Sub
            Else
                MsgBox "您無權限查此人資料！", vbExclamation, "操作錯誤！"
                txtSales.SetFocus
                txtSales_GotFocus
                Cancel = True
                Exit Sub
            End If
        End If
        'end 2015/02/03
   End If
   'end 2014/05/19
   
   
'Removed by Morgan 2021/5/13 取消, 此處檢查不完整 cmdSearch_Click 內的 CheckSales 比較正確 Ex:82015會無法看10051
'   'Add By Sindy 2009/05/12
'   '若為帶人主管權限時,檢查其輸入的智權人員之第二級期限管制人是否為操作人員
'   'modify by sonia 2016/6/30 帶人主管條件改寫法
'   'If Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True Then
'   If Trim(txtSales) <> "" And PUB_GetST05Limits(strUserNum) = True And txtSales.Enabled = True Then
'      'Modify By Sindy 2014/8/28
'      'If txtSales <> strUserNum And PUB_GetST52(txtSales) <> strUserNum Then
'      If txtSales <> strUserNum And PUB_GetST52(txtSales, strUserNum) = False Then
'      '2014/8/28 END
'         '2011/8/10 add by sonia 中三區人員可單獨下68096看期限
'         If PUB_GetStaffST15(strUserNum, "1") = "S23" Or strUserNum = "68096" Then
'            If txtSales = "68096" Then Exit Sub
'         End If
'         '2011/8/10 end
'         MsgBox "您無權限查詢此人資料！", vbExclamation, "操作錯誤！"
'         txtSales.SetFocus
'         txtSales_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'   End If
'end 2021/5/13

End Sub

'Added by Morgan 2015/7/1
'設定抬頭選單
'Modified by Lydia 2022/01/04 As ComboBox=> Object
Public Sub SetReceiptTitle(Combo1 As Object, strCustNo As String)

   Dim StrSQLa As String, intR As Integer
   Dim rsA As ADODB.Recordset

   Combo1.Clear
   Combo1.AddItem CustomerQuery(strCustNo, 1)
   
   StrSQLa = "Select Distinct A0K04 From ACC0K0 Where A0K03='" & strCustNo & "' and a0k04<>'" & ChgSQL(Combo1) & "' and (a0k09 is null or a0k09=0) Order By 1 "
   intR = 1
   Set rsA = ClsLawReadRstMsg(intR, StrSQLa)
   If intR = 1 Then
      While Not rsA.EOF
         Combo1.AddItem "" & rsA.Fields(0).Value
         rsA.MoveNext
      Wend
   End If
   Combo1.ListIndex = 0
   Set rsA = Nothing
   
End Sub

'Added by Morgan 2015/7/1
'預設抬頭:該案號最後抬頭->該客戶最後抬頭
Private Function GetReceiptTitle(ByVal strCaseNo As String, ByVal strCustNo As String) As String

   Dim StrSQLa As String, intR As Integer
   Dim rsA As ADODB.Recordset
   
   strCaseNo = Replace(strCaseNo, "-", "")
   
   StrSQLa = "Select A0K04,A0K02 From ACC0K0,ACC0J0 Where A0J02='" & strCaseNo & "'" & _
      " and A0J13=A0K01(+) and (a0k09 is null or a0k09=0) and A0K04 is not null" & _
      " Order By A0K02 desc"
   intR = 1
   Set rsA = ClsLawReadRstMsg(intR, StrSQLa)
   If intR = 1 Then
      GetReceiptTitle = "" & rsA.Fields(0).Value
   Else
      StrSQLa = "Select A0K04 From ACC0K0 Where A0K03='" & strCustNo & "' and (a0k09 is null or a0k09=0) Order By A0K02 Desc"
      intR = 1
      Set rsA = ClsLawReadRstMsg(intR, StrSQLa)
      If intR = 1 Then
         GetReceiptTitle = "" & rsA.Fields(0).Value
      Else
         GetReceiptTitle = CustomerQuery(strCustNo, 1)
      End If
   End If
                   
End Function



