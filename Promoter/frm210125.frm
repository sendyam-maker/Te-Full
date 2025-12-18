VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210125 
   BorderStyle     =   1  '單線固定
   Caption         =   "來函期限查詢"
   ClientHeight    =   5610
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdOK 
      Caption         =   "最後收文人員(&S)"
      Height          =   400
      Left            =   5070
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   3
      Top             =   750
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6583
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtZone 
      Height          =   300
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2475
      TabIndex        =   2
      Top             =   435
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      Top             =   435
      Width           =   915
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "確認通知(&C)"
      Height          =   400
      Left            =   7396
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   60
      Width           =   1110
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   7444
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
      TabIndex        =   18
      Top             =   750
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
      Left            =   2250
      TabIndex        =   17
      Top             =   750
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "符號說明：●代表銷卷＊代表閉卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4905
      TabIndex        =   16
      Top             =   1050
      Width           =   2850
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
      Left            =   5760
      TabIndex        =   15
      Top             =   510
      Width           =   1350
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(  例： 97/5/1 來函， 97/5/6 不再顯示 )"
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
      Left            =   4905
      TabIndex        =   14
      Top             =   810
      Width           =   3090
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函超過                                 將視同確認不再顯示"
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
      Left            =   4905
      TabIndex        =   13
      Top             =   600
      Width           =   4020
   End
   Begin VB.Line Line2 
      X1              =   2220
      X2              =   2490
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   2295
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   210
      TabIndex        =   11
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   225
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   225
      TabIndex        =   9
      Top             =   435
      Width           =   720
   End
End
Attribute VB_Name = "frm210125"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/04 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblSalesName
'Memo by Lydia 2019/07/01 表單名稱:審查機關來函期限查詢=>來函期限查詢
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/4/23
Option Explicit

Dim ii As Integer
Dim stST05 As String, stST15 As String
'Add by Amy 2014/05/19
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/1
Dim arrID 'Add By Sindy 2022/2/21


Private Sub cmdConfirm_Click()
Dim ii As Integer, bChecked As Boolean
Dim strMsg As String 'Add by Amy 2015/02/03
   
   '2011/8/10 add by sonia 因68096的帳號,杜副總開給某些人使用,若開放確認會查不出是誰操作的故限制使用68096帳號登入者只可查詢不可做確認動作
   If strUserNum = "68096" Then
      MsgBox "使用68096帳號登入者只可查詢不可做確認動作！"
      Exit Sub
   End If
   '2011/8/10 end
   
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) <> "" Then
            'Add by Amy 2015/02/03 總經理業務工作代理人員可查詢不可確認
            'Mark by Amy 2021/10/28 拿掉 總經理業務工作代理人員可查詢不可確認,因文雄 無法確自己案子
'            If bolSpecMan = True And InStr(strSpecCode, "總經理業務工作代理人員") > 0 And InStr(Pub_GetSpecMan("總經理員工編號"), .TextMatrix(ii, 15)) = 0 Then
'                strMsg = "只可查詢不可做確認動作！"
'            Else
                bChecked = True
'            End If
            Exit For
         End If
      Next
      If bChecked = False Then
         'Add by Amy 2015/02/03 總經理業務工作代理人員可查詢不可確認
         'Mark by Amy 2021/10/28 拿掉 總經理業務工作代理人員可查詢不可確認,因文雄 無法確自己案子
'         If bolSpecMan = True And InStr(strSpecCode, "總經理業務工作代理人員") > 0 And strMsg <> "" Then
'         Else
            strMsg = "至少需點選一筆資料！"
'         End If
         MsgBox strMsg
         Exit Sub
      End If
   End With
   
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
   
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'2011/8/9 add by sonia
Private Sub cmdok_Click()
Dim i As Integer
Dim strCust As String
   
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexArrowHourGlass
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         grdDataList.col = 2
         strCust = grdDataList.Text
         grdDataList.col = 14
         '抓該客戶所有案件最後收文之智權人員,包含離職人員
         If grdDataList.Text <> "" Then
            strExc(0) = "select st02 from staff,(select max(cp05||cp09||cp13) cp13 from ( " & _
                        "      Select cp05,cp09,cp13 From patent, caseprogress Where pa26='" & GetPrjPeopleNum1(grdDataList.Text) & "' and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and cp09<'B' " & _
                        "union Select cp05,cp09,cp13 From trademark, caseprogress Where tm23='" & GetPrjPeopleNum1(grdDataList.Text) & "' and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and cp09<'B' " & _
                        "union Select cp05,cp09,cp13 From lawcase, caseprogress Where lc11='" & GetPrjPeopleNum1(grdDataList.Text) & "' and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and cp09<'B' " & _
                        "union Select cp05,cp09,cp13 From servicepractice, caseprogress Where sp08='" & GetPrjPeopleNum1(grdDataList.Text) & "' and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and cp09<'B' " & _
                        "union Select cp05,cp09,cp13 From hirecase, caseprogress Where hc05='" & GetPrjPeopleNum1(grdDataList.Text) & "' and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp09<'B' " & _
                        ")) aa where substr(aa.cp13,18)=st01(+) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                MsgBox strCust & "(" & grdDataList.Text & ")" & "所有案件最後收文智權人員為 " & RsTemp.Fields(0).Value & " ！"
            End If
         End If
      End If
   Next i
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub
'2011/8/9 end

Public Sub cmdSearch_Click()
Dim Cancel As Boolean
Dim intErrCol As Integer
   
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
   If txtSales <> "" And lblSalesName = "" Then
      MsgBox "智權人員編號輸入錯誤！"
      Exit Sub
   End If
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   'Modify By Sindy 2025/8/11 +, txtZone
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol, txtZone) = False Then
      If intErrCol = 0 Then
         txtSales.SetFocus
         txtSales_GotFocus
      ElseIf intErrCol = 1 Then
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
      Else
         txtSalesArea1.SetFocus
         txtSalesArea1_GotFocus
      End If
      Exit Sub
   End If
   
'   'add by sonia 2016/6/30 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
'      If txtSalesArea <> "P11" Or txtSalesArea1 <> "P11" Then
'         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'            MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea.SetFocus
'            txtSalesArea_GotFocus
'            Exit Sub
'         End If
'         If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'            MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea1.SetFocus
'            txtSalesArea1_GotFocus
'            Exit Sub
'         End If
'      Else
'         If Trim(txtSales) <> strUserNum Then
'            MsgBox "查專利處工程師時只可查自己的資料", vbExclamation
'            txtSales.SetFocus
'            txtSales_GotFocus
'            Exit Sub
'         End If
'      End If
'   End If
'
'   '簡金泉69005只可查北所業務區
''Removed by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
''   If strUserNum = "69005" Then
''      If txtSalesArea < "S1" Or txtSalesArea > "S19" Then
''         MsgBox "業務區起始條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''            Exit Sub
''      End If
''      If txtSalesArea1 < "S1" Or txtSalesArea1 > "S19" Then
''         MsgBox "業務區迄止條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea1.SetFocus
''         txtSalesArea1_GotFocus
''            Exit Sub
''      End If
''   End If
''end 2019/12/30
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      Exit Sub
'   End If
'   'end 2016/6/30
   
   GetData
   
   '2011/8/9 ADD BY SONIA
   'Modify By Sindy 2013/3/21 +M51
   If txtSales = "68006" Or txtSales = "68096" Or strUserNum = "68006" Or Pub_StrUserSt03 = "M51" Then
      cmdOK.Visible = True
   Else
      cmdOK.Visible = False
   End If
   '2011/8/9 END
End Sub

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
'   'Modify by Amy 2023/05/09 +stST05
'   ElseIf Combo3 = MsgText(601) And stST05 <> "00" And stST05 <> "01" And stST05 <> "08" Then
'        '下拉選單無區主管智權人員不可為空
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
   '2024/8/5 END
   End If
End Sub
'2022/2/21 END

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
   stST05 = PUB_GetST05(strUserNum)
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode, , , , , , True)
   
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If pub_CallNextForm = True Then
'      strSql = "select * from executelog where el01='frm210124' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI <> 1 Then
         pub_CallNextForm = True
         frm210124.Show
         frm210124.cmdSearch_Click
'      End If
   End If
   Set frm210125 = Nothing
End Sub

Private Sub SetDataListWidth(Optional ByVal p_bClear As Boolean = False)
Dim i As Integer
   
   With grdDataList
      If p_bClear = True Then
         .Clear
         .Rows = 2
      End If
      .FormatString = "選|智權人員|申請人|申請國家|本所號/分所號|案件名稱|輸來函日|來函性質|下一程序|本所期限|法定期限|本所案號"
      i = 0
      .ColWidth(i) = 300
      .ColAlignment(i) = flexAlignCenterCenter
      i = i + 1
      .ColWidth(i) = 670
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 1260
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 670
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 1335
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 1125
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 870
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 870
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 870
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 870
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 870
      .ColAlignment(i) = flexAlignLeftCenter
      i = i + 1
      .ColWidth(i) = 1335
      .ColAlignment(i) = flexAlignLeftCenter
      'Modify by Amy 2015/02/03 +NP10
      For i = i + 1 To .Cols - 1
          .ColWidth(i) = 0
      Next i
   End With
End Sub

Private Sub GetData()
Dim stDate As String, stCon As String, stSQL As String
   
   stDate = CompWorkDay(期限通知天數, strSrvDate(1), 1)
   stCon = " and LDI07>=" & stDate
   'Modify by Amy 2014/05/19 +開放專利處部份智權同仁資料給彥葶代為處理
   If txtSales <> "" Then
        If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
            '彥葶A2023登入不考慮業務區(因彥葶與開放的智權同仁業務區不同)
            stCon = stCon & " and np10 in (" & PUB_GetSalesList(txtSales) & ")"
        Else
            '2009/12/16 modify by sonia
            'stCon = stCon & " and np10 in (" & PUB_GetSalesList(txtSales) & ")"
            stCon = stCon & " and np10 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) & ")"
            '2009/12/16 end
        End If
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            stCon = stCon & " and np10 in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "')"
        End If
   End If
   'end 2014/05/19
    
   'add by sonia 2016/6/30 林柄佑要控制所別
   If strUserNum = "82026" Then
      stCon = stCon & " and st06 = '" & pub_strUserOffice & "'"
   End If
   'end 2016/6/30
   
   'Modify By Sindy 2009/05/12
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   'modify by sonia 2016/6/30 帶人主管條件改寫法
   'If (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
   If (Trim(txtSales) <> "" And PUB_GetST05Limits(strUserNum) = True And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
   ElseIf txtSales = "79075" Then   '2009/12/16 ADD BY SONIA 加巨京專利給郭雅娟79075看,所以不限制區別
   ElseIf Pub_StrST52 = True Then   '2010/5/12 ADD BY SONIA
   'Modify by Amy 2019/02/13 總經理業務工作代理人員,可處理總經理員工編號
   ElseIf bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) Then 'Add by Amy 2014/05/15
        '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   'add by sonia 2016/6/30 查自己資料不限制區別,因為有調區問題
   ElseIf txtSales = strUserNum Then
   'end 2016/6/30
   Else
      If txtSalesArea <> "" Then
         stCon = stCon & " and st15>='" & txtSalesArea & "'"
      End If
      If txtSalesArea1 <> "" Then
         stCon = stCon & " and st15<='" & txtSalesArea1 & "'"
      End If
   End If
   
'   If txtZone <> "" Then
'      stCon = stCon & " and st06='" & txtZone & "'"
'   End If
   
   '2011/8/9 MODIFY BY SONIA 加本所案號完整欄
   'Modify By Sindy 2013/3/21 +●代表銷卷＊代表閉卷
   'Modify by Amy 2015/02/03 +NP10
   '專利
   'modify by sonia 2019/3/22 加and cp04(+)=np05 and cp09 is not null條件,EPC核准不必出現子案期限,CFP-025661不必出現土耳其子案CFP-025661-0-38商業使用聲明期限
   stSQL = "SELECT NVL(ST02,NP10) 智權人員, NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),PA26) 申請人" & _
      ", NA03 申請國家, NVL(PA47,np02||'-'||np03||DECODE(np04||np05,'000','','-'||np04||'-'||np05))||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') 案號" & _
      ", NVL(PA05,NVL(PA07,PA06)) 案件名稱, SQLDATET(LDI07) 輸來函日, DECODE(NA01,'000',C2.CPM03,C2.CPM04) 來函性質" & _
      ", DECODE(NA01,'000',C1.CPM03,C1.CPM04) 下一程序, SQLDATET(NP08) 本所期限, SQLDATET(NP09) 法定期限,np02||'-'||np03||DECODE(np04||np05,'000','','-'||np04||'-'||np05) 本所案號,ldi01,ldi02,np02||'-'||np03||'-'||np04||'-'||np05,NP10" & _
      " From LetterDuedayInform, Nextprogress, PATENT, CUSTOMER, Nation, CASEPROPERTYMAP C1, caseprogress, CASEPROPERTYMAP C2, STAFF" & _
      " WHERE Ldi03 IS NULL AND np01(+)=ldi01 and np22(+)=ldi02" & _
      " and PA01(+)=np02 AND PA02(+)=np03 AND PA03(+)=np04 AND PA04(+)=np05 and pa01 is not null" & _
      " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
      " AND NA01(+)=PA09 AND C1.CPM01(+)=NP02 AND C1.CPM02(+)=NP07" & _
      " AND CP09(+)=NP01 and cp04(+)=np05 and cp09 is not null AND C2.CPM01(+)=CP01 AND C2.CPM02(+)=CP10" & _
      " AND ST01(+)=NP10" & stCon
   '商標
   'modify by sonia 2015/11/6 剔除程序管制期限 T-194428(催審-被異議理由)
   stSQL = stSQL & " union SELECT NVL(ST02,NP10) 智權人員, NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),TM23) 申請人" & _
      ", NA03 申請國家, NVL(TM34,np02||'-'||np03||DECODE(np04||np05,'000','','-'||np04||'-'||np05))||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') 案號" & _
      ", NVL(TM05,NVL(TM07,TM06)) 案件名稱, SQLDATET(LDI07) 輸來函日, DECODE(NA01,'000',C2.CPM03,C2.CPM04) 來函性質" & _
      ", DECODE(NA01,'000',C1.CPM03,C1.CPM04) 下一程序, SQLDATET(NP08) 本所期限, SQLDATET(NP09) 法定期限,np02||'-'||np03||DECODE(np04||np05,'000','','-'||np04||'-'||np05) 本所案號,ldi01,ldi02,np02||'-'||np03||'-'||np04||'-'||np05,NP10" & _
      " From LetterDuedayInform, Nextprogress, TRADEMARK, CUSTOMER, Nation, CASEPROPERTYMAP C1, caseprogress, CASEPROPERTYMAP C2, STAFF" & _
      " WHERE Ldi03 IS NULL AND np01(+)=ldi01 and np22(+)=ldi02" & strNpSqlOfNoSalesDuty & _
      " and TM01(+)=np02 AND TM02(+)=np03 AND TM03(+)=np04 AND TM04(+)=np05 and TM01 is not null" & _
      " AND CU01(+)=SUBSTR(TM23,1,8) AND CU02(+)=SUBSTR(TM23,9,1)" & _
      " AND NA01(+)=TM10 AND C1.CPM01(+)=NP02 AND C1.CPM02(+)=NP07" & _
      " AND CP09(+)=NP01 AND C2.CPM01(+)=CP01 AND C2.CPM02(+)=CP10" & _
      " AND ST01(+)=NP10" & stCon
   '服務
   stSQL = stSQL & " union SELECT NVL(ST02,NP10) 智權人員, NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),SP08) 申請人" & _
      ", NA03 申請國家, nvl(sp28,np02||'-'||np03||DECODE(np04||np05,'000','','-'||np04||'-'||np05))||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') 案號" & _
      ", NVL(SP05,NVL(SP07,SP06)) 案件名稱, SQLDATET(LDI07) 輸來函日, DECODE(NA01,'000',C2.CPM03,C2.CPM04) 來函性質" & _
      ", DECODE(NA01,'000',C1.CPM03,C1.CPM04) 下一程序, SQLDATET(NP08) 本所期限, SQLDATET(NP09) 法定期限,np02||'-'||np03||DECODE(np04||np05,'000','','-'||np04||'-'||np05) 本所案號,ldi01,ldi02,np02||'-'||np03||'-'||np04||'-'||np05,NP10" & _
      " From LetterDuedayInform, Nextprogress, SERVICEPRACTICE, CUSTOMER, Nation, CASEPROPERTYMAP C1, caseprogress, CASEPROPERTYMAP C2, STAFF" & _
      " WHERE Ldi03 IS NULL AND np01(+)=ldi01 and np22(+)=ldi02" & _
      " and SP01(+)=np02 AND SP02(+)=np03 AND SP03(+)=np04 AND SP04(+)=np05 and Sp01 is not null" & _
      " AND CU01(+)=SUBSTR(SP08,1,8) AND CU02(+)=SUBSTR(SP08,9,1)" & _
      " AND NA01(+)=SP09 AND C1.CPM01(+)=NP02 AND C1.CPM02(+)=NP07" & _
      " AND CP09(+)=NP01 AND C2.CPM01(+)=CP01 AND C2.CPM02(+)=CP10" & _
      " AND ST01(+)=NP10" & stCon
      
   stSQL = stSQL & " ORDER BY 1,2,3,4,5"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      Set grdDataList.Recordset = RsTemp.Clone
      SetDataListWidth
   Else
      SetDataListWidth True
      MsgBox "無資料！"
   End If
   
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
      iRow = .MouseRow
      If iRow > 0 And iRow < .Rows Then
         .Visible = False
         grdSelected iRow
         .Visible = True
      End If
   End With
End Sub

Private Function FormSave() As Boolean
Dim stLDI01 As String, stLDI02 As String
   
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) <> "" Then
            stLDI01 = .TextMatrix(ii, 12)
            stLDI02 = .TextMatrix(ii, 13)
            strSql = "UPDATE LetterDuedayInform SET LDI03='" & strUserNum & "',LDI04='" & strSrvDate(1) & "',LDI05=TO_CHAR(SYSDATE,'HH24MISS')" & _
               " where LDI01='" & stLDI01 & "' and LDI02='" & stLDI02 & "'"
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

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2022/2/21
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = StaffQuery(txtSales)
   'cancel by sonia 2016/6/30 因為S29不足五碼
   'Else
   '   lblSalesName = ""
   End If
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   If PUB_txtSales_Limit(txtSales, m_strListPer, txtZone, txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2024/9/18 END
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
