VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm077004 
   BorderStyle     =   1  '單線固定
   Caption         =   "介紹法律所案源查詢"
   ClientHeight    =   3840
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   5770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   5841.864
   Begin VB.TextBox txtcp04 
      Height          =   300
      Left            =   3780
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2340
      Width           =   375
   End
   Begin VB.TextBox txtcp03 
      Height          =   300
      Left            =   3465
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2340
      Width           =   255
   End
   Begin VB.TextBox txtcp02 
      Height          =   300
      Left            =   2550
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2340
      Width           =   855
   End
   Begin VB.TextBox txtcp01 
      Height          =   300
      Left            =   1950
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2340
      Width           =   550
   End
   Begin VB.TextBox txtQ 
      Height          =   300
      Index           =   6
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2820
      Width           =   880
   End
   Begin VB.TextBox txtQ 
      Height          =   300
      Index           =   7
      Left            =   3090
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2820
      Width           =   880
   End
   Begin VB.TextBox txtQ 
      Height          =   300
      Index           =   8
      Left            =   1950
      MaxLength       =   6
      TabIndex        =   13
      Top             =   3180
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   4350
      TabIndex        =   15
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   405
      Left            =   3300
      TabIndex        =   14
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2820
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1620
      Width           =   600
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1950
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1620
      Width           =   600
   End
   Begin VB.TextBox txtZone 
      Height          =   300
      Left            =   1950
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1260
      Width           =   345
   End
   Begin VB.TextBox txtQ 
      Height          =   300
      Index           =   1
      Left            =   3090
      MaxLength       =   7
      TabIndex        =   1
      Top             =   900
      Width           =   880
   End
   Begin VB.TextBox txtQ 
      Height          =   300
      Index           =   0
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   0
      Top             =   900
      Width           =   880
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1950
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1980
      Width           =   735
   End
   Begin MSForms.Label lblName 
      Height          =   300
      Left            =   2790
      TabIndex        =   24
      Top             =   3210
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   315
      Left            =   1950
      TabIndex        =   6
      Top             =   1980
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2880
      TabIndex        =   23
      Top             =   2010
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   180
      Left            =   750
      TabIndex        =   22
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "法務案收文日：　　　　 　－"
      Height          =   225
      Index           =   1
      Left            =   690
      TabIndex        =   21
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "法律所處理人員："
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   20
      Top             =   3210
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "介紹人："
      Height          =   225
      Index           =   5
      Left            =   1050
      TabIndex        =   19
      Top             =   2010
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "業務區：　　　　  －       　"
      Height          =   195
      Index           =   4
      Left            =   1050
      TabIndex        =   18
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "所　別：　　　（1北所 2中所 3南所 4高所）"
      Height          =   225
      Index           =   3
      Left            =   1050
      TabIndex        =   17
      Top             =   1290
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "介紹日期：    　　　 　－　　　"
      Height          =   225
      Index           =   2
      Left            =   1050
      TabIndex        =   16
      Top             =   960
      Width           =   2835
   End
End
Attribute VB_Name = "frm077004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/07 改成Form2.0 ; lblSalesName、lblName、Combo3
'Create by Sindy 2020/4/28 介紹案源查詢
Option Explicit

Dim arrID
Dim m_strListPer As String
Dim bolAreaMan As Boolean '下拉選單有區主管
Dim bolSpecMan As Boolean '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim stST05 As String, stST15 As String
Dim bolShowMsgBox As Boolean

Private Sub cmdExit_Click()
   Unload Me
End Sub

'Private Sub CountMonthToDay()
'   If Combo2.Text <> "" Then
'      If Val(Combo2.Text) > 0 Then
'         txtQ(0) = ChangeWDateStringToTString(DateAdd("d", -2, ChangeWStringToWDateString(strSrvDate(1))))
'         txtQ(1).Text = Val(Format(DateAdd("M", Val(Combo2.Text), ChangeTStringToWDateString(txtQ(0))), "YYYYMMDD")) - 19110000
'      ElseIf Val(Combo2.Text) < 0 Then
'         txtQ(1) = ChangeWDateStringToTString(DateAdd("d", -2, ChangeWStringToWDateString(strSrvDate(1))))
'         txtQ(0).Text = Val(Format(DateAdd("M", Val(Combo2.Text), ChangeTStringToWDateString(txtQ(1))), "YYYYMMDD")) - 19110000
'      End If
'      txtQ(0).Tag = txtQ(0).Text
'      txtQ(1).Tag = txtQ(1).Text
'   End If
'End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   ConstrainCheck = True
   
   If txtQ(0) = "" Then
     MsgBox "請輸入介紹日起日！", vbExclamation
     txtQ(0).SetFocus
     txtQ_GotFocus (0)
     ConstrainCheck = False
     Exit Function
   Else
     bolCancel = False
     Call txtQ_Validate(0, bolCancel)
     If bolCancel = True Then
        ConstrainCheck = False
        Exit Function
     End If
   End If
   If txtQ(1) = "" Then
     MsgBox "請輸入介紹日迄日！", vbExclamation
     txtQ(1).SetFocus
     txtQ_GotFocus (1)
     ConstrainCheck = False
     Exit Function
   Else
     bolCancel = False
     Call txtQ_Validate(1, bolCancel)
     If bolCancel = True Then
        ConstrainCheck = False
        Exit Function
     End If
   End If
   
   'Add by Amy 2020/03/25 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus 'Add By Sindy 2020/7/15 讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
      If Combo3 = MsgText(601) Then
         Call Combo3_Validate(bolCancel)
         If bolCancel = True Then
            Combo3.SetFocus
            ConstrainCheck = False
            Exit Function
         End If
      ElseIf txtSales = MsgText(601) Then
         txtSales = Mid(Combo3, 1, Val(InStr(Combo3, " ")) - 1)
      End If
   End If
   
   'Add By Sindy 2009/05/14
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      'Modify by Amy 2020/03/25 +有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      ElseIf txtSales.Enabled = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      ConstrainCheck = False
      Exit Function
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
      ConstrainCheck = False
      Exit Function
   End If
   
'   'add by sonia 2016/6/15 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
'      If txtSalesArea = "" Then
'         txtSalesArea = "S21"
'         txtSalesArea1 = "S29"
'      End If
'      If txtSalesArea <> "P11" Or txtSalesArea1 <> "P11" Then
'         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'            MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea.SetFocus
'            txtQ_GotFocus (3)
'            ConstrainCheck = False
'            Exit Function
'         End If
'         If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'            MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea1.SetFocus
'            txtQ_GotFocus (4)
'            ConstrainCheck = False
'            Exit Function
'         End If
'      Else
'         If Trim(txtSales) <> strUserNum Then
'            MsgBox "查專利處工程師時只可查自己的資料", vbExclamation
'            txtSales.SetFocus
'            txtSales_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'      End If
'   End If
'   'Added by Lydia 2020/06/15 簡協理可看所有智權人員
'   If strUserNum = "69005" Then
'      If Left(txtSalesArea, 1) <> "S" Then
'         MsgBox "業務區起始條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea.SetFocus
'         Call txtQ_GotFocus(3)
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'      If Left(txtSalesArea1, 1) <> "S" Then
'         MsgBox "業務區迄止條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea1.SetFocus
'         Call txtQ_GotFocus(4)
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'   End If
'   'end 2020/06/15
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtQ_GotFocus (3)
'      ConstrainCheck = False
'      Exit Function
'   End If
'
'   'add by nickc 2008/01/18 加入外商主管  可以輸入相同組別的
'   If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") Then
'      If Trim(txtSales) = "" Then
'         MsgBox "智權人員不可以空白！", vbExclamation, "操作錯誤！"
'         txtSales.SetFocus
'         txtSales_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(txtSales) Then
'         MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
'         txtSales.SetFocus
'         txtSales_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
End Function

Private Sub cmdQuery_Click()
Dim bolCancel As Boolean
   
   If Combo3.Visible = True Then
      bolCancel = False
      Call Combo3_Validate(bolCancel)
      If bolCancel = True Then
         Exit Sub
      End If
   End If
   
   If ConstrainCheck = True Then
      'Screen.MousePointer = vbHourglass
      bolShowMsgBox = True
      Call doQuery
      'Screen.MousePointer = vbDefault
   End If
End Sub

Public Function doQuery() As Boolean
Dim stCon As String, stConNo As String
Dim rsA As New ADODB.Recordset
Dim stIdList As String, stConId As String
Dim VatTmp As Variant, strVal As String
Dim idx As Integer
Dim stConLos1 As String, stConLos6 As String
   
On Error GoTo ErrHnd

   'LblCntTime.Caption = "執行時間：" & Format(ServerTime, "##:##:##")
   ClearQueryLog (Me.Name) 'Add By Sindy 2025/8/8 清除查詢印表記錄檔欄位
   
   'add by sonia 2016/6/7 林柄佑要控制所別
   If strUserNum = "82026" Then
      stCon = stCon & " and s2.st06 = '" & pub_strUserOffice & "'"
      stConNo = stConNo & " and s2.st06 = '" & pub_strUserOffice & "'"
      pub_QL05 = pub_QL05 & ";林柄佑控制所別：" & pub_strUserOffice 'Add By Sindy 2025/8/8
   End If
   'end 2016/6/7
   '所別
   If txtZone <> "" Then
      stCon = stCon & " and s2.st06 = '" & txtZone & "'"
      stConNo = stConNo & " and s2.st06 = '" & txtZone & "'"
      pub_QL05 = pub_QL05 & ";所別：" & txtZone 'Add By Sindy 2025/8/8
   End If
   
   '業務區
   'Modify by Amy 2019/02/12 總經理業務工作代理人員
   If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) _
      And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   
   'Modify By Sindy 98/03/11 若智權人員為80030時, 不限制區別
   'Modify By Sindy 2009/05/12 若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   '2009/12/16 MODIFY BY SONIA 加巨京專利給郭雅娟79075看,所以不限制區別
   'modify by sonia 2016/6/7 帶人主管條件改寫法
   ElseIf txtSales = "80030" Or txtSales = "79075" Or _
      (Trim(txtSales) <> "" And PUB_GetST05Limits(strUserNum) = True And txtSales.Enabled = True And _
      txtSales <> strUserNum) Then
      '不限制區別
  
   'add by sonia 2016/6/7 查自己資料不限制區別,因為有調區問題
   ElseIf txtSales = strUserNum Then
   'end 2016/6/7
   Else
      If txtSalesArea <> "" Then
         stCon = stCon & " and s2.st15>='" & txtSalesArea & "'"
         stConNo = stConNo & " and s2.st15>='" & txtSalesArea & "'"
      End If
      If txtSalesArea1 <> "" Then
         stCon = stCon & " and s2.st15<='" & txtSalesArea1 & "'"
         stConNo = stConNo & " and s2.st15<='" & txtSalesArea1 & "'"
      End If
   End If
   'Add By Sindy 2025/8/8
   If txtSalesArea <> "" Or txtSalesArea1 <> "" Then
      pub_QL05 = pub_QL05 & ";業務區：" & txtSalesArea & "-" & txtSalesArea1
   End If
   '2025/8/8 END
   
   '智權人員
   If Trim(txtSales) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txtSales & lblSalesName 'Add By Sindy 2025/8/8
      'Add by Amy 2014/05/15 +if
      'Modify by Amy 2019/02/12 總經理業務工作代理人員,可處理總經理員工編號
      If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And _
         txtSales <> strUserNum Then
         '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
         stIdList = PUB_GetSalesList(Trim(txtSales))
      Else
         'Add by Morgan 2010/1/29 若不是多員工編號時用 = 算符來加速查詢
         stIdList = PUB_GetSalesList(Trim(txtSales), txtSalesArea, txtSalesArea1, txtZone)
      End If
      'end 2014/05/15
      
'      If InStr(stIdList, ",") = 0 Then
'         stConId = " = " & stIdList & " "
'      Else
'         stConId = " in (" & stIdList & " ) "
'      End If
      
      '2010/5/10 add by sonia 因中所有跨區帶人故離職智權人員的帶人主管不考慮業務區條件
      If Pub_StrST52 Then
         stCon = "": stConNo = ""
      End If
      '2010/5/10 end
   'Modify by Amy 2014/05/15
   '智權人員 為空
   Else
      If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
         'A2023彥葶登入,未輸智權人員-設定查A7人員
         'stConId = " in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "')"
         stIdList = Replace(Pub_GetSpecMan("A7"), ";", ",")
      End If
   End If
   If stIdList <> "" Then
      stIdList = Replace(stIdList, "'", "")
      '組多人SQL
      VatTmp = Split(stIdList, ",")
      If UBound(VatTmp) = 0 Then
         strVal = " instr(LOS04,'" & VatTmp(idx) & "')>0 "
      Else
         For idx = 0 To UBound(VatTmp)
            strVal = strVal & "or instr(LOS04,'" & VatTmp(idx) & "')>0 "
         Next idx
         strVal = Mid(strVal, 3)
      End If
      stCon = stCon & " and (" & strVal & ")"
      stConNo = stConNo & " and (" & strVal & ")"
   End If
   
   '介紹日
   If txtQ(0) <> "" Then
      stCon = stCon & " and LOS12>=" & ChangeTStringToWString(txtQ(0))
      stConNo = stConNo & " and LOS12>=" & ChangeTStringToWString(txtQ(0))
      pub_QL05 = pub_QL05 & ";介紹日期：" & txtQ(0) 'Add By Sindy 2025/8/8
   End If
   If txtQ(1) <> "" Then
      stCon = stCon & " and LOS12<=" & ChangeTStringToWString(txtQ(1))
      stConNo = stConNo & " and LOS12<=" & ChangeTStringToWString(txtQ(1))
      pub_QL05 = pub_QL05 & "-" & txtQ(1) 'Add By Sindy 2025/8/8
   End If
   
   '法務案收文日
   If txtQ(6) <> "" Then
      stCon = stCon & " and c1.CP05>=" & ChangeTStringToWString(txtQ(6))
      pub_QL05 = pub_QL05 & ";法務案收文日：" & txtQ(6) 'Add By Sindy 2025/8/8
   End If
   If txtQ(7) <> "" Then
      stCon = stCon & " and c1.CP05<=" & ChangeTStringToWString(txtQ(7))
      pub_QL05 = pub_QL05 & "-" & txtQ(7) 'Add By Sindy 2025/8/8
   End If
   '抓成案資料
   If txtQ(6) <> "" Or txtQ(7) <> "" Then
      stCon = stCon & " and LOS06 is not null"
   End If
   '法律所處理人員
   If Trim(txtQ(8)) <> "" Then
      'Modified by Lydia 2020/07/22 debug
      'stCon = stCon & " and c1.CP13=" & ChangeTStringToWString(txtQ(8))
      'If txtQ(6) = "" And txtQ(7) = "" Then '放棄案源,抓放棄人員資料
      '   stConNo = stConNo & " and LOS08=" & ChangeTStringToWString(txtQ(8))
      'End If
      stCon = stCon & " and c1.cp13=" & CNULL(Trim(txtQ(8)))
      'end 2020/07/22
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Trim(txtQ(8)) & lblName 'Add By Sindy 2025/8/8
   End If
   
   'Add By Sindy 2021/2/5
   If txtcp01 <> "" And txtcp02 <> "" Then
      If txtcp03 = "" Then txtcp03 = "0"
      If txtcp04 = "" Then txtcp04 = "00"
      stConLos1 = stConLos1 & " AND c2.cp01='" & txtcp01 & "' AND c2.cp02='" & txtcp02 & "' AND c2.cp03='" & txtcp03 & "' AND c2.cp04='" & txtcp04 & "'"
      stConLos6 = stConLos6 & " AND ((c1.cp01='" & txtcp01 & "' AND c1.cp02='" & txtcp02 & "' AND c1.cp03='" & txtcp03 & "' AND c1.cp04='" & txtcp04 & "')" & _
                                    " or (c2.cp01='" & txtcp01 & "' AND c2.cp02='" & txtcp02 & "' AND c2.cp03='" & txtcp03 & "' AND c2.cp04='" & txtcp04 & "'))"
      pub_QL05 = pub_QL05 & ";" & Label3 & txtcp01 & "-" & txtcp02 & "-" & txtcp03 & "-" & txtcp04   'Add By Sindy 2025/8/8
   End If
   '2021/2/5 END
   
   '放棄的案源資料
   '1.剔除放棄人員非法律所人員的資料
   If txtQ(6) = "" And txtQ(7) = "" Then '無法務案收文日查詢,才需讀取放棄案件
      'Modify By Sindy 2025/7/23 AND LOS01=c2.CP09(+) 改使用 AND LOS15=c2.CP162(+) and instr(c2.CP01(+),'L')=0
      strSql = "SELECT '' as V,substr(sqldatet(LOS12),1,10) as 介紹日,substr(sqldatet(LOS16),1,10) as 管制日期,a0902 as 業務區,LOS04 as 介紹人,nvl(CRA07,CRA08) as 介紹客戶,substr(CRL57,1,500) as 介紹內容" & _
               ",decode(c2.cp01,null,'',c2.cp01||'-'||c2.cp02||decode(c2.cp03||c2.cp04,'000','','-'||c2.cp03||'-'||c2.cp04)) as 案源案號" & _
               ",'放棄' as 法律所案號,substr(sqldatet(LOS07),1,10) as 收文日,substr(LOS09,1,500) as 進度備註" & _
               ",s1.ST02 as 法律所業務,LOS15,LOS01,LOS06,LOS17,LOS18" & _
               " FROM LawOfficeSource,STAFF s1,STAFF s2,acc090,consultrecapp,consultrecordlist,CaseProgress c2" & _
               " WHERE LOS08 is not null AND LOS08=s1.ST01(+) AND s1.ST03 like 'L%'" & _
               " AND substr(LOS04,1,5)=s2.ST01(+) AND a0901(+)=s2.ST15" & _
               " AND LOS17=CRL01(+) AND CRL01=CRA01(+) AND CRA02(+)='1'" & stConNo & _
               " AND LOS15=c2.CP162(+) and instr(c2.CP01(+),'L')=0" & stConLos1 & _
               " union all "
   Else
      strSql = ""
   End If
   'Modify By Sindy 2023/2/9 +簽核檔 (因待收文區的資料不應該顯示) + and f0301(+)=CRL01 and (f0301 is null or f0309 is not null)
   'Modify By Sindy 2025/7/23 AND LOS01=c2.CP09(+) 改使用 AND LOS15=c2.CP162(+) and instr(c2.CP01(+),'L')=0
   strSql = strSql & " SELECT '' as V,substr(sqldatet(LOS12),1,10) as 介紹日,substr(sqldatet(LOS16),1,10) as 管制日期,a0902 as 業務區,LOS04 as 介紹人,nvl(CRA07,CRA08) as 介紹客戶,substr(CRL57,1,500) as 介紹內容" & _
            ",decode(c2.cp01,null,'',c2.cp01||'-'||c2.cp02||decode(c2.cp03||c2.cp04,'000','','-'||c2.cp03||'-'||c2.cp04)) as 案源案號" & _
            ",decode(c1.cp01,null,'',c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)) as 法律所案號,substr(sqldatet(c1.cp05),1,10) as 收文日" & _
            ",substr(decode(c1.cp09,null,'',decode(sign(instr('3,4',sk02)),1,decode(c1.cp46,19221111,'回執退件日:'||sqldatet(c1.cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(c1.cp47)||';',''),'')||c1.cp64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),c1.cp10)||')'),1,500) as 進度備註" & _
            ",s3.st02 as 法律所業務,LOS15,LOS01,LOS06,LOS17,LOS18" & _
            " FROM LawOfficeSource,STAFF s2,STAFF s3,CaseProgress c1,SystemKind,lawcase" & _
            ",consultrecapp,consultrecordlist,acc090,casepropertymap,CaseProgress c2,flow003" & _
            " WHERE LOS08 is null AND substr(LOS04,1,5)=s2.ST01(+) AND a0901(+)=s2.ST15" & _
            " AND (los17 IS NOT NULL OR los06 IS NOT NULL)" & _
            " AND LOS06=c1.cp09(+) AND c1.cp01=SK01(+)" & _
            " AND c1.cp01=LC01(+) AND c1.cp02=LC02(+) AND c1.cp03=LC03(+) AND c1.cp04=LC04(+)" & _
            " AND c1.cp01=cpm01(+) and c1.cp10=cpm02(+) AND c1.cp13=s3.st01(+)" & _
            " AND LOS17=CRL01(+) AND CRL01=CRA01(+) AND CRA02(+)='1'" & stCon & _
            " AND LOS15=c2.CP162(+) and instr(c2.CP01(+),'L')=0" & stConLos6 & _
            " and f0301(+)=CRL01 and (f0301 is null or f0309 is not null)" & _
            " order by 介紹日,業務區,介紹人"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      'Debug.Print Timer
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      'Debug.Print Timer
      'LblCntTime.Caption = LblCntTime.Caption & " ~ " & Format(ServerTime, "##:##:##") & " 共 " & .RecordCount & " 筆" 'Add By Sindy 2014/6/12
      If .RecordCount > 0 Then
         If pub_QL04 <> "" Then InsertQueryLog (.RecordCount) 'Add By Sindy 2025/8/8
         Call frm077004_1.doQuery(AdoRecordSet3)
         frm077004_1.Show
         Me.Hide
      Else
         If bolShowMsgBox = True Then
            If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/8
            MsgBox "無符合資料！", vbInformation
         End If
      End If
   End With
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Private Sub Combo2_Change()
'   Call CountMonthToDay
'End Sub
'Private Sub Combo2_Click()
'   Call CountMonthToDay
'End Sub

'Modified by Lydia 2022/01/07 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
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
Dim stTmp As String
   
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      '只能輸入下拉選單中已有的人員
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      'Modify By Sindy 2020/6/15
'      If InStr(m_strListPer, stTmp) = 0 And stTmp <> strUserNum And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "不可輸入下拉選單以外的人員！"
'         Cancel = True
'         Combo3.SetFocus
'         Exit Sub
'      End If
      
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
'        'Modify By Sindy 2020/7/14
'        'If bolAreaMan = False And Pub_StrUserSt03 <> "M51" Then
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'        '2020/7/14 END
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
   '2024/8/5 END
   End If
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
   CloseIme
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   
   txtcp01.Text = UCase(txtcp01.Text)
   If IsEmptyText(txtcp01) = False Then
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         txtcp01_GotFocus
         Exit Sub
      End If
   End If
   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
   CloseIme
End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
   CloseIme
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
   CloseIme
End Sub

Private Sub txtSales_Change()
   lblSalesName.Caption = "" 'Added by Lydia 2020/07/22
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(Trim(txtSales), True)
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_LostFocus()
   If Trim(txtSales) = "" Then
       lblSalesName = ""
   End If
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   
   If PUB_ChkLCompStaff(strUserNum) = False Then 'Added by Lydia 2025/06/09 法律所的人員,不限制權限
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
   End If 'Added by Lydia 2025/06/09
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolShowMsgBox = False
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   bolAreaMan = False
      
   'Added by Morgan 2023/1/17
   '法律所的人員,不限制權限
   If PUB_ChkLCompStaff(strUserNum) = False Then
   'end 2023/1/17
   
      '檢查當時是否需要為他人職代
      Combo3.Clear
      Combo3.AddItem strUserNum & " " & strUserName
      Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
      
   End If 'Added by Morgan 2023/1/17
   
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
      '判斷下拉選單是否有區主管
      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
         bolAreaMan = True
      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Added by Lydia 2022/01/19
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   
'   If pub_CallNextForm = True Then 'APP開啟時,自動Run
      '系統日前後2天,共5日
'      txtQ(0) = ChangeWDateStringToTString(DateAdd("d", -2, ChangeWStringToWDateString(strSrvDate(1))))
'      txtQ(1) = ChangeWDateStringToTString(DateAdd("d", 2, ChangeWStringToWDateString(strSrvDate(1))))
'   Else
'      Combo2.ListIndex = 7 '3個月
'   End If
   txtQ(0) = strSrvDate(2)
   txtQ(1) = strSrvDate(2)
   
   '法律所的人員,不限制權限
   If PUB_ChkLCompStaff(strUserNum) = False Then
      
      'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
      Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
      
   End If
   
   '記錄原操作人可以查詢的業務區及所別
   txtZone.Tag = txtZone.Text
   txtSalesArea.Tag = txtSalesArea.Text
   txtSalesArea1.Tag = txtSalesArea1.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm077004 = Nothing
End Sub

Private Sub txtQ_GotFocus(Index As Integer)
    TextInverse txtQ(Index)
End Sub

Private Sub txtQ_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
         Case 8 '法律所處理人員
            KeyAscii = UpperCase(KeyAscii)
    End Select
End Sub

Private Sub txtQ_Validate(Index As Integer, Cancel As Boolean)
Dim strMsg As String

    Select Case Index
         Case 0, 1 '介紹日期區間
            If IsEmptyText(txtQ(Index)) = False Then
                If CheckIsTaiwanDate(txtQ(Index), False) = False Then
                   strMsg = "日期格式不正確！"
                End If
            End If
            If Index = 1 And Trim(txtQ(Index)) <> "" Then
               If RunNick(txtQ(Index - 1), txtQ(Index)) = True Then
                  Cancel = True
                  Exit Sub
               End If
            End If
         Case 6, 7 '法務案收文日
            If IsEmptyText(txtQ(Index)) = False Then
                If CheckIsTaiwanDate(txtQ(Index), False) = False Then
                   strMsg = "日期格式不正確！"
                End If
            End If
         Case 8 '法律所處理人員
            If IsEmptyText(txtQ(Index)) = False Then
                lblName.Caption = GetStaffName(txtQ(Index), True)
                If lblName.Caption = "" Then
                   strMsg = "法律所處理人員代號不存在！"
                End If
            Else
                lblName.Caption = ""
            End If
    End Select
    
    If strMsg <> "" Then GoTo ExceptExit
    
    Cancel = False
    Exit Sub
    
ExceptExit:
    Cancel = True
    MsgBox strMsg, vbExclamation, "檢核資料"
    
    txtQ(Index).SetFocus
    txtQ_GotFocus (Index)
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
   CloseIme
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_Validate(Cancel As Boolean)
   If Trim(txtSalesArea1) <> "" Then
      If RunNick(txtSalesArea, txtSalesArea1) = True Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
