VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210136 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標著作權案件齊備管制"
   ClientHeight    =   5740
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5740
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋(&Q)"
      Height          =   315
      Left            =   4785
      TabIndex        =   11
      Top             =   1560
      Width           =   765
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "回覆補充資料(&R)"
      Height          =   375
      Index           =   1
      Left            =   5010
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   465
      Width           =   1620
   End
   Begin VB.TextBox txtRecvDate 
      Height          =   285
      Index           =   0
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   5
      Top             =   630
      Width           =   915
   End
   Begin VB.TextBox txtRecvDate 
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   6
      Top             =   630
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "收文日期："
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   630
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "所有未發文資料"
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   930
      Width           =   1875
   End
   Begin VB.OptionButton Option1 
      Caption         =   "所有未齊備資料"
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   330
      Value           =   -1  'True
      Width           =   2265
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "齊備日或急件維護(&M)"
      Height          =   375
      Index           =   0
      Left            =   6920
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   870
      Width           =   1980
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8100
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   75
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發E-Mail(&S)"
      Height          =   375
      Index           =   3
      Left            =   7790
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   465
      Width           =   1110
   End
   Begin VB.TextBox txtCU2 
      Height          =   285
      Left            =   2400
      MaxLength       =   9
      TabIndex        =   9
      Top             =   1260
      Width           =   970
   End
   Begin VB.TextBox txtCU1 
      Height          =   285
      Left            =   1305
      MaxLength       =   9
      TabIndex        =   8
      Top             =   1260
      Width           =   970
   End
   Begin VB.TextBox txtSales 
      Height          =   285
      Left            =   3375
      MaxLength       =   6
      TabIndex        =   2
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   0
      Top             =   30
      Width           =   435
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "承辦進度(&E)"
      Height          =   375
      Index           =   1
      Left            =   6660
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   465
      Width           =   1110
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   7260
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   75
      Width           =   800
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   285
      Left            =   1890
      MaxLength       =   3
      TabIndex        =   1
      Top             =   30
      Width           =   435
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3765
      Left            =   30
      TabIndex        =   18
      Top             =   1950
      Width           =   8895
      _ExtentX        =   15699
      _ExtentY        =   6632
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      Left            =   3375
      TabIndex        =   30
      Top             =   30
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
   Begin MSForms.TextBox txtCuName 
      Height          =   300
      Left            =   1680
      TabIndex        =   10
      Top             =   1560
      Width           =   3075
      VariousPropertyBits=   671107099
      MaxLength       =   40
      Size            =   "5424;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCuNam 
      Caption         =   "客戶中文名稱："
      Height          =   180
      Left            =   330
      TabIndex        =   29
      Top             =   1590
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "未發文逾承辦期限案件"
      Height          =   180
      Left            =   3960
      TabIndex        =   28
      Top             =   1110
      Width           =   1800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "    "
      Height          =   180
      Left            =   3780
      TabIndex        =   27
      Top             =   1110
      Width           =   180
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "    "
      Height          =   180
      Left            =   3780
      TabIndex        =   26
      Top             =   900
      Width           =   180
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "未發文逾指定會稿日且無會稿日案件"
      Height          =   180
      Left            =   3960
      TabIndex        =   25
      Top             =   900
      Width           =   2880
   End
   Begin VB.Line Line5 
      X1              =   2220
      X2              =   2490
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS：請自行點選資料排序條件（點選該欄位標題）"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3780
      TabIndex        =   24
      Top             =   1365
      Width           =   4860
   End
   Begin VB.Line Line3 
      X1              =   2220
      X2              =   2490
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   330
      TabIndex        =   23
      Top             =   1305
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   22
      Top             =   75
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   2430
      TabIndex        =   21
      Top             =   75
      Width           =   900
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   4350
      TabIndex        =   20
      Top             =   75
      Width           =   1710
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   105
      TabIndex        =   19
      Top             =   5430
      Width           =   45
   End
   Begin VB.Line Line2 
      X1              =   1710
      X2              =   1980
      Y1              =   150
      Y2              =   150
   End
End
Attribute VB_Name = "frm210136"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/07/15 更名「台灣商標案件齊備管制」=>「商標著作權案件齊備管制」
'Memo by Morgan 2022/1/18 改成Form2.0 (grdDataList,txtCuName,lblSalesName)
'Memo by Lydia 2019/11/06 更名「商標案件齊備管制」=>「台灣商標案件齊備管制」(P.S. 加上避免人員誤解)
'Memo by Lydia 2019/07/01 表單名稱:台灣商標案齊備日輸入=>商標案件齊備管制
'Memo by Lydia 2018/12/10 表單名稱從「台灣商標爭議案齊備日輸入」改成「台灣商標案齊備日輸入」
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Create By Sindy 2012/5/7
Option Explicit

Dim bolShowMsgBox As Boolean, bolSelData As Boolean
''紀錄作用按鍵
Public cmdState As Integer
Dim i As Integer, j As Integer
Dim isLoad As Boolean
Dim stST05 As String, stST15 As String
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim strOrderBy As String
Public m_EP06 As String, m_CP48 As String
Public strSubject As String
Public strContent As String
Public m_CP14 As String
'Add by Amy 2014/05/20
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim m_strListPer As String 'Add By Sindy 2020/7/1
Dim arrID 'Add By Sindy 2022/2/21


Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

arrGridHeadText = Array("V", "收文日", "本所案號", "案件性質", "案件名稱" _
                  , "智權人員", "承辦人", "本所期限", "法定期限", "齊備日", "承辦期限" _
                  , "指定會稿日", "是否會稿", "會稿日", "發文日", "費用", "點數" _
                  , "進度備註", "總收文號")
arrGridHeadWidth = Array(200, 850, 1200, 850, 1200 _
                     , 850, 850, 850, 850, 850, 850 _
                     , 850, 850, 850, 850, 800, 800 _
                     , 1000, 1000)
grdDataList.MergeCells = flexMergeRestrictColumns
grdDataList.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grdDataList.Cols - 1
   grdDataList.row = 0
   grdDataList.col = iRow
   grdDataList.Text = arrGridHeadText(iRow)
   grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grdDataList.CellAlignment = flexAlignLeftCenter
Next
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   ConstrainCheck = True
   
   'Add By Sindy 2022/2/21 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus '讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
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
   '2022/2/21 END
   
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
'      txtSales.SetFocus
'      txtSales_GotFocus
      'Add By Sindy 2022/2/21
      '有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      '排除隱藏
      'ElseIf txtSales.Enabled = True Then
      ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      '2022/2/21 END
      ConstrainCheck = False
      Exit Function
   End If
      
   If Option1(2).Value = True Then
      If txtRecvDate(0) = "" Then
         MsgBox "請輸入收文日期(起)！", vbExclamation
         txtRecvDate(0).SetFocus
         txtRecvDate_GotFocus (0)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtRecvDate_Validate(0, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
      If txtRecvDate(1) = "" Then
         MsgBox "請輸入收文日期(迄)！", vbExclamation
         txtRecvDate(1).SetFocus
         txtRecvDate_GotFocus (1)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtRecvDate_Validate(1, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
   End If
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol) = False Then
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
   
'   '林永生71003檢查業務區範圍
'   If strUserNum = "71003" Then
'      If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'         MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'         MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If txtSalesArea > txtSalesArea1 Then
'         MsgBox "業務區範圍條件錯誤！", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
'   '簡金泉69005檢查業務區範圍
'   If strUserNum = "69005" Then
'      'Modified by Lydia 2020/07/02 改成全所
'      'If txtSalesArea < "S1" Or txtSalesArea > "S19" Then
'      '   MsgBox "業務區起始條件錯誤！只可查北所業務區", vbExclamation
'      If txtSalesArea < "S1" Or txtSalesArea > "S49" Then
'         MsgBox "業務區起始條件錯誤！只可查智權部", vbExclamation
'      'end 2020/07/02
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'      'Modified by Lydia 2020/07/02 改成全所
'      'If txtSalesArea1 < "S1" Or txtSalesArea1 > "S19" Then
'      '   MsgBox "業務區迄止條件錯誤！只可查北所業務區", vbExclamation
'      If txtSalesArea1 < "S1" Or txtSalesArea1 > "S49" Then
'         MsgBox "業務區迄止條件錯誤！只可查智權部", vbExclamation
'      'end 2020/07/02
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'      If txtSalesArea > txtSalesArea1 Then
'         MsgBox "業務區範圍條件錯誤！", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
'   'Added by Lydia 2020/07/02 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
'      If txtSalesArea = "" Then
'         txtSalesArea = "S21"
'         txtSalesArea1 = "S29"
'      End If
'      If txtSalesArea <> "P11" Or txtSalesArea1 <> "P11" Then
'         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'            MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea.SetFocus
'            txtSalesArea_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'         If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'            MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea1.SetFocus
'            txtSalesArea1_GotFocus
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
'   'end 2020/07/02
'   '加入外商主管  可以輸入相同組別的
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
   
   '申請人
   If Trim(txtCU1) <> "" Or Trim(txtCU2) <> "" Then
      If Mid(txtCU1, 1, 6) <> Mid(txtCU2, 1, 6) Then
          MsgBox "申請人前6碼必須相同！", vbExclamation
          txtCU1.SetFocus
          txtCU1_GotFocus
          ConstrainCheck = False
          Exit Function
      End If
   End If
   
End Function

Public Function doQuery() As Boolean
Dim stCon As String, strInData As String, stCon_A As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim stIdList As String, stConId As String
Dim stCon_B As String 'Added by Lydia 2022/09/02

On Error GoTo ErrHnd
   
   doQuery = False
   stCon = "": stCon_A = ""
   stCon_B = "" 'Added by Lydia 2022/09/02
   
   '陳經理查詢所有智權人員要控制系統類別
   If strUserNum = "68005" And txtSales <> "68005" Then
      stCon = stCon & " and cp01='FCT'"
   End If
   
   '區別
   '若智權人員為80030時, 不限制區別
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   '加巨京專利給郭雅娟79075看,所以不限制區別
   If txtSales = "80030" Or txtSales = "79075" Or _
      (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
   'Add by Amy 2014/05/20
   ElseIf bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   'end 2014/05/20
   Else
      If txtSalesArea <> "" Then
         stCon = stCon & " and s1.st15||''>='" & txtSalesArea & "'" 'Modify By Sindy 2021/8/4 CP12 => s1.st15
      End If
      If txtSalesArea1 <> "" Then
         stCon = stCon & " and s1.st15||''<='" & txtSalesArea1 & "'" 'Modify By Sindy 2021/8/4 CP12 => s1.st15
      End If
   End If
   
   '智權人員
   If txtSales <> "" Then
        If (strUserNum <> "80030" And txtSales <> "80030") Then
            'Modify by Amy 2014/05/20 +if
            If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And txtSales <> strUserNum Then
                '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
                stIdList = PUB_GetSalesList(txtSales)
            Else
               'Modify By Sindy 2014/12/16
               'stIdList = PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice)
               stIdList = PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1)
               '2014/12/16 END
            End If
            'end 2014/05/20
            
            '若不是多員工編號時用 = 算符合加速查詢
            If InStr(stIdList, ",") = 0 Then
               stConId = "=" & stIdList
            Else
               stConId = "in (" & stIdList & ")"
            End If
            stCon = stCon & " and cp13 " & stConId
        Else
            '查87027陳淑芳時同時查20001台中所
            '查80030洪琬姿時同時查F4103
            If txtSales = "80030" Then
               StrSQLa = "select ST01 from STAFF where ST04<>'1' and ST03 like 'F1%' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               strInData = "'80030','F4103'"
               If rsA.RecordCount > 0 Then
                  rsA.MoveFirst
                  Do While rsA.EOF = False
                     strInData = strInData & ",'" & rsA.Fields(0).Value & "'"
                     rsA.MoveNext
                  Loop
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               stCon = stCon & " and cp13||'' IN (" & strInData & ")"
            Else
               stCon = stCon & " and cp13='" & txtSales & "'"
            End If
        End If
   'Modify by Amy 2014/05/20
   '智權人員 為空
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            'A2023彥葶登入,未輸智權人員-設定查A7人員
            stConId = " in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') "
            stCon = stCon & " and cp13 " & stConId
        End If
   'end 2014/05/20
   End If
   
   '所有未發文資料
   stCon = stCon & " and cp27 is null and cp57 is null"
   '所有未齊備資料(且未發文的)
   If Option1(0).Value = True Then
      stCon_A = stCon_A & " and (ep06 is null or ep06=0)"
   Else
      '收文日期(且未發文的)
      If Option1(2).Value = True Then
         If txtRecvDate(0) <> "" Then
            stCon = stCon & " and cp05>=" & ChangeTStringToWString(txtRecvDate(0))
         End If
         If txtRecvDate(1) <> "" Then
            stCon = stCon & " and cp05<=" & ChangeTStringToWString(txtRecvDate(1))
         End If
      End If
   End If
   
   stCon_B = stCon_A 'Added by Lydia 2022/09/02
   '申請人
   If Trim(txtCU1) <> "" Then
       txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
       txtCU2 = Mid(txtCU2 & "000000000", 1, 9)
       stCon_A = stCon_A & " and ((tm23>='" & txtCU1 & "' and tm23<='" & txtCU2 & "') or (tm78>='" & txtCU1 & "' and tm78<='" & txtCU2 & "') or (tm79>='" & txtCU1 & "' and tm79<='" & txtCU2 & "') or (tm80>='" & txtCU1 & "' and tm80<='" & txtCU2 & "') or (tm81>='" & txtCU1 & "' and tm81<='" & txtCU2 & "')) "
       'Added by Lydia 2022/09/02 著作權案
       stCon_B = stCon_B & " and ((sp08>='" & txtCU1 & "' and sp08<='" & txtCU2 & "') or (sp58>='" & txtCU1 & "' and sp58<='" & txtCU2 & "') or (sp59>='" & txtCU1 & "' and sp59<='" & txtCU2 & "') or (sp65>='" & txtCU1 & "' and sp65<='" & txtCU2 & "') or (sp66>='" & txtCU1 & "' and sp66<='" & txtCU2 & "')) "
   End If
   
'cancel by sonia 2014/6/9
'   '蔣律師要控制所別
'   If strUserNum = "79037" Then
'      stCon_A = stCon_A & " and s1.st06='" & pub_strUserOffice & "'"
'   End If
'end 2014/6/9
   
   '查詢SQL
   'TMdebate : 爭議案件性質
   '*** Memo by Amy 2015/05/21 測M51時請於 子查詢select + /*+index(caseprogress IDXCP0501092757)*/ 查 否則可能抓錯index跑很久 ***
   'Modify By Sindy 2012/6/11 +有承辦人時,業務區第一碼不為F
   'Modify By Sindy 2012/11/27 取消6/11所寫程式" and (C1.cp14 is null or (C1.cp14 is not null and substr(C1.cp12,1,1)<>'F'))"
   'Modify By Sindy 2012/11/28 排除有分案且承辦人是外商的資料" and (C1.cp14 is null or (C1.cp14 is not null and substr(s2.st03,1,1)<>'F'))"
   'Modify By Sindy 2014/6/30 +Index : IDXCP010510132757
   'Modified by Lydia 2022/07/15 T大陸案之齊備日管控: tm10='000' => tm10 in ('000','020') 、cpm03 => decode(tm10,'000',cpm03,cpm04)
   'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
   strSql = "select '' as V,sqldatet(C1.cp05) as 收文日,C1.cp01||'-'||C1.cp02||'-'||C1.cp03||'-'||C1.cp04 as 本所案號,decode(tm10,'000',cpm03,cpm04) as 案件性質,tm05 as 案件名稱,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(C1.cp06) as 本所期限,sqldatet(C1.cp07) as 法定期限,sqldatet(ep06) as 齊備日,sqldatet(C1.cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(C1.cp27) as 發文日,C1.cp16 As 費用, C1.cp18 As 點數, C1.cp64 As 進度備註, C1.cp09 As 總收文號" & _
            " from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp16,cp18,cp27,cp48,cp64 from caseprogress,staff s1" & _
            " where cp01 in('T','FCT') and cp05>=" & TMdebateStarDT & _
            " and cp10 in (" & TMdebate & ") And Not(cp01='FCT' And InStr(" & FCT_NotTMdebate & ", cp10)>0) and cp13=s1.st01(+) " & stCon & _
            " ) C1,trademark,engineerprogress,casepropertymap,staff s1,staff s2" & _
            " where C1.cp01=tm01(+) and C1.cp02=tm02(+) and C1.cp03=tm03(+) and C1.cp04=tm04(+)" & _
            " and tm10 in ('000','020') and tm29 is null and tm57 is null" & _
            " and C1.cp09=ep02(+)" & _
            " and C1.cp01=cpm01(+) and C1.cp10=cpm02(+)" & _
            " and C1.cp13=s1.st01(+)" & _
            " and C1.cp14=s2.st01(+)" & stCon_A
           ' " and (C1.cp14 is null or (C1.cp14 is not null and substr(s2.st03,1,1)<>'F'))"  'cancel by sonia 2021/10/19 因為外商林靖傑開始辦爭議案FCT-047420(異議),FCT-047243(申請意見書)
           
   'Added by Lydia 2018/12/10 開放T台灣案管控文件齊備(A類收文)
   If strSrvDate(1) >= T案收文齊備啟用日 Then
        'Modified by Lydia 2022/07/15 T大陸案之齊備日管控: tm10='000' => tm10 in ('000','020')、cpm03 => decode(tm10,'000',cpm03,cpm04)
        strSql = strSql & " Union select '' as V,sqldatet(C1.cp05) as 收文日,C1.cp01||'-'||C1.cp02||'-'||C1.cp03||'-'||C1.cp04 as 本所案號,decode(tm10,'000',cpm03,cpm04) as 案件性質,tm05 as 案件名稱,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(C1.cp06) as 本所期限,sqldatet(C1.cp07) as 法定期限,sqldatet(ep06) as 齊備日,sqldatet(C1.cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(C1.cp27) as 發文日,C1.cp16 As 費用, C1.cp18 As 點數, C1.cp64 As 進度備註, C1.cp09 As 總收文號" & _
                 " from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp16,cp18,cp27,cp48,cp64 from caseprogress,staff s1" & _
                 " where cp01 ='T' and cp05>=" & T案收文齊備啟用日 & _
                 " and cp10 not in (" & TMdebate & "," & T案收文齊備排除 & ") and substr(cp09,1,1)='A' and cp13=s1.st01(+) " & stCon & _
                 " ) C1,trademark,engineerprogress,casepropertymap,staff s1,staff s2" & _
                 " where C1.cp01=tm01(+) and C1.cp02=tm02(+) and C1.cp03=tm03(+) and C1.cp04=tm04(+)" & _
                 " and tm10 in ('000','020') and tm29 is null and tm57 is null" & _
                 " and C1.cp09=ep02(+)" & _
                 " and C1.cp01=cpm01(+) and C1.cp10=cpm02(+)" & _
                 " and C1.cp13=s1.st01(+)" & _
                 " and C1.cp14=s2.st01(+)" & stCon_A & _
                 " and (C1.cp14 is null or (C1.cp14 is not null and substr(s2.st03,1,1)<>'F'))"
   End If
   'end 2018/12/10
   'Added by Lydia 2022/07/15 TC案之文件齊備日管控: 臺灣、大陸
   'Modified by Lydia 2022/09/02 著作權案區分條件  stCon_A=>  stCon_B
   strSql = strSql & "Union select '' as V,sqldatet(C1.cp05) as 收文日,C1.cp01||'-'||C1.cp02||'-'||C1.cp03||'-'||C1.cp04 as 本所案號,decode(sp09,'000',cpm03,cpm04) as 案件性質,sp05 as 案件名稱,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(C1.cp06) as 本所期限,sqldatet(C1.cp07) as 法定期限,sqldatet(ep06) as 齊備日,sqldatet(C1.cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(C1.cp27) as 發文日,C1.cp16 As 費用, C1.cp18 As 點數, C1.cp64 As 進度備註, C1.cp09 As 總收文號" & _
            " from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp16,cp18,cp27,cp48,cp64 from caseprogress,staff s1" & _
            " where cp01 in('TC') and cp05>=20200101 " & _
            " and substr(cp09,1,1)='A' and cp13=s1.st01(+) " & stCon & _
            " ) C1,servicepractice,engineerprogress,casepropertymap,staff s1,staff s2" & _
            " where C1.cp01=sp01(+) and C1.cp02=sp02(+) and C1.cp03=sp03(+) and C1.cp04=sp04(+)" & _
            " and sp09 in ('000','020') and sp16 is null and sp61 is null" & _
            " and C1.cp09=ep02(+)" & _
            " and C1.cp01=cpm01(+) and C1.cp10=cpm02(+)" & _
            " and C1.cp13=s1.st01(+)" & _
            " and C1.cp14=s2.st01(+)" & stCon_B
   'end 2022/07/15
   
   If strOrderBy <> "" Then
      strSql = strSql & " order by " & strOrderBy
   End If
   CheckOC3
   grdDataList.Rows = 2
   grdDataList.Clear
   SetDataListWidth
   grdDataList.FixedCols = 0
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         grdDataList.FixedCols = 7
         Call SetColor
         Label3.Caption = "PS：請自行點選資料排序條件（點選該欄位標題） 共 " & .RecordCount & " 筆"
      Else
         Label3.Caption = "PS：請自行點選資料排序條件（點選該欄位標題） 共 0 筆"
         If bolShowMsgBox = True Then
            MsgBox "無符合資料！", vbInformation
         End If
      End If
   End With
   
   doQuery = True
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetColor()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   grdDataList.Visible = False
   For i = 1 To grdDataList.Rows - 1
      Call GetColor(i)
   Next i
   grdDataList.Visible = True
End Sub

Private Function GetColor(iRow As Integer) As Boolean
   GetColor = False
   
   grdDataList.row = iRow
   '未發文逾指定會稿日且無會稿日案件以淺紅色顯示
   If grdDataList.TextMatrix(iRow, 14) = "" And _
      grdDataList.TextMatrix(iRow, 11) <> "" And _
      Val(DBDATE(grdDataList.TextMatrix(iRow, 11))) < Val(strSrvDate(1)) And _
      grdDataList.TextMatrix(iRow, 13) = "" Then
      For j = 0 To grdDataList.Cols - 1
         grdDataList.col = j
         grdDataList.CellBackColor = &H8080FF '淺紅
      Next j
      GetColor = True
   '未發文逾承辦期限案件以黃色顯示
   ElseIf grdDataList.TextMatrix(iRow, 14) = "" And _
      grdDataList.TextMatrix(iRow, 10) <> "" And _
      Val(DBDATE(grdDataList.TextMatrix(iRow, 10))) < Val(strSrvDate(1)) Then
      For j = 0 To grdDataList.Cols - 1
         grdDataList.col = j
         grdDataList.CellBackColor = &HFFFF& '黃色
      Next j
      GetColor = True
   End If
End Function


'齊備日或急件維護
Private Sub cmdModify_Click(Index As Integer)
Dim intRow As Integer

On Error GoTo ErrHnd

   Me.Hide
   For intRow = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = intRow
      If Trim(grdDataList.Text) = "V" Then
         If Not IsNull(grdDataList.TextMatrix(intRow, 18)) Then
            '開啟視窗
            m_EP06 = grdDataList.TextMatrix(intRow, 9)  '齊備日
            m_CP48 = grdDataList.TextMatrix(intRow, 10) '承辦期限
            strSubject = ""
            strContent = ""
            m_CP14 = ""
            'Add By Sindy 2012/10/24
            '依使用者點選項目進入作業
            If Index = 0 Then
               frm210136_1.WorkType = "1"
               frm210136_1.Frame1.Visible = True
               frm210136_1.Label1(8).Visible = True
               'frm210136_1.Label1(9).Caption = "齊備日歷史記錄："
               frm210136_1.Frame2.Visible = False
               'Modified by Lydia 2018/12/10 台灣商標爭議案=>台灣商標案
               'Modified by Lydia 2019/07/01
               'frm210136_1.Caption = "台灣商標案齊備日或急件維護"
               'Modified by Lydia 2019/11/06
               'frm210136_1.Caption = "商標案件齊備管制或急件維護"
               'Modified by Lydia 2022/07/15
               'frm210136_1.Caption = "台灣商標案件齊備管制或急件維護"
               frm210136_1.Caption = "商標著作權案件齊備管制或急件維護"
            ElseIf Index = 1 Then
               frm210136_1.WorkType = "2"
               frm210136_1.Frame1.Visible = False
               frm210136_1.Label1(8).Visible = False
               'frm210136_1.Label1(9).Caption = "通知補充資料記錄："
               frm210136_1.Frame2.Visible = True
               'Modified by Lydia 2018/12/10 台灣商標爭議案=>台灣商標案
               'Modified by Lydia 2022/07/15
               'frm210136_1.Caption = "台灣商標案回覆補充資料作業"
               frm210136_1.Caption = "商標著作權案件回覆補充資料作業"
            End If
            '2012/10/24 End
            If frm210136_1.Process(grdDataList.TextMatrix(intRow, 18)) Then
               'Add By Sindy 2013/1/11
               If frm210136_1.bolNotData = True Then '若有待回覆補充資料時一定要進入此作業
                  Index = 1
                  frm210136_1.WorkType = "2"
                  frm210136_1.Frame1.Visible = False
                  frm210136_1.Label1(8).Visible = False
                  frm210136_1.Frame2.Visible = True
                  'Modified by Lydia 2018/12/10 台灣商標爭議案=>台灣商標案
                  'Modified by Lydia 2022/07/15
                  'frm210136_1.Caption = "台灣商標案回覆補充資料作業"
                  frm210136_1.Caption = "商標著作權案件回覆補充資料作業"
               End If
               '2013/1/11 End
               'Modify By Sindy 2012/10/24
               If Index = 1 And frm210136_1.bolNotData = False Then
                  Me.Show
                  MsgBox "無待回覆的補充資料 !!"
               Else
               '2012/10/24 End
                  frm210136_1.Show vbModal
                  grdDataList.TextMatrix(intRow, 9) = m_EP06
                  grdDataList.TextMatrix(intRow, 10) = m_CP48
                  'E-Mail通知承辦人
                  If strSubject <> "" And m_CP14 <> "" Then
                     PUB_SendMail strUserNum, m_CP14, "", strSubject, strContent, ""
                  End If
               End If
            End If
            Unload frm210136_1
            Set frm210136_1 = Nothing
            '資料列恢復原狀
            grdDataList.TextMatrix(intRow, 0) = ""
            If GetColor(intRow) = False Then
               For i = 0 To grdDataList.Cols - 1
                  grdDataList.col = i
                  If i <= 6 Then
                     grdDataList.CellBackColor = QBColor(7)
                  Else
                     grdDataList.CellBackColor = QBColor(15)
                  End If
               Next i
            End If
         End If
      End If
   Next intRow
   Me.Show
'   bolShowMsgBox = False
'   Call doQuery
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   strOrderBy = "收文日 asc,本所案號 asc,總收文號 asc"
   If ConstrainCheck = True Then
      SetDataListWidth
      bolShowMsgBox = True
      Call doQuery
   End If
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
   m_blnColOrderAsc = True    '2010/9/28 ADD BY SONIA
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

Private Sub Form_Activate()
If isLoad = False Then
   MoveFormToCenter Me
   isLoad = True
End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolShowMsgBox = False
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   'Add By Sindy 2022/2/21
   '檢查當時是否需要為他人職代
   Combo3.Clear
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
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, , txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
   
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'
'   Select Case strUserNum
'      '外商陳經理可看全所CFT,FCT,S,CFC
'      Case "68005"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
''cancel by sonia 2014/6/9
''      '蔣律師可看中所全部
''      Case "79037"
''         txtSalesArea.Enabled = True
''         txtSalesArea1.Enabled = True
''         txtSales.Enabled = True
''         txtSalesArea = "S2"
''         txtSalesArea1 = "S29"
''end 2014/6/9
'      '小真,杜副總可看全部
'      'modify by sonia 2014/6/9 +美珍77027
'      'Modify by Amy 2015/02/03 拿掉美珍 改寫至特殊人員(總經理業務工作代理人員)
'      Case "65001", "68006"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         '副總預設所有智權人員
'         If strUserNum = "68006" Then
'            txtSalesArea = "S"
'            txtSalesArea1 = "S99"
'         End If
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'         txtSales.Enabled = True
'      '王協理可看專利處
'      Case "71011"
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'         txtSales.Enabled = True
'      'Added by Lydia 2020/07/02 柄佑可看中所全部但業務區仍預設自已部門
'      Case "82026"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = GetST15(strUserNum)
'         txtSalesArea1 = txtSalesArea
'      'end 2002/07/02
'      Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "08"
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'            '分所財務人員可看該所全部
''            Case "C1", "NM", "KM"
''               txtSalesArea.Enabled = True
''               txtSalesArea1.Enabled = True
''               txtSales.Enabled = True
'            '各區主管
'            Case "SM"
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               '原羅文旭72009可兼看中一區,94/7/1只可看S22
'               '2005/7/5林永生71003可看中所全部,但預設S23
'               If strUserNum = "71003" Then
'                  txtSalesArea = "S23"
'                  txtSalesArea1 = "S23"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'               '簡協理可看北所全部但預設S15
'               If strUserNum = "69005" Then
'                  txtSalesArea = "S15"
'                  txtSalesArea1 = "S15"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'               End If
'               End If
'               txtSales.Enabled = True
'            '加入外商主管  王宗珮、洪琬姿、葉易雲
'            Case "21", "26", "28"
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               txtSales.Enabled = True
'            '其他只能看自己
'            Case Else
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               'Added by Lydia 2017/07/25 多使用者權限,則增加部門範圍
'               strExc(1) = PUB_GetSalesList(strUserNum, , , , , strExc(2), strExc(3))
'               If strExc(3) <> "" And strExc(3) > txtSalesArea1 Then
'                  txtSalesArea1 = strExc(3)
'               End If
'               'end 2017/07/25
'         End Select
'   End Select
'
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'
'   'Add by Amy 2015/02/03 +總經理業務工作代理人員
'   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'       bolSpecMan = True
'       strSpecCode = "總經理業務工作代理人員"
'   'Modify  by Amy 2014/05/20 開放專利處部份智權同仁資料給彥葶代為處理
'   ElseIf CheckLevel(strUserNum, "A8") = True Then
'        bolSpecMan = True
'        strSpecCode = "A8"
'   End If
'
'   If bolSpecMan = True Then
'        'Add by Amy 2015/02/03 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
'            txtSalesArea.Enabled = True: txtSalesArea = ""
'            txtSalesArea1.Enabled = True: txtSalesArea1 = ""
'            txtSales.Enabled = True
'        End If
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True: txtSales = ""
'   Else
'        txtSales = strUserNum
'   End If
'   'end 2014/05/20

   SetDataListWidth
   
'   'Add By Sindy 2020/7/1 記錄原操作人可以查詢的業務區及所別
'   'txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2020/7/1 END
    
    PUB_AddExcuteLog Me.Name 'Added by Lydia 2021/01/11 登入記錄
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pub_CallNextForm = False
   Set frm210136 = Nothing
End Sub

'取得正確的 row & col
Public Sub getGrdColRow(ByRef oObj As MSHFlexGrid, ByVal x As Single, ByVal y As Single, ByRef col As Long, ByRef row As Long)
Dim nIndex As Integer
col = 0: row = 0
For nIndex = 0 To oObj.Rows - 1
    If y > oObj.RowHeight(nIndex) Then
        row = row + 1
        y = y - oObj.RowHeight(nIndex)
    ElseIf y > 0 Then
        row = row + 1
        Exit For
    End If
Next nIndex
For nIndex = 0 To oObj.Cols - 1
    If x > oObj.ColWidth(nIndex) Then
        col = col + 1
        x = x - oObj.ColWidth(nIndex)
    ElseIf x > 0 Then
        col = col + 1
        Exit For
    End If
Next nIndex
col = col - 1 + IIf(oObj.LeftCol <> oObj.FixedCols And oObj.LeftCol <> 0, oObj.LeftCol - oObj.FixedCols, 0)
row = row - 1 + IIf(oObj.TopRow <> oObj.FixedRows And oObj.TopRow <> 0, oObj.TopRow - oObj.FixedRows, 0)

If col > oObj.Cols - 1 Then col = oObj.Cols - 1
If row > oObj.Rows - 1 Then row = oObj.Rows - 1
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grdDataList, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   grdDataList.col = nCol
   grdDataList.row = nRow
   Me.Enabled = False
   If Me.grdDataList.row < 1 Then
      If Me.grdDataList.Text = "費用" Or Me.grdDataList.Text = "點數" Then
'         If m_blnColOrderAsc = True Then
'            strOrderBy = Me.grdDataList.Text & " asc,總收文號 asc" '昇冪
'            m_blnColOrderAsc = False
'         Else
'            strOrderBy = Me.grdDataList.Text & " desc,總收文號 desc" '降冪
'            m_blnColOrderAsc = True
'         End If
'         bolShowMsgBox = False
'         Call doQuery
         If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 3 '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdDataList.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdDataList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
   Me.Enabled = True
End Sub

Private Sub grdDataList_SelChange()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
   If grdDataList.Text = "V" Then
      grdDataList.Text = ""
      
      If GetColor(grdDataList.MouseRow) = False Then
         For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            If i <= 6 Then
               grdDataList.CellBackColor = QBColor(7)
            Else
               grdDataList.CellBackColor = QBColor(15)
            End If
         Next i
      End If
   Else
      grdDataList.Text = "V"
      For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
      Next i
   End If
End If
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 2
         txtRecvDate(0).SetFocus
   End Select
End Sub

'CANCEL BY SONIA 2014/6/26
'Private Sub txtCU1_LostFocus()
'   txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
'End Sub
'END 2014/6/26

Private Sub txtRecvDate_Click(Index As Integer)
   Option1(2).Value = True
End Sub

Private Sub txtRecvDate_GotFocus(Index As Integer)
   TextInverse txtRecvDate(Index)
   CloseIme
End Sub

Private Sub txtRecvDate_Validate(Index As Integer, Cancel As Boolean)
   If txtRecvDate(Index) <> "" Then
      If ChkDate(txtRecvDate(Index)) = False Then
         Cancel = True
         txtRecvDate(Index).SetFocus
         txtRecvDate_GotFocus Index
         Exit Sub
      End If
      If Index = 1 Then
         If RunNick2(txtRecvDate(0), txtRecvDate(1)) = True Then
            txtRecvDate(Index).SetFocus
            txtRecvDate_GotFocus Index
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtCU1_GotFocus()
   TextInverse txtCU1
   CloseIme
End Sub

Private Sub txtCU1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCU2_GotFocus()
   'MODIFY BY SONIA 2014/6/26
   'If Len(txtCU1) = 9 Then
   If txtCU1 <> "" Then
   'END 2014/6/26
      txtCU2 = Left(txtCU1, 6) & "ZZZ"
      txtCU2.SelStart = 6
      txtCU2.SelLength = 3
   End If
   TextInverse txtCU2
   CloseIme
End Sub

Private Sub txtCU2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales, True)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2022/2/21
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   If PUB_txtSales_Limit(txtSales, m_strListPer, , txtSalesArea, txtSalesArea1, _
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

Private Sub cmdok_Click(Index As Integer)
cmdState = Index
PubShowNextData
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim intRow As Integer
Dim strCP01 As String 'Add By Sindy 2012/5/21
   
   Select Case cmdState
   Case 1 '承辦進度
      Me.Enabled = False
      For intRow = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = intRow
         If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            If GetColor(intRow) = False Then
               For i = 0 To grdDataList.Cols - 1
                  grdDataList.col = i
                  If i <= 6 Then
                     grdDataList.CellBackColor = QBColor(7)
                  Else
                     grdDataList.CellBackColor = QBColor(15)
                  End If
               Next i
            End If
            grdDataList.col = 18 '總收文號
            If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               strCP01 = GetCaseProData(Trim(Pub_RplStr(grdDataList.Text)), "CP01")
               frm100101_K.CmdFormName = UCase(Me.Name)
               frm100101_K.Show
               frm100101_K.Process Pub_RplStr(grdDataList.Text)
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
         End If
      Next intRow
      Me.Enabled = True
   Case 3 '發E-Mail
      Me.Enabled = False
      For intRow = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = intRow
         If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            If GetColor(intRow) = False Then
               For i = 0 To grdDataList.Cols - 1
                  grdDataList.col = i
                  If i <= 6 Then
                     grdDataList.CellBackColor = QBColor(7)
                  Else
                     grdDataList.CellBackColor = QBColor(15)
                  End If
               Next i
            End If
            If Not IsNull(grdDataList.TextMatrix(intRow, 18)) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm210136_2.Show
               frm210136_2.Process Pub_RplStr(grdDataList.TextMatrix(intRow, 18)) '總收文號
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
         End If
      Next intRow
      Me.Enabled = True
   Case Else
   End Select
End Sub

Private Sub txtCuName_GotFocus()
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   cmdSearch.Default = False: cmdFind.Default = True
   
   TextInverse Me.txtCuName
   OpenIme
End Sub

'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
Private Sub txtCuName_LostFocus()
   cmdSearch.Default = True: cmdFind.Default = False
End Sub
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Private Sub cmdFind_Click()
   If Me.txtCuName.Text = "" Then
      MsgBox "請輸入客戶中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txtCuName.SetFocus
      txtCuName_GotFocus
      Exit Sub
   End If
   frm090801_1.m_strCustChnName = Me.txtCuName.Text
   frm090801_1.lblName.Caption = Me.txtCuName.Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   If m_blnOneRec = True And m_strCustCode <> "" Then
      Me.txtCU1.Text = m_strCustCode
      Me.txtCU2.Text = IIf(Right(m_strCustCode, 3) = "000", Mid(m_strCustCode, 1, 6) & "ZZZ", IIf(Right(m_strCustCode, 1) = "0", Mid(m_strCustCode, 1, 8) & "Z", m_strCustCode))
      Me.txtCuName.Text = GetCustomerName(m_strCustCode)
   End If
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   If Me.txtCU1.Text <> "" And Me.txtCU2.Text <> "" Then
      Call cmdSearch_Click
   End If
   
End Sub
