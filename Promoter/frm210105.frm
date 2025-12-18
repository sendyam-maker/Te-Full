VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210105 
   BorderStyle     =   1  '單線固定
   Caption         =   "暫收款查詢"
   ClientHeight    =   5760
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9430
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋(&Q)"
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   510
      Width           =   765
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   6
      Top             =   870
      Width           =   915
   End
   Begin VB.TextBox txtX 
      Height          =   300
      Index           =   1
      Left            =   2940
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "X"
      Top             =   120
      Width           =   195
   End
   Begin VB.TextBox txtX 
      Height          =   300
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "X"
      Top             =   120
      Width           =   195
   End
   Begin VB.TextBox txtCustNo 
      Height          =   300
      Index           =   1
      Left            =   3150
      MaxLength       =   8
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtCustNo 
      Height          =   300
      Index           =   0
      Left            =   1755
      MaxLength       =   8
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7620
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4425
      Left            =   135
      TabIndex        =   9
      Top             =   1275
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   7796
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
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
   Begin MSForms.TextBox txtCuName 
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   510
      Width           =   3075
      VariousPropertyBits=   671105051
      Size            =   "5424;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2550
      TabIndex        =   14
      Top             =   900
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
      Caption         =   "本暫收款以關係企業代號母號為主，金額為關係企業總合計"
      Height          =   180
      Left            =   4500
      TabIndex        =   13
      Top             =   960
      Width           =   4680
   End
   Begin VB.Label lblCuNam 
      Caption         =   "申請人中文名稱："
      Height          =   180
      Left            =   105
      TabIndex        =   12
      Top             =   570
      Width           =   1485
   End
   Begin VB.Label Label4 
      Caption         =   "智權人員："
      Height          =   180
      Left            =   105
      TabIndex        =   11
      Top             =   930
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2745
      X2              =   2925
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label2 
      Caption         =   "客戶編號："
      Height          =   180
      Left            =   105
      TabIndex        =   10
      Top             =   165
      Width           =   900
   End
End
Attribute VB_Name = "frm210105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'原”客戶代號”改為”客戶編號”
'原”客戶名稱”改為”申請人名稱”
'原”客戶中文名稱”改為”申請人中文名稱”
'end 2021/07/27
'Memo by Lydia 2021/07/13 改成Form2.0 ; lblSalesName、txtCuName、grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'Added by Lydia 2015/06/18 +申請人中文名稱查詢
Public m_strCustCode As String
Public m_blnOneRec As Boolean
'Add by Sindy 2023/12/22
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'2023/12/22 END


'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 3: .Cols = 5: .FixedRows = 2: .FixedCols = 0
      End If
      .row = 0
      
      .col = 0: .ColWidth(.col) = 1400: .Text = "客戶編號"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 4700: .Text = "申請人名稱"
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 2: .ColWidth(.col) = 1500: .Text = "暫收餘額"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 3: .ColWidth(.col) = 1200: .Text = "智權人員"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 4: .ColWidth(.col) = 0
      
      .row = 1
      .col = 1: .Text = "合計："
      .CellAlignment = flexAlignRightCenter
      .col = 2: .Text = ""
      .CellAlignment = flexAlignRightCenter
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = &H90EE90
      Next
      .Refresh
      .Visible = True
   End With
End Sub

Private Function ConstrainCheck() As Boolean
   Dim bolCancel As Boolean
   
   ConstrainCheck = True
   If txtSales = "" And txtCustNo(0) = "" And txtCustNo(1) = "" Then
      MsgBox "客戶編號與智權人員不可同時空白！", vbExclamation
      ConstrainCheck = False
      txtCustNo_GotFocus 0
   ElseIf txtSales = "" Then
      If (txtCustNo(0) = "") Then
         MsgBox "客戶編號起不可空白！", vbExclamation
         ConstrainCheck = False
         txtCustNo_GotFocus 0
      ElseIf (txtCustNo(1) = "") Then
         MsgBox "客戶編號迄不可空白！", vbExclamation
         ConstrainCheck = False
         txtCustNo_GotFocus 1
      End If
   End If
   
   'Add By Sindy 2021/5/24
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      txtSales.SetFocus
      txtSales_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   
End Function

Private Function doQuery() As Boolean

   Dim stConCust As String, stConSales As String
   
   stConCust = "": stConSales = ""
   '客戶
   If txtCustNo(0) <> "" Then
      stConCust = stConCust & " and ax208 >= 'X" & txtCustNo(0) & "'"
   End If
   If txtCustNo(1) <> "" Then
      stConCust = stConCust & " and ax208 <= 'X" & txtCustNo(1) & "'"
   End If
   
   '智權人員
   If txtSales <> "" Then
      'Modify by Morgan 2005/12/20 部門不限定TOT
      '2014/1/21 MODIF BY SONIA 取消ax201||''='1'條件,將下方的ax205||'' = '2401' 移上來
      'stConSales = stConSales & " and ax208 in (select distinct ax208 from acc021 where  ax201||''='1'"
      'Modify Sindy 2022/6/17 74018和82026查出來的客戶X43988帶出來的智權人員為有不同: ax208 改 substr(ax208,1,6)
      stConSales = stConSales & " and substr(ax208,1,6) in (select distinct substr(ax208,1,6) from acc021 where  ax205||'' = '2401' "
      '2014/1/21 END
      'Modify by Morgan 2008/4/28 若區主管查自己資料時要函該區已離職及虛建編號
      'stConSales = stConSales & " and ax205||'' = '2401' and ax209 = '" & txtSales & "'" & stConCust & ")"
      '2010/5/13 modify by sonia考慮中所跨區帶人離職時,帶人主管要看到離職智權人員資料,故傳入操作人員部門
      'stConSales = stConSales & " and ax205||'' = '2401' and ax209 in (" & PUB_GetSalesList(txtSales) & ")" & stConCust & ")"
      Select Case strUserNum
         '蔣律師,杜副總,杜燕文,劉大愛,王協理,葉經理,小真,林永生,簡協理不限制
         'modify by sonia 2014/6/9 +美珍77027並取消蔣律師79037
         'modify by sonia 2016/2/24 +69008
         'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
         Case "68006", "74018", "79053", "71011", "67002", "69008", "65001", "77027", "71003", "69005"
JumpToDef: 'Added by Lydia 2023/04/24
            '2014/1/21 MODIF BY SONIA 將ax205||'' = '2401' 移上去
            'stConSales = stConSales & " and ax205||'' = '2401' and ax209 in (" & PUB_GetSalesList(txtSales) & ")" & stConCust & ")"
            stConSales = stConSales & " and ax209 in (" & PUB_GetSalesList(txtSales) & ")" & stConCust & ")"
         Case Else
            'Adde by Lydia 2023/04/24 修改王副總退休之相關控制; 5/1加入李柏翰99050
            If strUserNum = "99050" And strSrvDate(1) >= "20230501" Then
                GoTo JumpToDef
            End If
            'end 2023/04/24
            Select Case PUB_GetST05(strUserNum)
               '電腦中心,財務,總經理看全部
               '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
               Case "00", "01", "08"
                  '2014/1/21 MODIF BY SONIA 將ax205||'' = '2401' 移上去
                  'stConSales = stConSales & " and ax205||'' = '2401' and ax209 in (" & PUB_GetSalesList(txtSales) & ")" & stConCust & ")"
                  stConSales = stConSales & " and ax209 in (" & PUB_GetSalesList(txtSales) & ")" & stConCust & ")"
               Case Else
                  '2014/1/21 MODIF BY SONIA 將ax205||'' = '2401' 移上去
                  'stConSales = stConSales & " and ax205||'' = '2401' and ax209 in (" & PUB_GetSalesList(txtSales, PUB_GetStaffST15(txtSales, 1), PUB_GetStaffST15(txtSales, 1)) & ")" & stConCust & ")"
                  stConSales = stConSales & " and ax209 in (" & PUB_GetSalesList(txtSales, PUB_GetStaffST15(txtSales, 1), PUB_GetStaffST15(txtSales, 1)) & ")" & stConCust & ")"
            End Select
      End Select
      '2010/5/13 END
   End If
   
On Error GoTo ErrHnd
   
   'Modify by Morgan 2005/12/20 部門不限定TOT,客戶抓母號加總
   'Modify by Morgan 2006/1/10 不同業務也要合併,智權人員抓最大傳票日的
   'Modified by Morgan 2013/5/8 修正百年問題
   '2014/1/21 MODIF BY SONIA 取消ax201||''='1'條件
   strSql = "select ax208, cu04, X0, st02, st01" & _
      " from ( select substr(ax208,1,6) ax208, substr(max(substr('0'||a0205,-7)||ax209),8) ax209,nvl(sum(ax207),0)-nvl(sum(ax206),0) X0 from acc021, acc020" & _
      " where ax205||'' = '2401'" & stConCust & stConSales & _
      " and a0201(+)=ax201 and a0202(+)=ax202 group by substr(ax208,1,6)) X, staff,customer " & _
      " where st01(+)=ax209 and cu01(+)=ax208||'00' and cu02(+)='0' and X0<>0" & _
      " order by ax208 asc, st01 asc"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         Call SetDataListWidth(True)
         Call Calculate
      Else
         MsgBox "無符合資料！", vbInformation
      End If
   End With
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If ConstrainCheck = True Then
      Call SetDataListWidth
      Call doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   Call SetDataListWidth
   
   'Modify By Sindy 2021/5/21 設定員編權限
   'Modify By Sindy 2025/3/18 +Me.Name
   Call PUB_SetFormSaleDept(strUserNum, , , , txtSales, , , , , , , , , , Me.Name)
   
   'Modify By Sindy 2025/3/17 mark
'   'Add By Sindy 2021/6/2 客戶應收帳款查詢frm210122、暫收款查詢frm210105：開放分所管理人員(部門M71)可查詢該所人員資料
'   If Pub_StrUserSt03 = "M71" Then
'      txtSales.Enabled = True
'   End If
'   '2021/6/2 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MenuEnabled
   Set frm210105 = Nothing
End Sub

Private Sub txtCustNo_GotFocus(Index As Integer)
   TextInverse txtCustNo(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtCustNo(Index).IMEMode = 2
   CloseIme
   If Index = 1 And Len(txtCustNo(0)) = 8 Then
      txtCustNo(Index) = Left(txtCustNo(0), 5) & "ZZZ"
      txtCustNo(Index).SelStart = 5
      txtCustNo(Index).SelLength = 3
   End If
End Sub

Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And (KeyAscii < Asc(0) Or KeyAscii > Asc(9)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCustNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 And Len(txtCustNo(Index)) > 4 Then
      txtCustNo(Index) = Left(txtCustNo(Index) & "000", 8)
   End If
End Sub

Private Sub Calculate()
   Dim ii As Integer, dblSum As Double
   With grdDataList
      .Visible = False
      For ii = 2 To .Rows - 1
         dblSum = dblSum + Val(.TextMatrix(ii, 2))
         .TextMatrix(ii, 2) = Format(.TextMatrix(ii, 2), "###,###,###.00")
      Next ii
      .TextMatrix(1, 2) = Format(dblSum, "###,###,###.00")
      .Visible = True
   End With
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = StaffQuery(txtSales)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSales.IMEMode = 2
   CloseIme
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_LostFocus()
   If Trim(txtSales) = "" Then
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
Dim arrPer As Variant, idx As Integer 'Add By Sindy 2016/5/4
Dim strSalesArea As String, strSalesArea1 As String 'Add By Sindy 2020/6/11
   
   'Modify By Sindy 2025/3/17
   'Modify By Sindy 2025/3/17 +Me.Name
   If PUB_txtSales_Limit(txtSales, "", , , , _
                         bolSpecMan, strSpecCode, lblSalesName, Me.Name) = False Then
      If txtSales.Visible = True Then '排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If
      Cancel = True
      Exit Sub
   End If
   '2025/3/17 END
End Sub

'Added by Lydia 2015/06/18 +申請人中文名稱查詢
Private Sub txtCuName_GotFocus()
   'Added by Lydia 2015/06/25 輸入申請人名稱按Enter自動執行搜尋後,執行查詢
   cmdSearch.Default = False: cmdFind.Default = True
   
   TextInverse Me.txtCuName
   OpenIme
End Sub

'Added by Lydia 2015/06/25 輸入申請人名稱按Enter自動執行搜尋後,執行查詢
Private Sub txtCuName_LostFocus()
   cmdSearch.Default = True: cmdFind.Default = False
End Sub
'Added by Lydia 2015/06/18 +申請人中文名稱查詢
Private Sub cmdFind_Click()
   If Me.txtCuName.Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
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
      Me.txtX(0).Text = Mid(m_strCustCode, 1, 1)
      Me.txtCustNo(0).Text = Mid(m_strCustCode, 2, Len(m_strCustCode) - 1)
      Me.txtX(1).Text = Mid(m_strCustCode, 1, 1)
      Me.txtCustNo(1).Text = IIf(Right(m_strCustCode, 3) = "000", Mid(m_strCustCode, 2, 5) & "ZZZ", IIf(Right(m_strCustCode, 1) = "0", Mid(m_strCustCode, 2, 7) & "Z", Mid(m_strCustCode, 2, Len(m_strCustCode) - 1)))
      Me.txtCuName.Text = GetCustomerName(m_strCustCode)
   End If
   'Added by Lydia 2015/06/25 輸入申請人名稱按Enter自動執行搜尋後,執行查詢
   If Me.txtCustNo(0).Text <> "" And Me.txtCustNo(1).Text <> "" Then
      Call cmdSearch_Click
   End If
End Sub
