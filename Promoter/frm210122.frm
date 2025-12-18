VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210122 
   BorderStyle     =   1  '單線固定
   Caption         =   "應收帳款查詢"
   ClientHeight    =   6000
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9380
   Begin VB.CommandButton CmdNext 
      Caption         =   "下一筆"
      Height          =   400
      Left            =   6840
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   1695
      Left            =   60
      TabIndex        =   32
      Top             =   750
      Width           =   9285
      Begin VB.CommandButton cmdEdit 
         Caption         =   "預定收款日期輸入"
         Height          =   405
         Left            =   7440
         TabIndex        =   21
         Top             =   510
         Width           =   1725
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "搜尋(&Q)"
         Height          =   315
         Left            =   4680
         TabIndex        =   23
         Top             =   1373
         Width           =   765
      End
      Begin VB.TextBox txtCU2 
         Height          =   300
         Left            =   2250
         MaxLength       =   9
         TabIndex        =   20
         Top             =   1029
         Width           =   970
      End
      Begin VB.TextBox txtCU1 
         Height          =   300
         Left            =   1035
         MaxLength       =   9
         TabIndex        =   19
         Top             =   1029
         Width           =   970
      End
      Begin VB.TextBox systemkind 
         Height          =   300
         Left            =   4965
         TabIndex        =   18
         Text            =   "ALL"
         Top             =   681
         Width           =   2130
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   1035
         MaxLength       =   3
         TabIndex        =   14
         Top             =   681
         Width           =   405
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   15
         Top             =   681
         Width           =   675
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   2190
         MaxLength       =   1
         TabIndex        =   16
         Top             =   681
         Width           =   240
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   2460
         MaxLength       =   2
         TabIndex        =   17
         Top             =   681
         Width           =   315
      End
      Begin VB.TextBox txtKind 
         Height          =   285
         Left            =   1035
         MaxLength       =   1
         TabIndex        =   12
         Top             =   348
         Width           =   240
      End
      Begin VB.CheckBox Check1 
         Caption         =   "只顯示智權公司資料"
         Height          =   195
         Left            =   3930
         TabIndex        =   13
         Top             =   405
         Width           =   2445
      End
      Begin VB.TextBox txtbaseLine 
         Height          =   300
         Index           =   0
         Left            =   2100
         MaxLength       =   7
         TabIndex        =   5
         Top             =   0
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "總額下限"
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "總額"
         Height          =   300
         Index           =   1
         Left            =   3180
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.TextBox txtbaseLine 
         Height          =   300
         Index           =   1
         Left            =   4020
         MaxLength       =   7
         TabIndex        =   7
         Top             =   0
         Width           =   800
      End
      Begin VB.TextBox txtbaseLine 
         Height          =   300
         Index           =   2
         Left            =   5100
         MaxLength       =   7
         TabIndex        =   8
         Top             =   0
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "收據金額"
         Height          =   300
         Index           =   2
         Left            =   6180
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txtbaseLine 
         Height          =   300
         Index           =   3
         Left            =   7260
         MaxLength       =   7
         TabIndex        =   10
         Top             =   0
         Width           =   800
      End
      Begin VB.TextBox txtbaseLine 
         Height          =   285
         Index           =   4
         Left            =   8340
         MaxLength       =   7
         TabIndex        =   11
         Top             =   0
         Width           =   800
      End
      Begin VB.Label Label7 
         Caption         =   "註：列欄 N:未列印收據 Y:待列印收據                  Z.已開立INVOCIE"
         ForeColor       =   &H000000C0&
         Height          =   336
         Left            =   5664
         TabIndex        =   43
         Top             =   1248
         Width           =   3012
      End
      Begin MSForms.TextBox txtCuName 
         Height          =   300
         Left            =   1530
         TabIndex        =   22
         Top             =   1380
         Width           =   3015
         VariousPropertyBits=   671105051
         Size            =   "5318;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCuNam 
         Caption         =   "申請人中文名稱："
         Height          =   180
         Left            =   30
         TabIndex        =   39
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Line Line3 
         X1              =   2070
         X2              =   2220
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Left            =   0
         TabIndex        =   38
         Top             =   1089
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "系統類別："
         Height          =   180
         Left            =   3930
         TabIndex        =   37
         Top             =   741
         Width           =   900
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1170
         X2              =   2625
         Y1              =   870
         Y2              =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Left            =   0
         TabIndex        =   36
         Top             =   741
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "查詢資料："
         Height          =   180
         Left            =   0
         TabIndex        =   35
         Top             =   400
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(1.收據明細  2.客戶總額)"
         Height          =   180
         Left            =   1320
         TabIndex        =   34
         Top             =   405
         Width           =   1920
      End
      Begin VB.Line Line4 
         X1              =   4860
         X2              =   5010
         Y1              =   150
         Y2              =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "應收金額："
         Height          =   180
         Left            =   0
         TabIndex        =   33
         Top             =   60
         Width           =   900
      End
      Begin VB.Line Line5 
         X1              =   8130
         X2              =   8265
         Y1              =   150
         Y2              =   150
      End
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1095
      TabIndex        =   1
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7590
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8385
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtZone 
      Height          =   300
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2310
      TabIndex        =   2
      Top             =   0
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1095
      MaxLength       =   6
      TabIndex        =   3
      Top             =   360
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3435
      Left            =   60
      TabIndex        =   31
      Top             =   2490
      Width           =   9285
      _ExtentX        =   16387
      _ExtentY        =   6050
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.5
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
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2070
      TabIndex        =   42
      Top             =   420
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      Caption         =   "M51(看客戶編號)："
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   3870
      TabIndex        =   41
      Top             =   390
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblCuNo 
      Caption         =   "lblCuNo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5430
      TabIndex        =   40
      Top             =   390
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2310
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   5040
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   60
      TabIndex        =   29
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   3480
      TabIndex        =   28
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   60
      TabIndex        =   27
      Top             =   52
      Width           =   720
   End
End
Attribute VB_Name = "frm210122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/05/04 為了方便主管查看不同客戶，應收帳款查詢改為只呼叫畫面不限制帶入客戶代號
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'申請人改為”申請人名稱”，業務區改為”業務區；
'將條件”應收金額、查詢資料、本所案號、客戶編號、申請人中文名稱”改放在Frame1，
'將條件”應收金額、查詢資料、本所案號、客戶編號、申請人中文名稱”改放在Frame1，
'若從請款單明細表frm210146呼叫本程式則隱藏Frame1，帶入前一畫面的值，預設查詢條件為總額下限金額1元，查收據明細。
'end 2021/07/27
'Memo by Lydia 2021/07/14 Form2.0已修改(lblSalesName、txtCuName）；grdDataList改字型=新細明體-ExtB
'Memo by Lydia 2019/07/01 表單名稱:客戶應收帳款查詢=>應收帳款查詢
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改'
Option Explicit

Dim stST05 As String, stST15 As String, bolSelData As Boolean
Dim i As Integer, m_row As Integer
Public SetDate As String
Public SetData As String
Public SetKey As String
Dim BolCanEditDate As Boolean
'Add by Amy 2014/05/21
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim m_strListPer As String 'Add By Sindy 2020/7/28
'Added by Lydia 2021/07/27 智權-調整財務系統(20200909)：外部呼叫使用
Dim m_PrevForm As Form  '前一畫面
Dim m_SalesNo As String '前一畫面的智權人員編號
Dim m_NowKey As Integer, m_MaxKey As Integer  '目前的客戶編號索引和最大筆數
Dim m_ArrayKey As Variant  '傳入的客戶編號(可多筆，用,區隔)
'end 2021/07/27


'Added by Lydia 2021/07/27 外部呼叫使用
Public Sub SetParent(ByVal pForm As Form, ByVal pKeyNo As String, Optional ByVal pSalesNo As String)
   
   Set m_PrevForm = pForm
   m_ArrayKey = Empty
   m_NowKey = -1: m_MaxKey = -1
      
   If pKeyNo <> "" Then
        m_ArrayKey = Split(pKeyNo, ",")
        m_MaxKey = UBound(m_ArrayKey)
        If pSalesNo <> "" Then m_SalesNo = pSalesNo
   End If
End Sub
'end 2021/07/27

Private Sub cmdEdit_Click()
Dim bolIsEdit As Boolean
Dim tmpvar As Variant
Dim m_salesst05 As String  '2015/8/24 add by sonia

   bolIsEdit = False
   SetDate = ""
   If m_row <> 0 Then
       grdDataList.row = m_row
       grdDataList.col = 10
       If grdDataList.CellBackColor = &HFFC0C0 Then
           If grdDataList.TextMatrix(m_row, 10) = "" Then
               tmpvar = Split(grdDataList.TextMatrix(m_row, 11), " ")
               '2015/8/24 modify by sonia 簡協理可查北所全部,但跨區只可改區主管的資料,蘇特助開放可查分所資料,但跨區只可改北五,中一,南所,高所,四區主管的資料
               'If tmpvar(1) = strUserNum Or BolCanEditDate = True Then
                   'SetDate = ""
                   'bolIsEdit = True
               If tmpvar(1) = strUserNum Then
                  SetDate = ""
                  bolIsEdit = True
               ElseIf BolCanEditDate = True Then
                  m_salesst05 = PUB_GetST05(tmpvar(1))
                  Select Case strUserNum
                     Case "69005", "69010"
                        If Mid(tmpvar(0), 2) = stST15 Then
                           SetDate = ""
                           bolIsEdit = True
                        ElseIf m_salesst05 = "SM" Or (tmpvar(1) = "74018" And strUserNum = "69010") Then
                           SetDate = ""
                           bolIsEdit = True
                        Else
                           MsgBox "權限不足(跨區)！", vbCritical, "操作錯誤！"
                           Exit Sub
                        End If
                     Case Else
                        SetDate = ""
                        bolIsEdit = True
                  End Select
               '2015/8/24 end
               Else
                   '2009/6/4 MODIFY BY SONIA若為帶人主管權限時,也可以改預定收款日
                   'MsgBox "權限不足！", vbCritical, "操作錯誤！"
                   'Exit Sub
                   If Trim(txtSales) <> "" And stST05 = "SA" And txtSales.Enabled = True Then
                      'Modify By Sindy 2014/8/28
                      'If txtSales <> strUserNum And PUB_GetST52(txtSales) = strUserNum Then
                      If txtSales <> strUserNum And PUB_GetST52(txtSales, strUserNum) = True Then
                      '2014/8/28 END
                        SetDate = ""
                        bolIsEdit = True
                      Else
                        MsgBox "權限不足！", vbCritical, "操作錯誤！"
                        Exit Sub
                      End If
                   Else
                      MsgBox "權限不足！", vbCritical, "操作錯誤！"
                      Exit Sub
                   End If
                   '2009/6/4 END
                End If
           Else
               If BolCanEditDate = True Then
                  '2015/8/24 modify by sonia 修改智權部區主管權限
                  'SetDate = TAIWANDATE(grdDataList.TextMatrix(m_row, 10))
                  'bolIsEdit = True
                  tmpvar = Split(grdDataList.TextMatrix(m_row, 11), " ")
                  '區主管不可以改自己的資料日期,但由無日期到有日期可以
                  If tmpvar(1) = strUserNum And (PUB_GetST05(strUserNum) = "SM" Or strUserNum = "74018") Then
                     MsgBox "權限不足(自己的資料不能自己改日期)！", vbCritical, "操作錯誤！"
                     Exit Sub
                  End If
                  m_salesst05 = PUB_GetST05(tmpvar(1))
                  Select Case strUserNum
                     '簡協理可查北所全部,但跨區只可改區主管的資料,蘇特助開放可查分所資料,但跨區只可改北五,中一,南所,高所,四區主管的資料
                     Case "69005", "69010"
                        If Mid(tmpvar(0), 2) = stST15 Then
                           SetDate = TAIWANDATE(grdDataList.TextMatrix(m_row, 10))
                           bolIsEdit = True
                        ElseIf m_salesst05 = "SM" Or (tmpvar(1) = "74018" And strUserNum = "69010") Then
                           SetDate = TAIWANDATE(grdDataList.TextMatrix(m_row, 10))
                           bolIsEdit = True
                        Else
                           MsgBox "權限不足(跨區)！", vbCritical, "操作錯誤！"
                           Exit Sub
                        End If
                     '其他人可以改自己的
                     Case Else
                        SetDate = TAIWANDATE(grdDataList.TextMatrix(m_row, 10))
                        bolIsEdit = True
                  End Select
                  '2015/8/24 end
               Else
                   '2009/6/4 MODIFY BY SONIA若為帶人主管權限時,也可以改預定收款日
                   'MsgBox "權限不足！", vbCritical, "操作錯誤！"
                   'Exit Sub
                   If Trim(txtSales) <> "" And stST05 = "SA" And txtSales.Enabled = True Then
                      'Modify By Sindy 2014/8/28
                      'If txtSales <> strUserNum And PUB_GetST52(txtSales) = strUserNum Then
                      If txtSales <> strUserNum And PUB_GetST52(txtSales, strUserNum) = True Then
                      '2014/8/28 END
                        SetDate = TAIWANDATE(grdDataList.TextMatrix(m_row, 10))
                        bolIsEdit = True
                      Else
                        MsgBox "權限不足！", vbCritical, "操作錯誤！"
                        Exit Sub
                      End If
                   Else
                      MsgBox "權限不足！", vbCritical, "操作錯誤！"
                      Exit Sub
                   End If
                   '2009/6/4 END
               End If
           End If
       Else
           MsgBox "請先選取一筆資料！", vbCritical, "操作錯誤！"
           Exit Sub
       End If
   Else
       MsgBox "請先選取一筆資料！", vbCritical, "操作錯誤！"
       Exit Sub
   End If
   If bolIsEdit = True Then
       SetData = "申請人名稱：" & grdDataList.TextMatrix(m_row, 3) & vbCrLf & "本所案號：" & grdDataList.TextMatrix(m_row, 4) & vbCrLf & "收據日期：" & grdDataList.TextMatrix(m_row, 5) & vbCrLf & "收據號碼：" & grdDataList.TextMatrix(m_row, 6) & vbCrLf & "收據金額：" & grdDataList.TextMatrix(m_row, 8) & vbCrLf & "未收金額：" & grdDataList.TextMatrix(m_row, 9) & vbCrLf
       SetKey = grdDataList.TextMatrix(m_row, 6)
       Set frm210122_1.UpForm = Me
       frm210122_1.Show vbModal
       grdDataList.TextMatrix(m_row, 10) = ChangeTStringToTDateString(SetDate)
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'Added by Lydia 2021/07/27 外部呼叫使用：查詢下一筆
Private Sub cmdNext_Click()
    m_NowKey = m_NowKey + 1
    If m_NowKey > m_MaxKey Then
        MsgBox "已經是最後一筆！", vbInformation
        Call cmdExit_Click
        Exit Sub
    Else
        '帶入客戶編號
        If m_ArrayKey(m_NowKey) <> "" Then
            txtCU1 = Left(m_ArrayKey(m_NowKey) & String(8, "0"), 9)
            txtCU2 = Left(txtCU1, 6) & "ZZZ" '原先畫面輸入客戶編號起值不管幾碼，迄值第6~9碼自動帶出ZZZ
            lblCuNo.Caption = txtCU1 & "~" & txtCU2
            Call ProcQuery
        Else
            MsgBox "已經是最後一筆！", vbInformation
            Call cmdExit_Click
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
   'Memo by Lydia 2021/07/27 改成模組
   Call ProcQuery
End Sub

'Added by Lydia 2021/07/27 原本在cmdSearch_Click，改成模組
Private Sub ProcQuery()
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   bolSelData = False
   If ConstrainCheck = True Then
      grdDataList.Clear
      grdDataList.Rows = 2
      SetDataListWidth
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
      Call doQuery
      m_row = 0
   End If
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub
'end 2021/07/27

Private Sub Form_Load()
   MoveFormToCenter Me
   BolCanEditDate = CheckUse("frm210122", strEdit, False)
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
'   txtZone.Enabled = False
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'
'   Select Case strUserNum
'      '2008/6/10 CANCEL BY SONIA 取消蔣律師可看中所全部
'      'Case "79037"
'      '   txtZone = pub_strUserOffice
'      '   txtSalesArea.Enabled = True
'      '   txtSalesArea1.Enabled = True
'      '   txtSales.Enabled = True
'      '   txtSalesArea = "S2"
'      '   txtSalesArea1 = "S29"
'      '2008/6/10 END
'      '小真,杜副總可看全部
'      'modify by sonia 2014/6/9 +美珍77027
'      'Modify by Amy 2015/02/04 拿掉美珍 改寫至特殊人員(總經理業務工作代理人員)
'      Case "65001", "68006"
'         txtZone.Enabled = True
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
'         txtZone = pub_strUserOffice
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
'      '外商陳經理可看外商
'      Case "68005"
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'       'add by sonia 2016/12/21 柄佑可看中所全部但業務區仍預設自已業務區
'      Case "82026"
'         txtZone = pub_strUserOffice
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
'         txtSales = strUserNum
'      'end 2016/12/21
'     Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "08"
'               txtZone.Enabled = True
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'            '各區主管
'            Case "SM"
'               txtZone = pub_strUserOffice
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               '2015/8/21 modify by sonia 取消71003,加顏經理73009可看中所全部,蘇特助69010可看中南高全部
'               ''71003可看中所全部,但預設S23
'               'If strUserNum = "71003" Then
'               '   txtSalesArea = "S23"
'               '   txtSalesArea1 = "S23"
'               '   txtSalesArea.Locked = False
'               '   txtSalesArea1.Locked = False
'               '   txtSalesArea.Enabled = True
'               '   txtSalesArea1.Enabled = True
'               '73009可看中所全部,
'               If strUserNum = "73009" Then
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               '69010可看中南高全部
'               ElseIf strUserNum = "69010" Then
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               '2015/8/21
'               '簡協理可看北所全部但預設S15
'               ElseIf strUserNum = "69005" Then
'                  txtZone.Enabled = True 'Added by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所(預設S15)
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'               End If
'               txtSales.Enabled = True
'            '外商主管  王宗珮、洪琬姿、葉易雲
'            Case "21", "26", "28"
'               txtZone = pub_strUserOffice
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               txtSales.Enabled = True
'            '其他只能看自己
'            Case Else
'               txtZone = pub_strUserOffice
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               'Added by Lydia 2017/07/25 多使用者權限,則增加業務區範圍
'               strExc(1) = PUB_GetSalesList(strUserNum, , , , , strExc(2), strExc(3))
'               If strExc(3) <> "" And strExc(3) > txtSalesArea1 Then
'                  txtSalesArea1 = strExc(3)
'               End If
'               'end 2017/07/25
'         End Select
'   End Select
'
'   'Add By Sindy 2009/05/12
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'
'   'Add by Amy 2015/02/04 +總經理業務工作代理人員
'   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'       bolSpecMan = True
'       strSpecCode = "總經理業務工作代理人員"
'   'Modify  by Amy 2014/05/21 開放專利處部份智權同仁資料給彥葶代為處理
'   ElseIf CheckLevel(strUserNum, "A8") = True Then
'        bolSpecMan = True
'        strSpecCode = "A8"
'   End If
'
'   If bolSpecMan = True Then
'        'Add by Amy 2015/02/04 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
'            txtZone.Enabled = True
'            txtSalesArea.Enabled = True: txtSalesArea = ""
'            txtSalesArea1.Enabled = True: txtSalesArea1 = ""
'            txtSales.Enabled = True
'        End If
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True: txtSales = ""
'   Else
'         txtSales = strUserNum
'   End If
'   'end 2014/05/21
   
   'Modify By Sindy 2020/7/28 設定員編,業務區,所別權限
   'Modified by Lydia 2021/07/27 外部呼叫改用前一畫面的智權人員編號
   'Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
   'Modify By Sindy 2025/3/18 +Me.Name
   Call PUB_SetFormSaleDept(IIf(m_SalesNo <> "", m_SalesNo, strUserNum), txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode, , , , , , , , Me.Name)
   
   'Modify By Sindy 2025/3/17 mark
'   'Add By Sindy 2013/1/9 分所開放可使用此作業
'   'MODIFY BY SONIA 2015/6/1 分所出納改用業務區M71判斷,使用權限已在外層限制,否則85003不能看
'   'If stST05 = "KM" Or stST05 = "NM" Or stST05 = "C1" Then
'   If Pub_StrUserSt03 = "M71" Then
'      txtSalesArea.Enabled = True
'      txtSalesArea1.Enabled = True
'      txtSales.Enabled = True
'   End If
'   '2013/1/9 End
   
   SetDataListWidth
   bolSelData = False
   
   'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   Option1(0).Value = True
   Option1_Click 0
   
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   cmdEdit.Visible = False
   
   'Added by Lydia 2021/07/27  智權-調整財務系統(20200909)：外部呼叫使用
   'Mark by Lydia 2023/05/04 為了方便主管查看不同客戶，應收帳款查詢改為只呼叫畫面不限制帶入客戶代號
'   If TypeName(m_PrevForm) <> "Nothing" Then
'       cmdSearch.Visible = False: cmdSearch.Enabled = False
'       CmdNext.Top = cmdSearch.Top
'       CmdNext.Visible = True
'       If Pub_StrUserSt03 = "M51" Then
'            Label11.Visible = True
'            lblCuNo.Visible = True
'       End If
'       '拉高Grid
'       Frame1.Visible = False
'       grdDataList.Top = 750
'       grdDataList.Height = 4815
'       '預設查詢條件為總額下限金額1元，查收據明細
'       Option1(0).Value = 1
'       txtbaseLine(0) = "1"
'       txtKind = "1"
'       Call cmdNext_Click
'   Else
'       m_NowKey = -1: m_MaxKey = -1
'       Frame1.Visible = True
'       grdDataList.Top = 2670
'       grdDataList.Height = 2895
'   End If
'   'end 2021/07/27
   If TypeName(m_PrevForm) <> "Nothing" And m_MaxKey >= 0 Then '若前畫面有點選客戶編號，則帶入第一筆客戶
       '預設查詢條件為總額下限金額1元，查收據明細
       Option1(0).Value = 1
       txtbaseLine(0) = "1"
       txtKind = "1"
       Call cmdNext_Click
   End If
   'end 2023/05/04
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'Added by Lydia 2021/07/27 回前一畫面
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   'end 2021/07/27
   
   MenuEnabled
   Set frm210122 = Nothing
End Sub

Private Sub grdDataList_SelChange()
Dim m_mouseRow As Integer

   grdDataList.Visible = False
   m_mouseRow = grdDataList.MouseRow
   grdDataList.col = 0
   If m_mouseRow <> 0 And InStr(1, grdDataList.TextMatrix(m_mouseRow, 4), "小計：") = 0 Then
       If m_row <> 0 Then
           grdDataList.row = m_row
            For i = 0 To grdDataList.Cols - 1
                 grdDataList.col = i
                 If grdDataList.CellBackColor = &HFFC0C0 Then
                   grdDataList.CellBackColor = &H80000018
                 Else
                   grdDataList.CellBackColor = &HFFC0C0 '&H80000018 '&H8080FF
                 End If
           Next i
       End If
       If m_row <> m_mouseRow Then
           grdDataList.row = m_mouseRow
           m_row = m_mouseRow
            For i = 0 To grdDataList.Cols - 1
                 grdDataList.col = i
                 If grdDataList.CellBackColor = &HFFC0C0 Then
                   grdDataList.CellBackColor = &H80000018
                 Else
                   grdDataList.CellBackColor = &HFFC0C0
                 End If
           Next i
       Else
           m_row = 0
       End If
   End If
   grdDataList.Visible = True
   
End Sub

Private Sub Option1_Click(Index As Integer)
'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   txtbaseLine(0).Enabled = False
   txtbaseLine(1).Enabled = False
   txtbaseLine(2).Enabled = False
   'Modified by Lydia 2018/08/15 增加收據金額
   'For intI = 0 To 2
   txtKind.Text = ""
   txtKind.Enabled = True
   For intI = 0 To 4
   'end 2018/08/15
       txtbaseLine(intI).Enabled = False
       txtbaseLine(intI).Text = ""
   Next intI
   
If Index = 0 Then '總額下限
   txtbaseLine(0).Enabled = True
'Modified by Lydia 2018/08/15
'Else
ElseIf Index = 1 Then '總額區間
   txtbaseLine(1).Enabled = True
   txtbaseLine(2).Enabled = True
'Added by Lydia 2018/08/15 增加收據金額
ElseIf Index = 2 Then
   '選收據金額時，查詢資料選項自動改為１收據明細
   txtKind.Text = "1'"
   txtKind.Enabled = False
   txtbaseLine(3).Enabled = True
   txtbaseLine(4).Enabled = True
'end 2018/08/15
End If
End Sub

Private Sub systemkind_GotFocus()
   TextInverse systemkind
   CloseIme
End Sub

Private Sub systemkind_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub



Private Sub txtbaseLine_GotFocus(Index As Integer)

'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   TextInverse txtbaseLine(Index)
   CloseIme

End Sub

Private Sub txtbaseLine_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'CANCEL BY SONIA 2014/6/26
'Private Sub txtCU1_LostFocus()
'   txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
'End Sub
'END 2014/6/26

Private Sub txtkind_GotFocus()
   TextInverse txtKind
   CloseIme
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

Private Sub txtKind_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtKind_Validate(Cancel As Boolean)
   If txtKind <> "" Then
      Select Case txtKind
      Case "", "1", "2"
      Case Else
          MsgBox "查詢方式只可以輸入 1 到 2！", vbInformation, "輸入錯誤！"
          Cancel = True
      End Select
   End If
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
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   'Modify By Sindy 2025/3/17 + Me.Name
   If PUB_txtSales_Limit(txtSales, m_strListPer, txtZone, txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName, Me.Name) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2024/9/18 END
   
'   'Add By Sindy 2015/6/26 若有異動智權人員,需重新查詢業務區和所別
'   'modify by sonia 2016/6/15 加入帶人主管條件
'   'If txtSalesArea.Enabled = True Then 'Modify By Sindy 2016/5/5 + if
'   If txtSalesArea.Enabled = True Or PUB_GetST05Limits(strUserNum) = True Then
'      If txtSales.Text <> "" And txtSales.Text <> txtSales.Tag Then
'         txtZone = PUB_GetST06(Trim(txtSales))
'         txtSalesArea = PUB_GetStaffST15(Trim(txtSales), "1")
'         txtSalesArea1 = PUB_GetStaffST15(Trim(txtSales), "1")
'      End If
'   Else
'      'Add By Sindy 2016/5/6 還原(原操作人)可以查詢的業務區及所別
'      txtZone = txtZone.Tag
'      txtSalesArea = txtSalesArea.Tag
'      txtSalesArea1 = txtSalesArea1.Tag
'      '2016/5/6 END
'   End If
'
'   txtSales.Tag = txtSales.Text
'
'   'add by sonia 2016/12/21 取消智權人員編號時,無跨所別權限者重新預設所別
'   If Trim(txtSales) = "" And txtZone.Enabled = False Then
'      txtZone = pub_strUserOffice
'   End If
'   'end 2016/12/21
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

'Remove by Lydia 2017/07/24 改成共用模組,已不使用
'Function GetNotInOfficeAndFalseStaff(oStr As String, oStr2 As String) As String
'GetNotInOfficeAndFalseStaff = ""
'Dim rsTmp2 As New ADODB.Recordset
'Dim sqlTmp2 As String
'
'   sqlTmp2 = "select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='2' "
'   sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01<'6' "
'   Select Case strUserNum
'      Case "71011"  '王協理
'         sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01 in ('96031','96032') "
'
'      Case "67002" '葉經理
'         sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01 in ('96029','96030') "
'      Case Else
'   End Select
'   Set rsTmp2 = New ADODB.Recordset
'   With rsTmp2
'       .CursorLocation = adUseClient
'       .Open sqlTmp2, cnnConnection, adOpenStatic, adLockReadOnly
'       If .RecordCount <> 0 Then
'           .MoveFirst
'           Do While Not .EOF
'               GetNotInOfficeAndFalseStaff = GetNotInOfficeAndFalseStaff & "'" & CheckStr(.Fields(0)) & "',"
'               .MoveNext
'           Loop
'       End If
'   End With
'   Set rsTmp2 = Nothing
'End Function

Private Sub SetDataListWidth()
grdDataList.Visible = False
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, m_i As Integer
   Select Case txtKind
   Case "1", ""
        'Modify By Sindy 2014/5/21 +Grid中在收據號碼欄後加'列'欄, 顯示a0k32的值
        arrGridHeadText = Array("所別", "業務區", "智權人員", "申請人名稱" _
                  , "本所案號", "收據日期", "收據號碼", "列" _
                  , "收據金額", "未收金額", "預定收款日", "")
        'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
        'arrGridHeadWidth = Array(480, 680, 680, 2000 _
                           , 1200, 800, 900, 300 _
                           , 900, 900, 1100, 0)
        arrGridHeadWidth = Array(480, 680, 680, 2000 _
                           , 1200, 800, 900, 300 _
                           , 900, 900, 0, 0)
        grdDataList.Cols = UBound(arrGridHeadText) + 1
        grdDataList.MergeCells = flexMergeFree
        grdDataList.MergeCol(0) = True
        grdDataList.MergeCol(1) = True
        grdDataList.MergeCol(2) = True
        grdDataList.MergeCol(3) = True
        grdDataList.MergeCol(4) = False
        grdDataList.MergeCol(5) = False
        grdDataList.MergeCol(6) = False
        grdDataList.MergeCol(7) = False
        grdDataList.MergeCol(8) = False
        grdDataList.MergeCol(9) = False
        grdDataList.MergeCol(10) = False 'Add By Sindy 2014/5/21
        grdDataList.ColAlignment(8) = flexAlignRightCenter '7
        grdDataList.ColAlignment(9) = flexAlignRightCenter '8
        For iRow = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(iRow, 4) = "申請人小計：" Then
                grdDataList.MergeRow(iRow) = True
                grdDataList.row = iRow
                grdDataList.col = 4
                grdDataList.CellBackColor = &HFF80FF
            ElseIf grdDataList.TextMatrix(iRow, 3) = "智權人員小計：" Then
                grdDataList.MergeRow(iRow) = True
                grdDataList.row = iRow
                grdDataList.col = 3
                grdDataList.CellBackColor = &HFFFF80
            ElseIf grdDataList.TextMatrix(iRow, 2) = "業務區小計：" Then
                grdDataList.MergeRow(iRow) = True
                grdDataList.row = iRow
                grdDataList.col = 2
                grdDataList.CellBackColor = &H80FF80
            ElseIf grdDataList.TextMatrix(iRow, 1) = "所小計：" Then
                grdDataList.MergeRow(iRow) = True
                grdDataList.row = iRow
                grdDataList.col = 1
                grdDataList.CellBackColor = &HC0C0FF
            Else
                grdDataList.MergeRow(iRow) = False
            End If
        Next
   Case "2"
        arrGridHeadText = Array("所別", "業務區", "智權人員", "申請人名稱", "未收金額")
        arrGridHeadWidth = Array(480, 680, 680, 2000, 1000)
        grdDataList.Cols = UBound(arrGridHeadText) + 1
        grdDataList.MergeCells = flexMergeFree
        grdDataList.MergeCol(0) = True
        grdDataList.MergeCol(1) = True
        grdDataList.MergeCol(2) = True
        For iRow = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(iRow, 3) = "智權人員小計：" Then
                grdDataList.MergeRow(iRow) = True
                grdDataList.row = iRow
                grdDataList.col = 3
                grdDataList.CellBackColor = &HFFFF80
            ElseIf grdDataList.TextMatrix(iRow, 2) = "業務區小計：" Then
                grdDataList.MergeRow(iRow) = True
                grdDataList.row = iRow
                grdDataList.col = 2
                grdDataList.CellBackColor = &H80FF80
            ElseIf grdDataList.TextMatrix(iRow, 1) = "所小計：" Then
                grdDataList.MergeRow(iRow) = True
                grdDataList.row = iRow
                grdDataList.col = 1
                grdDataList.CellBackColor = &HC0C0FF
            Else
                grdDataList.MergeRow(iRow) = False
            End If
        Next
   End Select
   For iRow = 0 To grdDataList.Cols - 1
      grdDataList.row = 0
      grdDataList.col = iRow
      grdDataList.Text = arrGridHeadText(iRow)
      grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdDataList.CellAlignment = flexAlignLeftCenter
   Next
   grdDataList.Visible = True
End Sub

Public Function doQuery() As Boolean
'原先智權人員與業務區是抓  客戶檔的，但是因為客戶檔裡面的智權人員與業務區並沒有一併修正，還是掛著舊的、離職的智權人員，所以改抓 acc0k0 的業務區與智權人員
Dim stCon As String, stCon1 As String, stCon2 As String
Dim stConA0K As String, stConLOS As String, arrPer As Variant, ii As Integer, jj As Integer 'Add By Sindy 2024/2/16
Dim strCon As String, strCon1 As String, strCon2 As String, strCon3 As String, strCon4 As String
Dim stIdList As String, stConId As String '2010/5/11 ADD BY SONIA
   
   stCon1 = "": stCon2 = ""
   stConA0K = "": stConLOS = "" 'Add By Sindy 2024/2/16
   '蔣律師要控制所別
   'Add By Sindy 2013/1/9 分所開放可使用此作業,但只能查詢該所人員
   'If strUserNum = "79037" Then
   'modify by sonia 2014/6/9 取消79037
   'If strUserNum = "79037" Or stST05 = "KM" Or stST05 = "NM" Or stST05 = "C1" Then
   'MODIFY BY SONIA 2015/6/1 分所出納改用業務區M71判斷,使用權限已在外層限制,否則85003不能看
   'If stST05 = "KM" Or stST05 = "NM" Or stST05 = "C1" Then
   'Modify By Sindy 2025/3/17 mark;因輸入智權人員的欄位有控制此規則,所以查詢SQL不用再判斷
'   If Pub_StrUserSt03 = "M71" Then
'   '2013/1/9 End
'      stCon1 = stCon1 & " and st06 = '" & pub_strUserOffice & "'"
'   End If
   
   '區別
   'Modify By Sindy 2009/05/12
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   If (Trim(txtSales) <> "" And stST05 = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
  'Add by Amy 2014/05/21
  'Modify by Amy 2019/02/13 總經理業務工作代理人員
   ElseIf bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   'end 2014/05/21
   Else
      If txtSalesArea <> "" Then
         stCon1 = stCon1 & " and st15||'' >= '" & txtSalesArea & "'" 'Modify By Sindy 2021/8/4 a0k22 => st15
      End If
      If txtSalesArea1 <> "" Then
         stCon1 = stCon1 & " and st15||'' <= '" & txtSalesArea1 & "'" 'Modify By Sindy 2021/8/4 a0k22 => st15
      End If
      If txtSalesArea <> "" Or txtSalesArea1 <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1 & txtSalesArea & "-" & txtSalesArea1 'Add By Sindy 2010/12/23
      End If
   End If
   
   '智權人員
   stIdList = "" 'Add By Sindy 2024/2/16
   If txtSales <> "" Then
        '加入區主管若是輸入自己的編號時，要看見 自己的 + 離職智權人員 + 虛建智權人員的資料
        '加入69005之控制
'''edit by nickc 2008/04/25 改共用
'''        If (strUserNum = "74018" And txtSales = "74018") Or (strUserNum = "78007" And txtSales = "78007") Or (strUserNum = "71011" And txtSales = "71011") Or (strUserNum = "67002" And txtSales = "67002") Or (stST05 = "SM" And strUserNum <> "71003") Or (strUserNum = "71003" And txtSales = "71003" And txtSalesArea = "S23" And txtSalesArea1 = "S23") Or (stST05 = "SM" And strUserNum <> "69005") Or (strUserNum = "69005" And txtSales = "69005" And txtSalesArea = "S15" And txtSalesArea1 = "S15") Then
'''            stCon1 = stCon1 & " and a0k20||'' in (" & GetNotInOfficeAndFalseStaff(txtSalesArea, txtSalesArea1) & "'" & txtSales & "' ) "
'''        Else
'''            '查87027陳淑芳時同時查20001台中所
'''            If txtSales = "87027" Then
'''               stCon1 = stCon1 & " and a0k20||'' IN ('87027','20001') "
'''            Else
'''               stCon1 = stCon1 & " and a0k20||'' = '" & txtSales & "'"
'''            End If
'''        End If
        '2008/6/10 MODIFY BY SONIA 取消所別控制, 否則杜副總不能查分所
        'stCon1 = stCon1 & " and a0k20||'' in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice) & ") "
        '2010/5/11 MODIFY BY SONIA
        'stCon1 = stCon1 & " and a0k20||'' in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) & ") "
        'Modify by Amy 2014/05/21 +if
        'Modify by Amy 2019/02/12 總經理業務工作代理人員
        If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
            stIdList = PUB_GetSalesList(txtSales)
        Else
            stIdList = PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1)
        End If
        'end 2014/05/21
        
        If Pub_StrST52 Then
           stCon1 = "": stCon2 = ""
           stConA0K = "": stConLOS = "" 'Add By Sindy 2024/2/16
        End If
'        If InStr(stIdList, ",") = 0 Then
'           stCon1 = stCon1 & " and a0k20||'' = " & stIdList & " "
'        Else
'           stCon1 = stCon1 & " and a0k20||'' in (" & stIdList & " ) "
'        End If
        '2010/5/11 END
        pub_QL05 = pub_QL05 & ";" & Label4 & txtSales & lblSalesName 'Add By Sindy 2010/12/23
   'Modify by Amy 2014/05/21
   '智權人員 為空
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            'A2023彥葶登入,未輸智權人員-設定查A7人員
            stIdList = "'" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "'"
'            stCon1 = stCon1 & " and a0k20||'' in (" & stIdList & ") "
        End If
   'end 2014/05/21
   End If
   'Modify By Sindy 2024/2/16
   If stIdList <> "" Then
      If InStr(stIdList, ",") = 0 Then
         stConA0K = stConA0K & " and a0k20||'' = " & stIdList & " "
         stConLOS = stConLOS & " and instr(LOS04||''," & stIdList & ")>0 "
      Else
         stConA0K = stConA0K & " and a0k20||'' in (" & stIdList & " ) "
         arrPer = Split(stIdList, ",")
         stConLOS = " and ("
         For ii = 0 To UBound(arrPer)
            If ii > 0 Then stConLOS = stConLOS & " or "
            stConLOS = stConLOS & "instr(LOS04||''," & arrPer(ii) & ")>0"
         Next ii
         stConLOS = stConLOS & ")"
      End If
   End If
   '2024/2/16 END
   
   'Added by Lydia 2019/01/08 判斷收據金額(區間)
   If Option1(2).Value = True Then
      stCon1 = stCon1 & " and acase >= " & Val(txtbaseLine(3)) & " and acase <= " & Val(txtbaseLine(4))
   End If
   'end 2019/01/08
   
   If systemkind = "" Then
      systemkind = "ALL"
   End If
   pub_QL05 = pub_QL05 & ";" & Label9 & systemkind 'Add By Sindy 2010/12/23
   
   If Trim(txtCU1) <> "" Then
       txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
       txtCU2 = Mid(txtCU2 & "000000000", 1, 9)
       stCon1 = stCon1 & " and ((A0K03>='" & txtCU1 & "' and A0K03<='" & txtCU2 & "')) "
       pub_QL05 = pub_QL05 & ";" & Label8 & txtCU1 & "-" & txtCU2 'Add By Sindy 2010/12/23
   End If
   
   '2008/6/16 cancel by sonia 應收帳款總額下限並非未收金額下限
   'If txtbaseLine <> "" Then
      'stCon1 = stCon1 & " DIFF>=" & Val(txtbaseLine)
   'End If
   '2008/6/16 end
   
   If txt1(0) <> "" Then
        stCon1 = stCon1 & " and cp01='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
        stCon1 = stCon1 & " and cp02='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
        stCon1 = stCon1 & " and cp03='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
        stCon1 = stCon1 & " and cp04='" & txt1(3) & "' "
   End If
   If txt1(0) <> "" Or txt1(1) <> "" Or txt1(2) <> "" Or txt1(3) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label6 & txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/23
   End If
   stCon1 = stCon1 & " and cp01 in (" & GetAddStr(GetAllSysKind(systemkind)) & ") "
   'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
'     pub_QL05 = pub_QL05 & ";" & Label2 & txtbaseLine 'Add By Sindy 2010/12/23
   If Option1(0).Value = True Then
      'Modified by Lydia 2018/08/15
      'pub_QL05 = pub_QL05 & ";" & Label2 & txtbaseLine(0) 'Add By Sindy 2010/12/23
        pub_QL05 = pub_QL05 & ";" & Label2 & Option1(0).Caption & "(" & txtbaseLine(0) & ")"
   'Modified by Lydia 2018/08/15 總額區間
   'Else
   '  pub_QL05 = pub_QL05 & ";" & Label2 & txtbaseLine(1) & "-" & txtbaseLine(2)
   ElseIf Option1(1).Value = True Then
       pub_QL05 = pub_QL05 & ";" & Label2 & Option1(1).Caption & "(" & txtbaseLine(1) & "-" & txtbaseLine(2) & ")"
   'Added by Lydia 2018/08/15 +收據金額(區間)
   ElseIf Option1(2).Value = True Then
       pub_QL05 = pub_QL05 & ";" & Label2 & Option1(2).Caption & "(" & txtbaseLine(3) & "-" & txtbaseLine(4) & ")"
   'end 2018/08/15
   End If
   '.end  'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   
On Error GoTo ErrHnd

'2008/6/16 modify by sonia 應收帳款總額下限並非未收金額下限
'   strCon = "SELECT st06,a0k01,a0k02,a0k22,a0k20,ST02,A0K03,SUBSTR(CU04,1,10) as cu04,DIFF ,Acase ,pay,min(cp09) as ocp FROM ("
'   strCon = strCon & "  select a0k01,a0k02,A0K03,a0k20,a0k22,SUM(nvl(a0k06,0)+nvl(a0k07,0)-PAY) DIFF,SUM(nvl(a0k06,0)+nvl(a0k07,0)) Acase,sum(pay) pay FROM( "
'   strCon = strCon & " SELECT A0K01,a0k02,A0K03,a0k20,a0k22,nvl(A0K06,0) as a0k06,nvl(A0K07,0) as a0k07,sum(nvl(a1u04,0))+sum(nvl(a1u07,0))-sum(nvl(a1u08,0))+sum(nvl(a1u05,0))+sum(nvl(a1u09,0))-sum(nvl(a1u10,0)) PAY"
'   strCon = strCon & " from acc0k0,ACC1U0 WHERE (a0k09 is null or a0k09 = 0) AND A0K01=A1U02(+)"
'   strCon = strCon & " and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0)) "
'   strCon = strCon & " GROUP BY A0K01,a0k02,A0K03,a0k20,a0k22,nvl(A0K06,0),nvl(A0K07,0) ) AA"
'   strCon = strCon & " where (nvl(a0k06,0)+nvl(a0k07,0)) > PAY GROUP BY a0k01,a0k02,A0K03,a0k20,a0k22"
'   strCon = strCon & " ) NEW,CUSTOMER,STAFF,caseprogress "
'   strCon = strCon & " WHERE a0k01=cp60 and SUBSTR(A0K03,1,8)=CU01(+) AND SUBSTR(A0K03,9,1)=CU02(+) AND a0k20=ST01(+) " & stCon1
'   strCon = strCon & " group by st06,A0K01,a0k22,a0k20,ST02,a0k02,A0K03,SUBSTR(CU04,1,10),DIFF,Acase,pay "
   'Modify By Sindy 2014/5/21 +" & IIf(Check1.Value = 1, " and a0k11='J'", "") & "
   '                          +,a0k32
'Modify By Sindy 2024/2/16
'組2次語法, 第2次是要抓案源資料
For jj = 1 To IIf(stConLOS = "", 1, 2)
   strCon1 = "SELECT st06,a0k01,a0k32,a0k02,a0k22,a0k20,ST02,A0K03,NVL(SUBSTR(CU04,1,10),NVL(SUBSTR(cu05,1,10),SUBSTR(CU06,1,10))) as cu04,DIFF ,Acase ,pay,min(cp09) as ocp FROM ("
   strCon1 = strCon1 & " select a0k01,a0k32,a0k02,A0K03,a0k20,a0k22,SUM(nvl(a0k06,0)+nvl(a0k07,0)-PAY) DIFF,SUM(nvl(a0k06,0)+nvl(a0k07,0)) Acase,sum(pay) pay FROM( "
   strCon1 = strCon1 & " SELECT A0K01,a0k32,a0k02,A0K03,a0k20,a0k22,nvl(A0K06,0) as a0k06,nvl(A0K07,0) as a0k07,sum(nvl(a1u04,0))+sum(nvl(a1u07,0))-sum(nvl(a1u08,0))+sum(nvl(a1u05,0))+sum(nvl(a1u09,0))-sum(nvl(a1u10,0)) PAY"
   strCon1 = strCon1 & " from acc0k0,ACC1U0 WHERE (a0k09 is null or a0k09 = 0)" & IIf(Check1.Value = 1, " and a0k11='J'", "") & " AND A0K01=A1U02(+)"
   strCon1 = strCon1 & " and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0)) "
   strCon2 = " and a0k03 in (select a0k03 from ( "
   strCon3 = " GROUP BY A0K01,a0k32,a0k02,A0K03,a0k20,a0k22,nvl(A0K06,0),nvl(A0K07,0) ) AA"
   strCon3 = strCon3 & " where (nvl(a0k06,0)+nvl(a0k07,0)) > PAY GROUP BY a0k01,a0k32,a0k02,A0K03,a0k20,a0k22"
   'Modified by Morgan 2011/10/31 考慮拆收據情形加串0j0
   'strCon3 = strCon3 & " ) NEW,CUSTOMER,STAFF,caseprogress "
   'strCon3 = strCon3 & " WHERE a0k01=cp60 and SUBSTR(A0K03,1,8)=CU01(+) AND SUBSTR(A0K03,9,1)=CU02(+) AND a0k20=ST01(+) " & stCon1
   'Modify By Sindy 2024/2/16
   strCon3 = strCon3 & " ) NEW,CUSTOMER,STAFF,acc0j0,caseprogress" & IIf(jj = 1, "", ",lawofficesource") & " "
   strCon3 = strCon3 & " WHERE a0j13(+)=a0k01 and cp09(+)=a0j01 and SUBSTR(A0K03,1,8)=CU01(+) AND SUBSTR(A0K03,9,1)=CU02(+) AND a0k20=ST01(+) " _
                     & stCon1 & IIf(jj = 1, stConA0K, " and LOS15(+)=cp162 and cp162 is not null and a0j02<>'TT999999000'" & stConLOS)
   '2024/2/16 END
   'end 2011/10/31
   strCon3 = strCon3 & " group by st06,A0K01,a0k32,a0k22,a0k20,ST02,a0k02,A0K03,NVL(SUBSTR(CU04,1,10),NVL(SUBSTR(cu05,1,10),SUBSTR(CU06,1,10))),DIFF,Acase,pay "
  
  'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   ' strCon4 = " ) AA,acc090 where AA.a0k22=a0901(+) group by a0k03 having sum(AA.diff)>= " & Val(txtbaseLine) & " )"
   If Option1(0).Value = True Then
        strCon4 = " ) AA,acc090 where AA.a0k22=a0901(+) group by a0k03 having sum(AA.diff)>= " & Val(txtbaseLine(0)) & " )"
   'Modified by Lydia 2018/08/15
   'Else
   ElseIf Option1(1).Value = True Then
        strCon4 = " ) AA,acc090 where AA.a0k22=a0901(+) group by a0k03 having sum(AA.diff)>= " & Val(txtbaseLine(1)) & " and sum(AA.diff)<= " & Val(txtbaseLine(2)) & " )"
   'Added by Lydia 2018/08/15 +收據金額(區間)
   ElseIf Option1(2).Value = True Then
        'Modified by Lydia 2019/01/08 收據金額不應在此判斷
        'strCon4 = " ) AA,acc090 where AA.a0k22=a0901(+) group by a0k03 having sum(AA.diff)>= " & Val(txtbaseLine(3)) & " and sum(AA.diff)<= " & Val(txtbaseLine(4)) & " )"
        strCon4 = " ) AA,acc090 where AA.a0k22=a0901(+) group by a0k03) "
   End If
   '.end  'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   
   '組語法 strCon1 + strCon2 + strCon1 + strCon3 + strCon4 + strCon3
   If strCon = "" Then 'Add By Sindy 2024/2/16
      strCon = strCon1 & strCon2 & strCon1 & strCon3 & strCon4 & strCon3
   Else
      strCon = strCon & " union " & strCon1 & strCon2 & strCon1 & strCon3 & strCon4 & strCon3
   End If
   '2008/6/16 end
Next jj

   If txtKind = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label10 & "1.收據明細" 'Add By Sindy 2010/12/23
      '2008/6/16 add by sonia 應收帳款總額下限並非未收金額下限
'      stCon2 = " select a0k03 from (select a0k03,sum(AA.diff) totamt from ( " & strCon & " ) AA,acc090 where AA.a0k22=a0901(+) group by a0k03) where totamt>=" & Val(txtbaseLine)
'      strCon = " select * from ( " & strCon & " ) where a0k03 in ( " & stCon2 & " )"
      '2008/6/16 end
      'Modify By Sindy 2014/5/21 +列欄
      'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
                     'stCon = "select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,aa.st02 as 智權人員,aa.cu04 as 申請人,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,sqldatet(AA.a0k02) as 收據日期,AA.a0k01 as 收據號碼,AA.a0k32 as 列,TO_CHAR(AA.Acase,'999,999,999') 收據金額,TO_CHAR(AA.diff,'999,999,999') 未收金額,sqldatet(rd05) 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'1' as osort from (" & strCon & ") AA,acc090,(select * from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd01,rd02) in (select rd01,max(rd02) from ReceivablesDay group by rd01 ) group by rd01,rd02)) BB,caseprogress where AA.a0k22=a0901(+) and AA.ocp=cp09(+) and AA.ocp=BB.RD01(+) "
                     stCon = "select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,aa.st02 as 智權人員,aa.cu04 as 申請人,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,sqldatet(AA.a0k02) as 收據日期,AA.a0k01 as 收據號碼,AA.a0k32 as 列,TO_CHAR(AA.Acase,'999,999,999') 收據金額,TO_CHAR(AA.diff,'999,999,999') 未收金額,' ' 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'1' as osort from (" & strCon & ") AA,acc090,caseprogress where AA.a0k22=a0901(+) and AA.ocp=cp09(+)  "
      'end 2018/08/22
      stCon = stCon & " union select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,aa.st02 as 智權人員,aa.cu04 as 申請人,'申請人小計：','申請人小計：','申請人小計：','申請人小計：','申請人小計：',TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) group by AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他'),a0902,aa.st02,aa.cu04 "
      stCon = stCon & " union select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,aa.st02 as 智權人員,'智權人員小計：','智權人員小計：','智權人員小計：','智權人員小計：','智權人員小計：','智權人員小計：',TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) group by AA.st06||AA.a0k22||' '||AA.a0k20||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他'),a0902,aa.st02 "
      stCon = stCon & " union select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,'業務區小計：','業務區小計：','業務區小計：','業務區小計：','業務區小計：','業務區小計：','業務區小計：',TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||AA.a0k22||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) and aa.st06<>'3' and aa.st06<>'4' group by AA.st06||AA.a0k22||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他'),a0902 "
      stCon = stCon & " union select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,'所小計：','所小計：','所小計：','所小計：','所小計：','所小計：','所小計：','所小計：',TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) group by AA.st06||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') "
      stCon = stCon & " order by oSort "
   Else
      pub_QL05 = pub_QL05 & ";" & Label10 & "2.客戶總額" 'Add By Sindy 2010/12/23
      '2008/6/16 add by sonia 應收帳款總額下限並非未收金額下限
'      stCon2 = " select a0k03 from (select a0k03,sum(AA.diff) totamt from ( " & strCon & " ) AA,acc090 where AA.a0k22=a0901(+) group by a0k03) where totamt>=" & Val(txtbaseLine)
'      strCon = " select * from ( " & strCon & " ) where a0k03 in ( " & stCon2 & " )"
      '2008/6/16 end
                    stCon = " select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,aa.st02 as 智權人員,aa.cu04 as 申請人,TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) group by AA.st06||AA.a0k22||' '||AA.a0k20||' '||AA.cu04||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他'),a0902,aa.st02,aa.cu04 "
      stCon = stCon & " union select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,aa.st02 as 智權人員,'智權人員小計：',TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||AA.a0k22||' '||AA.a0k20||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) group by AA.st06||AA.a0k22||' '||AA.a0k20||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他'),a0902,aa.st02 "
      stCon = stCon & " union select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 業務區,'業務區小計：','業務區小計：',TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||AA.a0k22||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) and aa.st06<>'3' and aa.st06<>'4' group by AA.st06||AA.a0k22||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他'),a0902 "
      stCon = stCon & " union select decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,'所小計：','所小計：','所小計：',TO_CHAR(sum(AA.diff),'999,999,999') 未收金額,' ' 預定收款日,AA.st06||'Z' as osort from (" & strCon & ") AA,acc090 where AA.a0k22=a0901(+) group by AA.st06||'Z',decode(AA.st06,'1','北所','2','中所','3','南所','4','高所','其他') "
      stCon = stCon & " order by oSort "
   End If

   CheckOC3
   SetDataListWidth
   grdDataList.Rows = 2
   grdDataList.Clear
   SetDataListWidth
 
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open stCon, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/23
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'         If txtKind = "1" Then
'            cmdEdit.Enabled = True
'            '2015/8/21分所出納已控制不可改預定收款日,故加按鈕不可使用
'            If Pub_StrUserSt03 = "M71" Then
'               cmdEdit.Enabled = False
'            End If
'            '2015/8/21 END
'         Else
'            cmdEdit.Visible = False
'         End If
         'end 2018/08/22
         grdDataList.Visible = True
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/23
         MsgBox "無符合資料！", vbInformation
      End If
   End With
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   ConstrainCheck = True
   
   'Add By Sindy 2009/05/14
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      txtSales.SetFocus
      txtSales_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   
   If txtKind = "" Then
      MsgBox "請輸入查詢資料！", vbExclamation
      txtKind.SetFocus
      txtkind_GotFocus
      ConstrainCheck = False
      Exit Function
   Else
      bolCancel = False
      Call txtKind_Validate(bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
   End If
   
'   If txtbaseLine = "" Then
'      '2008/6/17 modify by sonia
'      'MsgBox "請輸入應收總額下限！", vbExclamation
'      'txtbaseLine.SetFocus
'      'txtbaseLine_GotFocus
'      'ConstrainCheck = False
'      'Exit Function
'      txtbaseLine = 0
'   End If
   
   'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   If Option1(0).Value = True And txtbaseLine(0) = "" Then
      txtbaseLine(0) = 0
   ElseIf Option1(1).Value = True Then
          For intI = 1 To 2
            If txtbaseLine(intI) = "" Then
                MsgBox "請輸入正確總額區間！", vbExclamation
                txtbaseLine(intI).SetFocus
                txtbaseLine_GotFocus intI
                ConstrainCheck = False
                Exit Function
            End If
          Next intI
          
          If Val(txtbaseLine(1)) > Val(txtbaseLine(2)) Then
            MsgBox "請輸入正確總額區間！", vbExclamation
            txtbaseLine(1).SetFocus
            txtbaseLine_GotFocus 1
            ConstrainCheck = False
            Exit Function
          End If
   'Added by Lydia 2018/08/15
   ElseIf Option1(2).Value = True Then
          For intI = 3 To 4
            If txtbaseLine(intI) = "" Then
                MsgBox "請輸入正確收據金額區間！", vbExclamation
                txtbaseLine(intI).SetFocus
                txtbaseLine_GotFocus intI
                ConstrainCheck = False
                Exit Function
            End If
          Next intI
          
          If Val(txtbaseLine(3)) > Val(txtbaseLine(4)) Then
            MsgBox "請輸入正確收據金額區間！", vbExclamation
            txtbaseLine(1).SetFocus
            txtbaseLine_GotFocus 1
            ConstrainCheck = False
            Exit Function
          End If
   'end 2018/08/15
   End If
   'end  'Add by Lydia 2014/10/28  金額條件改為"客戶應收帳款：⊙總額下限：(原)  ○總額：(新增區間)"
   
   'Modify By Sindy 2020/7/29 檢查業務區欄位
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
   
''2015/8/21 CANCEL BY SONIA
''   '林永生71003檢查業務區範圍
''   If strUserNum = "71003" Then
''      If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
''         MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''      If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
''         MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
''         txtSalesArea1.SetFocus
''         txtSalesArea1_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''      If txtSalesArea > txtSalesArea1 Then
''         MsgBox "業務區範圍條件錯誤！", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''   End If
''2015/8/21 END
'
''2015/8/21 ADD BY SONIA
'   '顏經理73009檢查業務區範圍
'   If strUserNum = "73009" Then
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
'   End If
'   '蘇特助69010檢查業務區範圍,可查分所所有人及簡協理69005
'   If strUserNum = "69010" Then
'      If txtSalesArea <> "S13" And (txtSalesArea < "S15" Or txtSalesArea > "S99") Then
'         MsgBox "不可查此業務區資料", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If txtSalesArea1 <> "S13" And (txtSalesArea1 < "S15" Or txtSalesArea1 > "S99") Then
'         MsgBox "不可查此業務區資料", vbExclamation
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If txtSalesArea = "S15" And txtSales <> "69005" Then
'         MsgBox "北五區只可查69005簡協理的資料！", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
''2015/8/21 END
'
'   '簡金泉69005檢查業務區範圍
''Removed by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
''   If strUserNum = "69005" Then
''      If txtSalesArea < "S1" Or txtSalesArea > "S19" Then
''         MsgBox "業務區起始條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''      If txtSalesArea1 < "S1" Or txtSalesArea1 > "S19" Then
''         MsgBox "業務區迄止條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea1.SetFocus
''         txtSalesArea1_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''   End If
''end 2019/12/30
'
'   'add by sonia 2016/12/21 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
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
'   'end 2016/12/21
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      ConstrainCheck = False
'      Exit Function
'   End If
'   If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") Then
'       If Trim(txtSales) = "" Then
'           MsgBox "智權人員不可以空白！", vbExclamation, "操作錯誤！"
'           txtSales.SetFocus
'           txtSales_GotFocus
'           ConstrainCheck = False
'           Exit Function
'       End If
'       If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(txtSales) Then
'           MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
'           txtSales.SetFocus
'           txtSales_GotFocus
'           ConstrainCheck = False
'           Exit Function
'       End If
'   End If
   
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
